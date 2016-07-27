'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-30 08:05:08+00 $
'$Log: frmStatusSelect.vb,v $
'Revision 1.0  2005-11-30 08:05:08+00  taylor
'Commented and moved to medscreen common gui
'
'Revision 1.2  2005-06-16 17:37:36+01  taylor
'Milestone Check in
'
'Revision 1.1  2004-05-04 10:19:58+01  taylor
'Header added to form
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'
Public Class frmStatusSelect
    Inherits System.Windows.Forms.Form

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible modes for form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Enum FormType
        '''<summary>Status</summary>
        Status
        '''<summary>Calendar</summary>
        Calendar
        '''<summary>Month Calendar</summary>
        MonthCalendar
    End Enum

    Private myFormType As FormType

#Region " Windows Form Designer generated code "

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Construct the form
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.CheckedListBox1.Items.Clear()
        Dim objPhrase As MedscreenLib.Glossary.Phrase
        Dim strPhrase As String
        For Each objPhrase In MedscreenLib.Glossary.Glossary.CollectionStatusList
            If objPhrase.PhraseID.Trim.Length > 0 Then  'skip blank ones added to list 
                Me.CheckedListBox1.Items.Add(objPhrase.PhraseID & " - " & objPhrase.PhraseText, False)
                'End If
            End If
        Next

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
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents CheckedListBox1 As System.Windows.Forms.CheckedListBox
    Friend WithEvents MonthCalendar1 As System.Windows.Forms.MonthCalendar
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmStatusSelect))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.CheckedListBox1 = New System.Windows.Forms.CheckedListBox()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 221)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(292, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(204, 8)
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
        Me.cmdOk.Location = New System.Drawing.Point(118, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.Items.AddRange(New Object() {"V Created", "C Confirmed", "W Assigned (Waiting to be sent)", "P Sent", "A Approved", "X Cancelled", "T Temporary", "  None "})
        Me.ComboBox1.Location = New System.Drawing.Point(64, 32)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(121, 21)
        Me.ComboBox1.TabIndex = 1
        Me.ComboBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(64, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Status"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "dd-MM-yyyy hh:mm"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(64, 32)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.ShowUpDown = True
        Me.DateTimePicker1.TabIndex = 3
        Me.DateTimePicker1.Visible = False
        '
        'CheckedListBox1
        '
        Me.CheckedListBox1.Items.AddRange(New Object() {"V Created", "C Confirmed", "W Assigned (Waiting to be sent)", "P Sent", "A Approved", "X Cancelled", "T Temporary"})
        Me.CheckedListBox1.Location = New System.Drawing.Point(64, 32)
        Me.CheckedListBox1.Name = "CheckedListBox1"
        Me.CheckedListBox1.Size = New System.Drawing.Size(200, 154)
        Me.CheckedListBox1.TabIndex = 4
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Location = New System.Drawing.Point(64, 32)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 5
        Me.MonthCalendar1.Visible = False
        '
        'frmStatusSelect
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(292, 261)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.MonthCalendar1, Me.Label1, Me.ComboBox1, Me.Panel1, Me.DateTimePicker1, Me.CheckedListBox1})
        Me.Name = "frmStatusSelect"
        Me.Text = "Select Status"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private strStatus As String

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        strStatus = Me.ComboBox1.Text
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Supplied status 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Status() As String
        Get
            Return strStatus
        End Get
        Set(ByVal Value As String)
            strStatus = Value
            Me.CheckedListBox1.Items.Clear()
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            Dim strPhrase As String
            For Each objPhrase In MedscreenLib.Glossary.Glossary.CollectionStatusList
                If objPhrase.PhraseID.Trim.Length > 0 Then  'skip blank ones added to list 
                    If InStr(strStatus, objPhrase.PhraseID) <> 0 Then
                        Me.CheckedListBox1.Items.Add(objPhrase.PhraseID & " - " & objPhrase.PhraseText, True)
                    Else
                        Me.CheckedListBox1.Items.Add(objPhrase.PhraseID & " - " & objPhrase.PhraseText, False)
                    End If
                End If
            Next
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Date (when in calendar mode)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SelectedDate() As Date
        Get
            If Me.FormTypeIs = FormType.Calendar Then
                Return Me.DateTimePicker1.Value
            Else
                Return Me.MonthCalendar1.SelectionStart
            End If
        End Get
        Set(ByVal Value As Date)
            Me.DateTimePicker1.Value = Value
            Me.MonthCalendar1.SetDate(Value)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set Form Type 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormTypeIs() As FormType
        Get
            Return myFormType
        End Get
        Set(ByVal Value As FormType)
            myFormType = Value
            Select Case myFormType
                Case FormType.Status
                    Me.Label1.Text = "Status"
                    Me.DateTimePicker1.Hide()
                    Me.ComboBox1.Hide()
                    Me.CheckedListBox1.Show()
                    Me.MonthCalendar1.Hide()
                Case FormType.Calendar
                    Me.Label1.Text = "Date"
                    Me.DateTimePicker1.Show()
                    Me.ComboBox1.Hide()
                    Me.CheckedListBox1.Hide()
                    Me.MonthCalendar1.Hide()
                Case FormType.MonthCalendar
                    Me.Label1.Text = "Date"
                    Me.DateTimePicker1.Hide()
                    Me.ComboBox1.Hide()
                    Me.CheckedListBox1.Hide()
                    Me.MonthCalendar1.Show()
            End Select
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get status type 
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

        Me.strStatus = "("
        If CheckedListBox1.CheckedItems.Count <> 0 Then
            ' If so, loop through all checked items and print results.
            Dim x As Integer
            Dim s As String = ""
            For x = 0 To CheckedListBox1.CheckedItems.Count - 1
                s = CheckedListBox1.CheckedItems(x).ToString.Chars(0)
                strStatus += "'" & s & "',"
            Next x
            strStatus = Mid(strStatus, 1, strStatus.Length - 1) & ")"
            'MessageBox.Show(strStatus)
        End If

    End Sub

End Class
