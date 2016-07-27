'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-27 12:55:43+00 $
'$Log: frmCallOutCMNO.vb,v $
'Revision 1.0  2005-11-27 12:55:43+00  taylor
'Checked in after commentating
'
Imports Intranet
Imports Intranet.intranet.Support
''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCallOutCMNO
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to get CM Number
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCallOutCMNO
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

        Me.DateTimePicker1.Value = Now

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
    Friend WithEvents cmdSystemAssign As System.Windows.Forms.Button
    Friend WithEvents cmdUserDefine As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents txtCMNumber As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCallOutCMNO))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdUserDefine = New System.Windows.Forms.Button()
        Me.cmdSystemAssign = New System.Windows.Forms.Button()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtCMNumber = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.cmdCancel, Me.cmdUserDefine, Me.cmdSystemAssign})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel1.Location = New System.Drawing.Point(272, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(152, 273)
        Me.Panel1.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Bitmap)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(12, 152)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(131, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "OK"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Button1.Visible = False
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(14, 192)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(131, 23)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdUserDefine
        '
        Me.cmdUserDefine.DialogResult = System.Windows.Forms.DialogResult.No
        Me.cmdUserDefine.Image = CType(resources.GetObject("cmdUserDefine.Image"), System.Drawing.Bitmap)
        Me.cmdUserDefine.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdUserDefine.Location = New System.Drawing.Point(14, 96)
        Me.cmdUserDefine.Name = "cmdUserDefine"
        Me.cmdUserDefine.Size = New System.Drawing.Size(131, 23)
        Me.cmdUserDefine.TabIndex = 1
        Me.cmdUserDefine.Text = "Manually Assigned"
        Me.cmdUserDefine.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdUserDefine, "User Define CM Number")
        '
        'cmdSystemAssign
        '
        Me.cmdSystemAssign.DialogResult = System.Windows.Forms.DialogResult.Yes
        Me.cmdSystemAssign.Image = CType(resources.GetObject("cmdSystemAssign.Image"), System.Drawing.Bitmap)
        Me.cmdSystemAssign.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSystemAssign.Location = New System.Drawing.Point(14, 32)
        Me.cmdSystemAssign.Name = "cmdSystemAssign"
        Me.cmdSystemAssign.Size = New System.Drawing.Size(131, 23)
        Me.cmdSystemAssign.TabIndex = 2
        Me.cmdSystemAssign.Text = "System Assign"
        Me.cmdSystemAssign.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSystemAssign, "Assign System CM Number")
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CustomFormat = "ddMMyyHHmm"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(144, 120)
        Me.DateTimePicker1.MinDate = New Date(2004, 1, 1, 0, 0, 0, 0)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(104, 20)
        Me.DateTimePicker1.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.DateTimePicker1, "Custom CM Number (Year Month Day Hours Minutes)")
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(128, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Custom CM Number"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label2.Location = New System.Drawing.Point(24, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(216, 80)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Call out Cm numbers may either be assigned by the system or may be user defined i" & _
        "n the format yymmddhhmi (Year Month Day Hours Minutes)"
        '
        'txtCMNumber
        '
        Me.txtCMNumber.Location = New System.Drawing.Point(144, 120)
        Me.txtCMNumber.Name = "txtCMNumber"
        Me.txtCMNumber.TabIndex = 0
        Me.txtCMNumber.Text = ""
        '
        'frmCallOutCMNO
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(424, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtCMNumber, Me.Label2, Me.Label1, Me.DateTimePicker1, Me.Panel1})
        Me.Name = "frmCallOutCMNO"
        Me.Text = "Enter CM Number"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private blnDefine As Boolean = True

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Returns CM number given
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property UserCMNo() As String
        Get
            If blnDefine Then
                Return Me.DateTimePicker1.Value.ToString("ddMMyyHHmm")
            Else
                Return Me.txtCMNumber.Text
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Says we are going to define CM number
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property DefineCMNumber() As Boolean
        Set(ByVal Value As Boolean)
            blnDefine = Value
            If blnDefine Then
                Me.Button1.Hide()
                Me.txtCMNumber.Hide()
                Me.DateTimePicker1.Show()
                Me.cmdSystemAssign.Show()
                Me.cmdUserDefine.Show()
            Else
                Me.Button1.Show()
                Me.txtCMNumber.Show()
                Me.DateTimePicker1.Hide()
                Me.cmdSystemAssign.Hide()
                Me.cmdUserDefine.Hide()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' User define option selected
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdUserDefine_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUserDefine.Click
        If cSupport.iCollection(False).CollectionExists(Me.DateTimePicker1.Value.ToString("yyMMddHHmm")) Then
            MsgBox("this id has already been used!", MsgBoxStyle.Exclamation Or MsgBoxStyle.OKOnly)
            Me.DialogResult = DialogResult.Cancel
        End If


    End Sub
End Class
