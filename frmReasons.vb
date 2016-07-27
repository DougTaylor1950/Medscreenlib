'$Revision$
'$Author$
'$Date$
'$Log$
'Revision 1.2  2005-06-17 07:05:20+01  taylor
'Milestone check in
'
'Revision 1.1  2004-05-06 14:15:47+01  taylor
'<>
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
Imports System.Windows.Forms
Public Class frmCancelReason
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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents DataGrid1 As System.Windows.Forms.DataGrid
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents XmlSchema11 As XMLSchema1
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DataGrid3 As System.Windows.Forms.DataGrid
    Friend WithEvents DataGridTableStyle1 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn1 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridBoolColumn1 As System.Windows.Forms.DataGridBoolColumn
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents DataGridTableStyle3 As System.Windows.Forms.DataGridTableStyle
    Friend WithEvents DataGridTextBoxColumn3 As System.Windows.Forms.DataGridTextBoxColumn
    Friend WithEvents DataGridBoolColumn3 As System.Windows.Forms.DataGridBoolColumn
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCancelReason))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.DataGrid1 = New System.Windows.Forms.DataGrid()
        Me.XmlSchema11 = New XMLSchema1()
        Me.DataGridTableStyle3 = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn3 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridBoolColumn3 = New System.Windows.Forms.DataGridBoolColumn()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.DataGrid3 = New System.Windows.Forms.DataGrid()
        Me.DataGridTableStyle1 = New System.Windows.Forms.DataGridTableStyle()
        Me.DataGridTextBoxColumn1 = New System.Windows.Forms.DataGridTextBoxColumn()
        Me.DataGridBoolColumn1 = New System.Windows.Forms.DataGridBoolColumn()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.txtComments = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.XmlSchema11, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel4.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(68, 379)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "TextBox1"
        Me.TextBox1.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(17, 377)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 23)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Price"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label1.Visible = False
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(184, 379)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 23)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Expiry"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label2.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(240, 379)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = "TextBox2"
        Me.TextBox2.Visible = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(352, 379)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Currency"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label3.Visible = False
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(408, 379)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.TabIndex = 5
        Me.TextBox3.Text = "TextBox3"
        Me.TextBox3.Visible = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(0, 344)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 23)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Item Code"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label4.Visible = False
        '
        'TextBox4
        '
        Me.TextBox4.Location = New System.Drawing.Point(65, 344)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.TabIndex = 7
        Me.TextBox4.Text = "TextBox4"
        Me.TextBox4.Visible = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(183, 344)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 23)
        Me.Label5.TabIndex = 10
        Me.Label5.Text = "Desc"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.Label5.Visible = False
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(239, 344)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(217, 20)
        Me.TextBox5.TabIndex = 9
        Me.TextBox5.Text = "TextBox5"
        Me.TextBox5.Visible = False
        '
        'DataGrid1
        '
        Me.DataGrid1.AlternatingBackColor = System.Drawing.Color.PeachPuff
        Me.DataGrid1.BackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid1.BackgroundColor = System.Drawing.Color.Silver
        Me.DataGrid1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGrid1.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.DataGrid1.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.CaptionText = "Chargeable Problems"
        Me.DataGrid1.DataMember = ""
        Me.DataGrid1.DataSource = Me.XmlSchema11
        Me.DataGrid1.Dock = System.Windows.Forms.DockStyle.Top
        Me.DataGrid1.FlatMode = True
        Me.DataGrid1.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.DataGrid1.ForeColor = System.Drawing.Color.Black
        Me.DataGrid1.GridLineColor = System.Drawing.Color.DimGray
        Me.DataGrid1.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.DataGrid1.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.DataGrid1.HeaderForeColor = System.Drawing.Color.White
        Me.DataGrid1.LinkColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid1.Name = "DataGrid1"
        Me.DataGrid1.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.DataGrid1.ParentRowsForeColor = System.Drawing.Color.Black
        Me.DataGrid1.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.DataGrid1.SelectionForeColor = System.Drawing.Color.White
        Me.DataGrid1.Size = New System.Drawing.Size(672, 312)
        Me.DataGrid1.TabIndex = 12
        Me.DataGrid1.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle3})
        Me.ToolTip1.SetToolTip(Me.DataGrid1, "Problems that may have been caused by customer or Medscreen staff errors, cancell" & _
        "ation charges will be determined by the reviewer later.")
        '
        'XmlSchema11
        '
        Me.XmlSchema11.DataSetName = "XMLSchema1"
        Me.XmlSchema11.Locale = New System.Globalization.CultureInfo("en-GB")
        Me.XmlSchema11.Namespace = "http://tempuri.org/XMLSchema1.xsd"
        '
        'DataGridTableStyle3
        '
        Me.DataGridTableStyle3.AlternatingBackColor = System.Drawing.Color.PaleTurquoise
        Me.DataGridTableStyle3.BackColor = System.Drawing.Color.LightBlue
        Me.DataGridTableStyle3.DataGrid = Me.DataGrid1
        Me.DataGridTableStyle3.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn3, Me.DataGridBoolColumn3})
        Me.DataGridTableStyle3.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle3.MappingName = "Reasons"
        '
        'DataGridTextBoxColumn3
        '
        Me.DataGridTextBoxColumn3.Format = ""
        Me.DataGridTextBoxColumn3.FormatInfo = Nothing
        Me.DataGridTextBoxColumn3.HeaderText = "Description"
        Me.DataGridTextBoxColumn3.MappingName = "Description"
        Me.DataGridTextBoxColumn3.Width = 400
        '
        'DataGridBoolColumn3
        '
        Me.DataGridBoolColumn3.FalseValue = False
        Me.DataGridBoolColumn3.MappingName = "HasReason"
        Me.DataGridBoolColumn3.NullValue = "False"
        Me.DataGridBoolColumn3.TrueValue = True
        Me.DataGridBoolColumn3.Width = 75
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label6})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 312)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(672, 32)
        Me.Panel2.TabIndex = 13
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(200, 7)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Other Reasons"
        '
        'DataGrid3
        '
        Me.DataGrid3.AlternatingBackColor = System.Drawing.Color.LightGray
        Me.DataGrid3.BackColor = System.Drawing.Color.Gainsboro
        Me.DataGrid3.BackgroundColor = System.Drawing.Color.Silver
        Me.DataGrid3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGrid3.CaptionBackColor = System.Drawing.Color.LightSteelBlue
        Me.DataGrid3.CaptionForeColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid3.CaptionText = "Non Chargeable Problems"
        Me.DataGrid3.DataMember = ""
        Me.DataGrid3.DataSource = Me.XmlSchema11
        Me.DataGrid3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGrid3.FlatMode = True
        Me.DataGrid3.Font = New System.Drawing.Font("Tahoma", 8.0!)
        Me.DataGrid3.ForeColor = System.Drawing.Color.Black
        Me.DataGrid3.GridLineColor = System.Drawing.Color.DimGray
        Me.DataGrid3.GridLineStyle = System.Windows.Forms.DataGridLineStyle.None
        Me.DataGrid3.HeaderBackColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid3.HeaderFont = New System.Drawing.Font("Tahoma", 8.0!, System.Drawing.FontStyle.Bold)
        Me.DataGrid3.HeaderForeColor = System.Drawing.Color.White
        Me.DataGrid3.LinkColor = System.Drawing.Color.MidnightBlue
        Me.DataGrid3.Location = New System.Drawing.Point(0, 344)
        Me.DataGrid3.Name = "DataGrid3"
        Me.DataGrid3.ParentRowsBackColor = System.Drawing.Color.DarkGray
        Me.DataGrid3.ParentRowsForeColor = System.Drawing.Color.Black
        Me.DataGrid3.SelectionBackColor = System.Drawing.Color.CadetBlue
        Me.DataGrid3.SelectionForeColor = System.Drawing.Color.White
        Me.DataGrid3.Size = New System.Drawing.Size(672, 241)
        Me.DataGrid3.TabIndex = 16
        Me.DataGrid3.TableStyles.AddRange(New System.Windows.Forms.DataGridTableStyle() {Me.DataGridTableStyle1})
        Me.ToolTip1.SetToolTip(Me.DataGrid3, "Problems that are due to errors and omissions of Medscreen staff")
        '
        'DataGridTableStyle1
        '
        Me.DataGridTableStyle1.AlternatingBackColor = System.Drawing.Color.LightBlue
        Me.DataGridTableStyle1.BackColor = System.Drawing.Color.PaleTurquoise
        Me.DataGridTableStyle1.DataGrid = Me.DataGrid3
        Me.DataGridTableStyle1.GridColumnStyles.AddRange(New System.Windows.Forms.DataGridColumnStyle() {Me.DataGridTextBoxColumn1, Me.DataGridBoolColumn1})
        Me.DataGridTableStyle1.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.DataGridTableStyle1.MappingName = "Reasons"
        '
        'DataGridTextBoxColumn1
        '
        Me.DataGridTextBoxColumn1.Format = ""
        Me.DataGridTextBoxColumn1.FormatInfo = Nothing
        Me.DataGridTextBoxColumn1.HeaderText = "Description"
        Me.DataGridTextBoxColumn1.MappingName = "Description"
        Me.DataGridTextBoxColumn1.Width = 400
        '
        'DataGridBoolColumn1
        '
        Me.DataGridBoolColumn1.FalseValue = False
        Me.DataGridBoolColumn1.HeaderText = "Yes"
        Me.DataGridBoolColumn1.MappingName = "HasReason"
        Me.DataGridBoolColumn1.NullValue = "False"
        Me.DataGridBoolColumn1.TrueValue = True
        Me.DataGridBoolColumn1.Width = 75
        '
        'Panel4
        '
        Me.Panel4.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel4.Location = New System.Drawing.Point(0, 553)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(672, 32)
        Me.Panel4.TabIndex = 19
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(580, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 20
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(492, 3)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 19
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtComments, Me.Label7})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 505)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(672, 48)
        Me.Panel1.TabIndex = 20
        '
        'txtComments
        '
        Me.txtComments.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.txtComments.Location = New System.Drawing.Point(136, 8)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.Size = New System.Drawing.Size(520, 32)
        Me.txtComments.TabIndex = 1
        Me.txtComments.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtComments, "Elaborate on reason for cancellation.")
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(8, 8)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Label7.Size = New System.Drawing.Size(120, 32)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "Provide detailed reason for cancellation"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'frmCancelReason
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 585)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1, Me.Panel4, Me.DataGrid3, Me.Panel2, Me.DataGrid1, Me.Label5, Me.TextBox5, Me.Label4, Me.TextBox4, Me.Label3, Me.TextBox3, Me.Label2, Me.TextBox2, Me.Label1, Me.TextBox1})
        Me.Name = "frmCancelReason"
        Me.Text = "Reasons for cancellation"
        CType(Me.DataGrid1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.XmlSchema11, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGrid3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel4.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private bm As BindingManagerBase
    Private DG As BindingManagerBase
    Dim Collectors As New XMLSchema1()
    Dim Others As New XMLSchema1()
    Dim Customer As New XMLSchema1()

    Private StrReason As String = ""
    Private intReasons As Long = 0


    Public Sub DoBinding(ByVal ds As DataSet)

        Dim tc As New DataGridTableStyle()
        Dim bgc As DataGridBoolColumn
        Dim tgc As DataGridTextBoxColumn
        tc.MappingName = "PHRASE"

        Dim CustomerFaults As DataRow()


        Dim r As DataRow
        Dim c As DataColumn
        Dim values() As Object = {False, ""}

        For Each r In ds.Tables(0).Rows
            If r.Item("PHRASE_ID") < "0000010000" Then
                Collectors.Reasons.AddReasonsRow(False, r.Item("PHRASE_TEXT"), r.Item("PHRASE_ID"))
            Else
                Customer.Reasons.AddReasonsRow(False, r.Item("PHRASE_TEXT"), r.Item("PHRASE_ID"))

            End If
        Next

        Customer.Reasons.AcceptChanges()
        Collectors.Reasons.AcceptChanges()
        Others.Reasons.AcceptChanges()

        DataGrid1.DataSource = Customer.Reasons
        DataGrid3.DataSource = Collectors.Reasons

    End Sub

    Private Sub DataGrid1_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'bm = DG
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim changes As New XMLSchema1()

        Dim temp As XMLSchema1
        Try
            changes.Merge(Customer.GetChanges)
        Catch ex As Exception
        End Try

        Try
            changes.Merge(Collectors.GetChanges)
        Catch ex As Exception
        End Try




        Dim tr As XMLSchema1.ReasonsRow

        For Each tr In changes.Reasons ' Go through all reasons clicked
            If tr.HasReason = True Then ' Check to see if we are still ticked
                Me.intReasons += CLng("&h" & tr.value)
                Me.StrReason += tr.Description & "; "
            End If
        Next
        Me.StrReason += Me.txtComments.Text

    End Sub

    Public Property ReasonString() As String
        Get
            Return StrReason
        End Get
        Set(ByVal Value As String)
            StrReason = Value
        End Set
    End Property

    Public Property NumericReason() As Long
        Get
            Return Me.intReasons
        End Get
        Set(ByVal Value As Long)
            Me.intReasons = Value
        End Set
    End Property

    Private Sub frmCancelReason_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Me.ErrorProvider1.SetError(Me.DataGrid1, "")
        Me.ErrorProvider1.SetError(Me.txtComments, "")


        If Me.DialogResult = DialogResult.OK And Me.intReasons = 0 Then
            MsgBox("A reason must be specified for the cancellation")
            Me.ErrorProvider1.SetError(Me.DataGrid1, "A reason must be specified for the cancellation")
            e.Cancel = True
        End If

        If Me.DialogResult = DialogResult.OK And Me.txtComments.Text.Length < 5 Then
            MsgBox("A descriptive commentry for the reason for cancelling must be given to procede")
            Me.ErrorProvider1.SetError(Me.txtComments, "A descriptive commentry for the reason for cancelling must be given to procede")
            e.Cancel = True
        End If
    End Sub
End Class
