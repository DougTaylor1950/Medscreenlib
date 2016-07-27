Public Class frmMatrix
    Inherits System.Windows.Forms.Form
    Private objMatrixPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objTSPhrase As MedscreenLib.Glossary.PhraseCollection
    Private objProduct As MedscreenLib.Glossary.PhraseCollection
    Private objFlawSpec As MedscreenLib.Glossary.PhraseCollection
    Private objFlawProduct As MedscreenLib.Glossary.PhraseCollection

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        objMatrixPhrase = New MedscreenLib.Glossary.PhraseCollection("MATRIX")
        objMatrixPhrase.Load()

        Me.objTSPhrase = New MedscreenLib.Glossary.PhraseCollection("Lib_cctool.GetTestSchedulesNonFlaws", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction, , ";")

        Me.objProduct = New MedscreenLib.Glossary.PhraseCollection("select identity,description from mlp_matrix where PRODUCT_GROUP = 'TOPLEVEL'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)

        Me.objFlawSpec = New MedscreenLib.Glossary.PhraseCollection("Lib_cctool.GetTestSchedulesFlaws", MedscreenLib.Glossary.PhraseCollection.BuildBy.PLSQLFunction, , ";")

        Me.objFlawProduct = New MedscreenLib.Glossary.PhraseCollection("select identity,description from mlp_header where PRODUCT_GROUP = 'FLAWS'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)

        Me.cbMatrix.DataSource = objMatrixPhrase
        Me.cbMatrix.DisplayMember = "PhraseText"
        Me.cbMatrix.ValueMember = "PhraseID"

        Me.cbTestSchedule.DataSource = objTSPhrase
        Me.cbTestSchedule.DisplayMember = "PhraseText"
        Me.cbTestSchedule.ValueMember = "PhraseID"

        Me.cbProduct.DataSource = objProduct
        Me.cbProduct.DisplayMember = "PhraseText"
        Me.cbProduct.ValueMember = "PhraseID"

        Me.cbFlawSchedule.DataSource = objFlawSpec
        Me.cbFlawSchedule.DisplayMember = "PhraseText"
        Me.cbFlawSchedule.ValueMember = "PhraseID"

        Me.cbFlawProduct.DataSource = objFlawProduct
        Me.cbFlawProduct.DisplayMember = "PhraseText"
        Me.cbFlawProduct.ValueMember = "PhraseID"


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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbMatrix As System.Windows.Forms.ComboBox
    Friend WithEvents cbTestSchedule As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbProduct As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cbFlawSchedule As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbFlawProduct As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMatrix))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbMatrix = New System.Windows.Forms.ComboBox()
        Me.cbTestSchedule = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbProduct = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cbFlawSchedule = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbFlawProduct = New System.Windows.Forms.ComboBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 205)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(672, 56)
        Me.Panel1.TabIndex = 11
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(584, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 1
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(504, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 0
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Sample type (Matrix)"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbMatrix
        '
        Me.cbMatrix.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbMatrix.Location = New System.Drawing.Point(176, 32)
        Me.cbMatrix.Name = "cbMatrix"
        Me.cbMatrix.Size = New System.Drawing.Size(480, 21)
        Me.cbMatrix.TabIndex = 13
        '
        'cbTestSchedule
        '
        Me.cbTestSchedule.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbTestSchedule.Location = New System.Drawing.Point(176, 64)
        Me.cbTestSchedule.Name = "cbTestSchedule"
        Me.cbTestSchedule.Size = New System.Drawing.Size(480, 21)
        Me.cbTestSchedule.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(48, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Test panel"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbProduct
        '
        Me.cbProduct.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbProduct.Location = New System.Drawing.Point(176, 96)
        Me.cbProduct.Name = "cbProduct"
        Me.cbProduct.Size = New System.Drawing.Size(480, 21)
        Me.cbProduct.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 96)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(136, 23)
        Me.Label3.TabIndex = 16
        Me.Label3.Text = "Product spec (cut-offs)"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbFlawSchedule
        '
        Me.cbFlawSchedule.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbFlawSchedule.Location = New System.Drawing.Point(176, 128)
        Me.cbFlawSchedule.Name = "cbFlawSchedule"
        Me.cbFlawSchedule.Size = New System.Drawing.Size(480, 21)
        Me.cbFlawSchedule.TabIndex = 19
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(48, 128)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "Flaws list"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbFlawProduct
        '
        Me.cbFlawProduct.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.cbFlawProduct.Location = New System.Drawing.Point(176, 160)
        Me.cbFlawProduct.Name = "cbFlawProduct"
        Me.cbFlawProduct.Size = New System.Drawing.Size(480, 21)
        Me.cbFlawProduct.TabIndex = 21
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 160)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 23)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "Cancellation criteria"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'frmMatrix
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(672, 261)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.cbFlawProduct, Me.Label5, Me.cbFlawSchedule, Me.Label4, Me.cbProduct, Me.Label3, Me.cbTestSchedule, Me.Label2, Me.cbMatrix, Me.Label1, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmMatrix"
        Me.Text = "Matrix editing"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#If r2004 Then
    Private myMatrix As Intranet.intranet.customerns.CustomerMatrix
    Private FormFilled As Boolean = False
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Matrix to edit
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	08/09/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Matrix() As Intranet.intranet.customerns.CustomerMatrix
        Get
            Return myMatrix
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.CustomerMatrix)
            myMatrix = Value
            Me.objProduct = New MedscreenLib.Glossary.PhraseCollection("select identity,description from mlp_matrix where PRODUCT_GROUP = 'TOPLEVEL' and matrices like '%" & myMatrix.Matrix & "%'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
            Me.cbProduct.DataSource = Me.objProduct
            Me.cbProduct.DisplayMember = "PhraseText"
            Me.cbProduct.ValueMember = "PhraseID"

            FillForm()
        End Set
    End Property

    Private Sub FillForm()
        If Me.myMatrix Is Nothing Then Exit Sub

        Me.cbMatrix.SelectedValue = myMatrix.Matrix
        Me.cbTestSchedule.SelectedValue = myMatrix.ScheduleId
        Me.cbProduct.SelectedValue = myMatrix.ProductId
        Me.cbFlawSchedule.SelectedValue = myMatrix.FlawSchedule
        Me.cbFlawProduct.SelectedValue = myMatrix.FlawSpec
        Me.FormFilled = True
    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Me.myMatrix Is Nothing Then Exit Sub

        Me.ErrorProvider1.SetError(Me.cbProduct, "")
        myMatrix.Matrix = Me.cbMatrix.SelectedValue
        myMatrix.ScheduleId = Me.cbTestSchedule.SelectedValue
        myMatrix.ProductId = Me.cbProduct.SelectedValue
        myMatrix.FlawSchedule = Me.cbFlawSchedule.SelectedValue
        myMatrix.FlawSpec = Me.cbFlawProduct.SelectedValue
        If myMatrix.ProductId.Trim.Length = 0 Then
            Me.DialogResult = Windows.Forms.DialogResult.None
            Me.ErrorProvider1.SetError(Me.cbProduct, "A cut-off must be provided")
        Else
            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If


    End Sub

    Private Sub cbMatrix_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbMatrix.SelectedValueChanged
        If Not Me.FormFilled Then Exit Sub
        If myMatrix Is Nothing Then Exit Sub

        Me.objProduct = New MedscreenLib.Glossary.PhraseCollection("select identity,description from mlp_matrix where PRODUCT_GROUP = 'TOPLEVEL' and matrices like '%" & Me.cbMatrix.SelectedValue & "%'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        Me.cbProduct.DataSource = Me.objProduct
        Me.cbProduct.DisplayMember = "PhraseText"
        Me.cbProduct.ValueMember = "PhraseID"
        Me.cbProduct.SelectedIndex = -1
        Me.cbProduct.SelectedValue = myMatrix.ProductId

    End Sub
#End If

    Private Sub cbProduct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbProduct.SelectedIndexChanged
        'If Not Me.FormFilled Then Exit Sub
        'If myMatrix Is Nothing Then Exit Sub

        'Me.objProduct = New MedscreenLib.Glossary.PhraseCollection("select identity,description from mlp_matrix where PRODUCT_GROUP = 'TOPLEVEL' and matrices like '%" & Me.cbMatrix.SelectedValue & "%'", MedscreenLib.Glossary.PhraseCollection.BuildBy.Query)
        'Me.cbProduct.DataSource = Me.objProduct
        'Me.cbProduct.DisplayMember = "PhraseText"
        'Me.cbProduct.ValueMember = "PhraseID"
        '' Me.cbProduct.SelectedValue = myMatrix.ProductId

    End Sub
End Class
