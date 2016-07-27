Imports MedscreenLib
Public Class frmCurrencyEntry
    Inherits System.Windows.Forms.Form
    Private myInputCurrency As String
    Private myOutputCurrency As String
    Private myInputAmount As Double
    Private myOutputAmount As Double
    Private outputCr As System.Globalization.CultureInfo
    Private InputCr As System.Globalization.CultureInfo

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim strCurrencyList As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CurrencyList")
        Dim objCurrencyList As MedscreenLib.Glossary.PhraseCollection = New MedscreenLib.Glossary.PhraseCollection(strCurrencyList, Glossary.PhraseCollection.BuildBy.Query)
        Me.cbInputCurrency.DataSource = objCurrencyList
        Me.cbInputCurrency.DisplayMember = "PhraseText"
        Me.cbInputCurrency.ValueMember = "PhraseID"
        Me.cbOutputCurrency.DataSource = objCurrencyList
        Me.cbOutputCurrency.DisplayMember = "PhraseText"
        Me.cbOutputCurrency.ValueMember = "PhraseID"

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
    Friend WithEvents pnlControls As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents pnlData As System.Windows.Forms.Panel
    Friend WithEvents cbInputCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbOutputCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents txtConvertedAmount As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cmdConvert As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmCurrencyEntry))
        Me.pnlControls = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.pnlData = New System.Windows.Forms.Panel()
        Me.cbInputCurrency = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbOutputCurrency = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.txtConvertedAmount = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cmdConvert = New System.Windows.Forms.Button()
        Me.pnlControls.SuspendLayout()
        Me.pnlData.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlControls
        '
        Me.pnlControls.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlControls.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.pnlControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlControls.Location = New System.Drawing.Point(0, 157)
        Me.pnlControls.Name = "pnlControls"
        Me.pnlControls.Size = New System.Drawing.Size(448, 48)
        Me.pnlControls.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(360, 16)
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
        Me.cmdOk.Location = New System.Drawing.Point(280, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.TabIndex = 4
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlData
        '
        Me.pnlData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlData.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdConvert, Me.txtConvertedAmount, Me.Label4, Me.txtAmount, Me.Label3, Me.Label2, Me.cbOutputCurrency, Me.Label1, Me.cbInputCurrency})
        Me.pnlData.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlData.Name = "pnlData"
        Me.pnlData.Size = New System.Drawing.Size(448, 157)
        Me.pnlData.TabIndex = 1
        '
        'cbInputCurrency
        '
        Me.cbInputCurrency.ItemHeight = 13
        Me.cbInputCurrency.Location = New System.Drawing.Point(144, 16)
        Me.cbInputCurrency.Name = "cbInputCurrency"
        Me.cbInputCurrency.Size = New System.Drawing.Size(48, 21)
        Me.cbInputCurrency.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Input Currency"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Output Currency"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbOutputCurrency
        '
        Me.cbOutputCurrency.ItemHeight = 13
        Me.cbOutputCurrency.Location = New System.Drawing.Point(144, 48)
        Me.cbOutputCurrency.Name = "cbOutputCurrency"
        Me.cbOutputCurrency.Size = New System.Drawing.Size(48, 21)
        Me.cbOutputCurrency.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(216, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Amount"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAmount
        '
        Me.txtAmount.Location = New System.Drawing.Point(328, 16)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.TabIndex = 10
        Me.txtAmount.Text = ""
        '
        'txtConvertedAmount
        '
        Me.txtConvertedAmount.Enabled = False
        Me.txtConvertedAmount.Location = New System.Drawing.Point(328, 48)
        Me.txtConvertedAmount.Name = "txtConvertedAmount"
        Me.txtConvertedAmount.TabIndex = 12
        Me.txtConvertedAmount.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(216, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Converted Amount"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdConvert
        '
        Me.cmdConvert.Image = CType(resources.GetObject("cmdConvert.Image"), System.Drawing.Bitmap)
        Me.cmdConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdConvert.Location = New System.Drawing.Point(336, 88)
        Me.cmdConvert.Name = "cmdConvert"
        Me.cmdConvert.Size = New System.Drawing.Size(80, 23)
        Me.cmdConvert.TabIndex = 97
        Me.cmdConvert.Text = "Convert"
        Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'frmCurrencyEntry
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(448, 205)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.pnlData, Me.pnlControls})
        Me.Name = "frmCurrencyEntry"
        Me.Text = "Currency Entry"
        Me.pnlControls.ResumeLayout(False)
        Me.pnlData.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Currency to convert from 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/01/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property InputCurrency() As String
        Get
            Return myInputCurrency
        End Get
        Set(ByVal Value As String)
            myInputCurrency = Value
            Me.cbInputCurrency.SelectedValue = Value
            Me.InputCr = New Globalization.CultureInfo(MedscreenLib.Glossary.Glossary.Currencies.Culture(Value))

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Currency to convert to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [12/01/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property OutputCurrency() As String
        Get
            Return myOutputCurrency
        End Get
        Set(ByVal Value As String)
            myOutputCurrency = Value
            Me.outputCr = New Globalization.CultureInfo(MedscreenLib.Glossary.Glossary.Currencies.Culture(Value))
        End Set
    End Property

    Public Property InputAmount() As Double
        Get
            Return myInputAmount
        End Get
        Set(ByVal Value As Double)
            myInputAmount = Value
            Me.txtAmount.Text = Value.ToString("c", InputCr)
        End Set
    End Property

    Public Property OutputAmount() As Double
        Get
            Return myOutputAmount
        End Get
        Set(ByVal Value As Double)
            myOutputAmount = Value
        End Set
    End Property

#End Region

    Private Sub cmdConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConvert.Click

        'If we don't have client info (should never occur)
        'Addtional Costs 
        Try
            'There may be a currency symbol in the first place, in a try block in case it goes do lalhi
            If Me.txtAmount.Text.Chars(0) < "0" Or Me.txtAmount.Text.Chars(0) > "9" Then
                Me.InputAmount = Mid(Me.txtAmount.Text, 2)
            Else
                InputAmount = Me.txtAmount.Text
            End If
        Catch ex As Exception
            InputAmount = 0
        End Try

        'See if the currency selected is the same as the customer's if not convert
        If Me.cbInputCurrency.Text <> Me.OutputCurrency Then
            Dim strin As String
            Dim strout As String

            'No match set up for conversion, conversion use ISO three character names.
            strin = Me.OutputCurrency
            'Deal with collecting officer currencies
            strout = Me.cbInputCurrency.SelectedValue
            'Use conversion routines and current conversion rate, 
            'if any errors will almost certainly be caused by the conversion rate not being in the table
            Me.OutputAmount = MedscreenLib.Glossary.Glossary.ExchangeRates.Convert(strout, strin, InputAmount)
            Me.txtConvertedAmount.Text = OutputAmount.ToString("c", outputCr)
            'Me.cbInputCurrency.SelectedValue = Me.CustomerCurrency
        End If
    End Sub

End Class
