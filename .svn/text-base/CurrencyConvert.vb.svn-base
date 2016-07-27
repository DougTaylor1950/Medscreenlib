Imports MedscreenLib
Public Class CurrencyConvert
    Inherits System.Windows.Forms.Form

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Starting Currency
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	25/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property InCurrency() As String
        Get
            Return strInCurrency
        End Get
        Set(ByVal Value As String)
            strInCurrency = Value
            Me.cbCurrency.SelectedValue = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' resultant currency
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	25/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property OutCurrency() As String
        Get
            Return strOutCurrency
        End Get
        Set(ByVal Value As String)
            strOutCurrency = Value
            Me.cbOutCurrency.SelectedValue = Value
            cr = New Globalization.CultureInfo(MedscreenLib.Glossary.Glossary.Currencies.Culture(Me.strOutCurrency))

        End Set
    End Property

    Public Property LockCurrency() As Boolean
        Get
            Return Not cbOutCurrency.Enabled
        End Get
        Set(ByVal Value As Boolean)
            cbOutCurrency.Enabled = Not Value
        End Set
    End Property
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim strCurrencyList As String = MedscreenCommonGUIConfig.CollectionQuery.Item("CurrencyList")
        Dim objCurrencyList As MedscreenLib.Glossary.PhraseCollection = New Glossary.PhraseCollection(strCurrencyList, Glossary.PhraseCollection.BuildBy.Query)
        Dim objCurrencyListOut As MedscreenLib.Glossary.PhraseCollection = New Glossary.PhraseCollection(strCurrencyList, Glossary.PhraseCollection.BuildBy.Query)
        Me.cbCurrency.DataSource = objCurrencyList
        Me.cbCurrency.DisplayMember = "PhraseText"
        Me.cbCurrency.ValueMember = "PhraseID"


        Me.cbOutCurrency.DataSource = objCurrencyListOut
        Me.cbOutCurrency.DisplayMember = "PhraseText"
        Me.cbOutCurrency.ValueMember = "PhraseID"
        cr = New Globalization.CultureInfo(MedscreenLib.Glossary.Glossary.Currencies.Culture(Me.strOutCurrency))

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
    Friend WithEvents cmdConvert As System.Windows.Forms.Button
    Friend WithEvents cbCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents txtAddCosts As System.Windows.Forms.TextBox
    Friend WithEvents lblAddCosts As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cbOutCurrency As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtResult As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CurrencyConvert))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.cmdConvert = New System.Windows.Forms.Button()
        Me.cbCurrency = New System.Windows.Forms.ComboBox()
        Me.txtAddCosts = New System.Windows.Forms.TextBox()
        Me.lblAddCosts = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbOutCurrency = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtResult = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 217)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(424, 56)
        Me.Panel1.TabIndex = 2
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(336, 16)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 11
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(256, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 10
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdConvert
        '
        Me.cmdConvert.Image = CType(resources.GetObject("cmdConvert.Image"), System.Drawing.Bitmap)
        Me.cmdConvert.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdConvert.Location = New System.Drawing.Point(328, 160)
        Me.cmdConvert.Name = "cmdConvert"
        Me.cmdConvert.Size = New System.Drawing.Size(80, 23)
        Me.cmdConvert.TabIndex = 100
        Me.cmdConvert.Text = "Convert"
        Me.cmdConvert.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCurrency
        '
        Me.cbCurrency.ItemHeight = 13
        Me.cbCurrency.Location = New System.Drawing.Point(88, 24)
        Me.cbCurrency.Name = "cbCurrency"
        Me.cbCurrency.Size = New System.Drawing.Size(152, 21)
        Me.cbCurrency.TabIndex = 97
        '
        'txtAddCosts
        '
        Me.txtAddCosts.Location = New System.Drawing.Point(328, 24)
        Me.txtAddCosts.Name = "txtAddCosts"
        Me.txtAddCosts.Size = New System.Drawing.Size(80, 20)
        Me.txtAddCosts.TabIndex = 98
        Me.txtAddCosts.Text = ""
        '
        'lblAddCosts
        '
        Me.lblAddCosts.Location = New System.Drawing.Point(0, 24)
        Me.lblAddCosts.Name = "lblAddCosts"
        Me.lblAddCosts.Size = New System.Drawing.Size(88, 23)
        Me.lblAddCosts.TabIndex = 99
        Me.lblAddCosts.Text = "input Currency"
        Me.lblAddCosts.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(0, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 23)
        Me.Label1.TabIndex = 101
        Me.Label1.Text = "Output Currency"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(272, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 102
        Me.Label2.Text = "amount"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cbOutCurrency
        '
        Me.cbOutCurrency.ItemHeight = 13
        Me.cbOutCurrency.Location = New System.Drawing.Point(88, 64)
        Me.cbOutCurrency.Name = "cbOutCurrency"
        Me.cbOutCurrency.Size = New System.Drawing.Size(152, 21)
        Me.cbOutCurrency.TabIndex = 103
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(272, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(56, 23)
        Me.Label3.TabIndex = 105
        Me.Label3.Text = "amount"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtResult
        '
        Me.txtResult.Location = New System.Drawing.Point(328, 64)
        Me.txtResult.Name = "txtResult"
        Me.txtResult.Size = New System.Drawing.Size(80, 20)
        Me.txtResult.TabIndex = 104
        Me.txtResult.Text = ""
        '
        'CurrencyConvert
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(424, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label3, Me.txtResult, Me.cbOutCurrency, Me.Label2, Me.Label1, Me.cmdConvert, Me.cbCurrency, Me.txtAddCosts, Me.lblAddCosts, Me.Panel1})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "CurrencyConvert"
        Me.Text = "Currency Convertor"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private cr As System.Globalization.CultureInfo
    Private strInCurrency As String = "GBP"
    Private strOutCurrency As String = "GBP"
    Private AgreedExpenses As Double
    Private blnStarting As Boolean = True

    Private Sub cmdConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConvert.Click
        'With myCollection

        'If we don't have client info (should never occur)
        'Addtional Costs 

        Try
            'There may be a currency symbol in the first place, in a try block in case it goes do lalhi
            If Me.txtAddCosts.Text.Chars(0) < "0" Or Me.txtAddCosts.Text.Chars(0) > "9" Then
                AgreedExpenses = Mid(Me.txtAddCosts.Text, 2)
            Else
                AgreedExpenses = Me.txtAddCosts.Text
            End If
        Catch ex As Exception
            AgreedExpenses = 0
        End Try

        'See if the currency selected is the same as the customer's if not convert
        If Me.cbCurrency.SelectedValue <> Me.strOutCurrency Then
            Dim strin As String
            Dim strout As String

            'No match set up for conversion, conversion use ISO three character names.
            strin = Me.strInCurrency
            'Deal with collecting officer currencies
            strout = Me.cbOutCurrency.SelectedValue
            'Use conversion routines and current conversion rate, 
            'if any errors will almost certainly be caused by the conversion rate not being in the table
            AgreedExpenses = MedscreenLib.Glossary.Glossary.ExchangeRates.Convert(strin, strout, AgreedExpenses)
        End If
        Me.txtResult.Text = AgreedExpenses.ToString("c", cr)
        'Me.cbCurrency.SelectedValue = Me.strInCurrency
        'End If
        'End With
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return the value 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Value() As Double
        Get
            Return AgreedExpenses
        End Get
        Set(ByVal Value As Double)
            If Not cr Is Nothing Then
                Me.txtAddCosts.Text = Value.ToString("c", cr)
            Else
                Me.txtAddCosts.Text = Value.ToString

            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the discovered culture 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property Culture() As Globalization.CultureInfo
        Get
            Return cr
        End Get
    End Property

    Private Sub cbCurrency_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCurrency.SelectedIndexChanged
        If blnStarting Then Exit Sub
        Me.strInCurrency = Me.cbCurrency.SelectedValue
        Me.cmdConvert_Click(Nothing, Nothing)
    End Sub

    Private Sub cbOutCurrency_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOutCurrency.SelectedIndexChanged
        If blnStarting Then Exit Sub
        Me.strOutCurrency = Me.cbOutCurrency.SelectedValue
        cr = New Globalization.CultureInfo(MedscreenLib.Glossary.Glossary.Currencies.Culture(Me.strOutCurrency))
        Me.cmdConvert_Click(Nothing, Nothing)
    End Sub

    Private Sub CurrencyConvert_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        blnStarting = False
    End Sub

    Private Sub txtAddCosts_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddCosts.Leave
        If blnStarting Then Exit Sub
        Me.cmdConvert_Click(Nothing, Nothing)
    End Sub
End Class
