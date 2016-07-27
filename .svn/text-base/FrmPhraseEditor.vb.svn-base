
Imports MedscreenLib.Glossary
Imports System.Drawing
Public Class FrmPhraseEditor
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Dim aColorName As String
        For Each aColorName In _
           System.Enum.GetNames _
           (GetType(System.Drawing.KnownColor))
            ForeColorCombo.Items.Add(Drawing.Color.FromName(aColorName))
        Next
        For Each aColorName In _
           System.Enum.GetNames _
           (GetType(System.Drawing.KnownColor))
            BackColorCombo.Items.Add(Drawing.Color.FromName(aColorName))
        Next

        'StatusPhrase = New PhraseCollection("PHRASESTAT")
        'StatusPhrase.Load()
        Me.cbStatus.PhraseType = "PHRASESTAT"
        'Me.cbStatus.DisplayMember = "PhraseText"
        'Me.cbStatus.ValueMember = "PhraseID"
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
    Friend WithEvents ColorDialog1 As System.Windows.Forms.ColorDialog
    Friend WithEvents FontDialog1 As System.Windows.Forms.FontDialog
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPhraseID As System.Windows.Forms.TextBox
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtIndex As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtFont As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ForeColorCombo As System.Windows.Forms.ComboBox
    Friend WithEvents BackColorCombo As System.Windows.Forms.ComboBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents udImageIndex As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cbStatus As MedscreenCommonGui.PhraseComboBox
    Friend WithEvents txtIcon As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmPhraseEditor))
        Me.ColorDialog1 = New System.Windows.Forms.ColorDialog()
        Me.FontDialog1 = New System.Windows.Forms.FontDialog()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtPhraseID = New System.Windows.Forms.TextBox()
        Me.txtDescription = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtIndex = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtFont = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.ForeColorCombo = New System.Windows.Forms.ComboBox()
        Me.BackColorCombo = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.udImageIndex = New System.Windows.Forms.NumericUpDown()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.cbStatus = New MedscreenCommonGui.PhraseComboBox()
        Me.txtIcon = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider()
        Me.pnlBottom.SuspendLayout()
        CType(Me.udImageIndex, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 421)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(768, 56)
        Me.pnlBottom.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(656, 16)
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
        Me.cmdOk.Location = New System.Drawing.Point(576, 16)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 2
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(24, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Phrase ID"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPhraseID
        '
        Me.txtPhraseID.Location = New System.Drawing.Point(136, 16)
        Me.txtPhraseID.Name = "txtPhraseID"
        Me.txtPhraseID.Size = New System.Drawing.Size(112, 20)
        Me.txtPhraseID.TabIndex = 2
        Me.txtPhraseID.Text = ""
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(136, 48)
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(608, 20)
        Me.txtDescription.TabIndex = 4
        Me.txtDescription.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(24, 48)
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Description"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtIndex
        '
        Me.txtIndex.Location = New System.Drawing.Point(136, 80)
        Me.txtIndex.Name = "txtIndex"
        Me.txtIndex.Size = New System.Drawing.Size(112, 20)
        Me.txtIndex.TabIndex = 6
        Me.txtIndex.Text = ""
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.TabIndex = 5
        Me.Label3.Text = "Order Num"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(24, 120)
        Me.Label4.Name = "Label4"
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Fore Colour"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(296, 120)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 9
        Me.Label5.Text = "Back Colour"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFont
        '
        Me.txtFont.Location = New System.Drawing.Point(136, 152)
        Me.txtFont.Name = "txtFont"
        Me.txtFont.Size = New System.Drawing.Size(392, 20)
        Me.txtFont.TabIndex = 12
        Me.txtFont.Text = ""
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 152)
        Me.Label6.Name = "Label6"
        Me.Label6.TabIndex = 11
        Me.Label6.Text = "Font"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ForeColorCombo
        '
        Me.ForeColorCombo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.ForeColorCombo.Location = New System.Drawing.Point(136, 120)
        Me.ForeColorCombo.Name = "ForeColorCombo"
        Me.ForeColorCombo.Size = New System.Drawing.Size(121, 21)
        Me.ForeColorCombo.TabIndex = 13
        '
        'BackColorCombo
        '
        Me.BackColorCombo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable
        Me.BackColorCombo.Location = New System.Drawing.Point(408, 120)
        Me.BackColorCombo.Name = "BackColorCombo"
        Me.BackColorCombo.Size = New System.Drawing.Size(121, 21)
        Me.BackColorCombo.TabIndex = 14
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(16, 192)
        Me.Label7.Name = "Label7"
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Image Index"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'udImageIndex
        '
        Me.udImageIndex.Location = New System.Drawing.Point(136, 192)
        Me.udImageIndex.Name = "udImageIndex"
        Me.udImageIndex.TabIndex = 16
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(296, 16)
        Me.Label8.Name = "Label8"
        Me.Label8.TabIndex = 17
        Me.Label8.Text = "Status"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbStatus
        '
        Me.cbStatus.Location = New System.Drawing.Point(408, 16)
        Me.cbStatus.Name = "cbStatus"
        Me.cbStatus.Size = New System.Drawing.Size(121, 21)
        Me.cbStatus.TabIndex = 18
        'Me.cbStatus.PhraseType = "PHRASESTAT"
        '
        'txtIcon
        '
        Me.txtIcon.Location = New System.Drawing.Point(408, 80)
        Me.txtIcon.Name = "txtIcon"
        Me.txtIcon.Size = New System.Drawing.Size(112, 20)
        Me.txtIcon.TabIndex = 20
        Me.txtIcon.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(296, 80)
        Me.Label9.Name = "Label9"
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "Icon"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.DataMember = Nothing
        '
        'FrmPhraseEditor
        '
        Me.AcceptButton = Me.cmdOk
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(768, 477)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtIcon, Me.Label9, Me.cbStatus, Me.Label8, Me.udImageIndex, Me.Label7, Me.BackColorCombo, Me.ForeColorCombo, Me.txtFont, Me.Label6, Me.Label5, Me.Label4, Me.txtIndex, Me.Label3, Me.txtDescription, Me.Label2, Me.txtPhraseID, Me.Label1, Me.pnlBottom})
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmPhraseEditor"
        Me.Text = "Phrase Editor"
        Me.pnlBottom.ResumeLayout(False)
        CType(Me.udImageIndex, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private myPhrase As Phrase
    Private Drawn As Boolean = False
    Private StatusPhrase As PhraseCollection

    Public Property Phrase() As Phrase
        Get
            GetForm()
            Return myPhrase
        End Get
        Set(ByVal Value As Phrase)
            myPhrase = Value
            FillForm()
        End Set
    End Property

    Private Sub FillForm()
        With myPhrase
            Me.txtPhraseID.Text = .PhraseID
            Me.txtDescription.Text = .PhraseText
            Me.txtIcon.Text = .Icon
            Dim aFont As Drawing.Font = .Formatter.GetFont
            If Not aFont Is Nothing Then
                Me.txtFont.Text = aFont.Name & " (" & aFont.SizeInPoints & ") "
                Me.txtFont.Font = aFont
                If aFont.Italic Then Me.txtFont.Text += "Italic "
                If aFont.Bold Then Me.txtFont.Text += "Bold "
                If aFont.Underline Then Me.txtFont.Text += "UnderLine "
                If aFont.Strikeout Then Me.txtFont.Text += "StrikeOut "
            End If
            Try
                Dim intIndex As Integer = CType(Me.cbStatus.DataSource, MedscreenLib.Glossary.PhraseCollection).IndexOf(.PhraseStatus)
                Me.cbStatus.SelectedIndex = intIndex
                'Me.cbStatus.SelectedValue = .PhraseStatus
                'If Not Me.cbStatus.SelectedItem Is Nothing Then CType(Me.cbStatus.SelectedItem, Phrase).Formatter.Decorate(cbStatus)
                Me.udImageIndex.Value = .Formatter.ImageIndex
                Dim intpos As Integer = Me.ForeColorCombo.FindString("Color [" & .Formatter.ForeColour & "]")
                If intpos > 0 Then Me.ForeColorCombo.SelectedIndex = intpos
                Me.ForeColorCombo.ForeColor = .Formatter.ForeColor
                Me.txtIndex.Text = .OrderNum().PadLeft(10, " ")
                intpos = Me.BackColorCombo.FindString("Color [" & .Formatter.BackColour & "]")
                If intpos > 0 Then Me.BackColorCombo.SelectedIndex = intpos
                Me.BackColorCombo.BackColor = Me.BackColorCombo.SelectedItem
                Me.BackColorCombo.ForeColor = Me.ForeColorCombo.SelectedItem
                Me.ForeColorCombo.BackColor = Me.BackColorCombo.SelectedItem
                Me.ForeColorCombo.ForeColor = Me.ForeColorCombo.SelectedItem
            Catch EX As Exception
            End Try
        End With
    End Sub

    Private mRand As New Random()

    Protected Sub colorCombo_MeasureItem( _
      ByVal sender As Object, ByVal e As _
      System.Windows.Forms.MeasureItemEventArgs) _
      Handles ForeColorCombo.MeasureItem, BackColorCombo.MeasureItem

        e.ItemHeight = mRand.Next(15, 16)

    End Sub


    Protected Sub colorCombo_DrawItem( _
      ByVal sender As Object, _
      ByVal e As System.Windows.Forms.DrawItemEventArgs) _
      Handles ForeColorCombo.DrawItem, BackColorCombo.DrawItem

        If e.Index < 0 Then
            e.DrawBackground()
            e.DrawFocusRectangle()
            Exit Sub
        End If
        ' Get the Color object from the Items list
        Dim aColor As Color = _
           CType(ForeColorCombo.Items(e.Index), Color)

        ' get a square using the bounds height
        Dim rect As Rectangle = New Rectangle _
           (2, e.Bounds.Top + 2, e.Bounds.Height, _
           e.Bounds.Height - 4)

        Dim br As Brush

        ' call these methods first
        e.DrawBackground()
        e.DrawFocusRectangle()

        ' change brush color if item is selected
        If e.State = _
           Windows.Forms.DrawItemState.Selected Then
            br = Brushes.White
        Else
            br = Brushes.Black
        End If

        ' draw a rectangle and fill it
        e.Graphics.DrawRectangle(New Pen(aColor), rect)
        e.Graphics.FillRectangle(New SolidBrush _
           (aColor), rect)

        ' draw a border
        rect.Inflate(1, 1)
        e.Graphics.DrawRectangle(Pens.Black, rect)

        ' draw the Color name
        e.Graphics.DrawString(aColor.Name, _
           ForeColorCombo.Font, br, e.Bounds.Height + 5, _
           ((e.Bounds.Height - ForeColorCombo.Font.Height) _
            \ 2) + e.Bounds.Top)
    End Sub


    Private Function GetForm()
        Dim blnRet As Boolean = True
        Me.ErrorProvider1.SetError(Me.txtPhraseID, "")
        With myPhrase
            .PhraseText = Me.txtDescription.Text
            .OrderNum = Me.txtIndex.Text.PadLeft(10, " ")
            .Formatter.ImageIndex = Me.udImageIndex.Value
            .PhraseStatus = Me.cbStatus.SelectedValue
            .Icon = Me.txtIcon.Text
            If Me.txtPhraseID.Text = "<Replace>" Then
                blnRet = False
                Me.ErrorProvider1.SetError(Me.txtPhraseID, "Id must be changed")
            End If
            .PhraseID = Me.txtPhraseID.Text
            .GenText1 = .Formatter.SerialiseToXML()
        End With
        Return blnRet
    End Function

    Private Sub txtFont_Doubleclick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFont.DoubleClick
        Dim aFont As Drawing.Font = myPhrase.Formatter.GetFont
        Me.FontDialog1.Font = aFont
        If Me.FontDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            aFont = Me.FontDialog1.Font
            myPhrase.Formatter.FontName = aFont.Name
            myPhrase.Formatter.FontItalic = aFont.Italic
            myPhrase.Formatter.FontBold = aFont.Bold
            myPhrase.Formatter.FontStrikeout = aFont.Strikeout
            myPhrase.Formatter.FontUnderline = aFont.Underline
            myPhrase.Formatter.FontSize = aFont.SizeInPoints
            FillForm()

        End If
    End Sub



    Private Sub txtBackColour_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.ColorDialog1.Color = myPhrase.Formatter.backColor
        Me.ColorDialog1.SolidColorOnly = True

        Me.ColorDialog1.AllowFullOpen = True
        If Me.ColorDialog1.ShowDialog Then
            myPhrase.Formatter.BackColour = Me.ColorDialog1.Color.Name
            Me.FillForm()
        End If
    End Sub

    Private Sub ForeColorCombo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ForeColorCombo.SelectedValueChanged
        If Drawn Then
            myPhrase.Formatter.ForeColour = CType(Me.ForeColorCombo.SelectedItem, Color).Name
            FillForm()
        End If
    End Sub

    Private Sub BackColorCombo_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles BackColorCombo.SelectedValueChanged
        If Drawn Then
            myPhrase.Formatter.BackColour = CType(Me.BackColorCombo.SelectedItem, Color).Name
            FillForm()
        End If
    End Sub

    Private Sub FrmPhraseEditor_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Drawn = True
    End Sub

    Private Sub cbStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbStatus.SelectedIndexChanged
        If Drawn Then
            '    If Not Me.cbStatus.SelectedItem Is Nothing Then CType(Me.cbStatus.SelectedItem, Phrase).Formatter.Decorate(cbStatus)
        End If
    End Sub

    Private Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        If Not GetForm() Then
            Me.AcceptButton = Nothing
            Me.DialogResult = Windows.Forms.DialogResult.None
        Else
            Me.DialogResult = Windows.Forms.DialogResult.OK
            Me.AcceptButton = cmdOk
        End If
    End Sub

    Private Sub txtPhraseID_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhraseID.Leave
        If Me.txtPhraseID.Text <> "<Replace>" Then
            Me.myPhrase.PhraseID = Me.txtPhraseID.Text

        End If
    End Sub
End Class
