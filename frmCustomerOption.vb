'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 08:50:55+00 $
'$Log: frmCustomerOption.vb,v $
'Revision 1.0  2005-11-29 08:50:55+00  taylor
'Commented
'
Imports MedscreenLib
Imports MedscreenLib.Medscreen
Imports Intranet.intranet.glossary
Imports MedscreenLib.Glossary
Imports System.Windows.Forms

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCustomerOption
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form for editing customer options
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [22/10/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCustomerOption
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
    Friend WithEvents lblOptionID As System.Windows.Forms.Label
    Friend WithEvents CKBool As System.Windows.Forms.CheckBox
    Friend WithEvents lblValue As System.Windows.Forms.Label
    Friend WithEvents txtValue As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents dtStartDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents DtEndDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents cbPhrase As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdNullEndDate As System.Windows.Forms.Button
    Friend WithEvents cklbMulti As System.Windows.Forms.CheckedListBox
    Friend WithEvents txtContractAmmend As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerOption))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.cmdNullEndDate = New System.Windows.Forms.Button
        Me.txtContractAmmend = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cbPhrase = New System.Windows.Forms.ComboBox
        Me.DtEndDate = New System.Windows.Forms.DateTimePicker
        Me.dtStartDate = New System.Windows.Forms.DateTimePicker
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtValue = New System.Windows.Forms.TextBox
        Me.lblValue = New System.Windows.Forms.Label
        Me.CKBool = New System.Windows.Forms.CheckBox
        Me.lblOptionID = New System.Windows.Forms.Label
        Me.cklbMulti = New System.Windows.Forms.CheckedListBox
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 233)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(504, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(420, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 9
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(340, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 8
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.cklbMulti)
        Me.Panel2.Controls.Add(Me.cmdNullEndDate)
        Me.Panel2.Controls.Add(Me.txtContractAmmend)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.cbPhrase)
        Me.Panel2.Controls.Add(Me.DtEndDate)
        Me.Panel2.Controls.Add(Me.dtStartDate)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.txtValue)
        Me.Panel2.Controls.Add(Me.lblValue)
        Me.Panel2.Controls.Add(Me.CKBool)
        Me.Panel2.Controls.Add(Me.lblOptionID)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(504, 233)
        Me.Panel2.TabIndex = 1
        '
        'cmdNullEndDate
        '
        Me.cmdNullEndDate.Image = CType(resources.GetObject("cmdNullEndDate.Image"), System.Drawing.Image)
        Me.cmdNullEndDate.Location = New System.Drawing.Point(231, 160)
        Me.cmdNullEndDate.Name = "cmdNullEndDate"
        Me.cmdNullEndDate.Size = New System.Drawing.Size(23, 23)
        Me.cmdNullEndDate.TabIndex = 11
        Me.cmdNullEndDate.UseVisualStyleBackColor = True
        '
        'txtContractAmmend
        '
        Me.txtContractAmmend.Location = New System.Drawing.Point(136, 184)
        Me.txtContractAmmend.Name = "txtContractAmmend"
        Me.txtContractAmmend.Size = New System.Drawing.Size(208, 20)
        Me.txtContractAmmend.TabIndex = 10
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 189)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(120, 23)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Contract Amendment"
        '
        'cbPhrase
        '
        Me.cbPhrase.Location = New System.Drawing.Point(88, 104)
        Me.cbPhrase.Name = "cbPhrase"
        Me.cbPhrase.Size = New System.Drawing.Size(384, 21)
        Me.cbPhrase.TabIndex = 8
        '
        'DtEndDate
        '
        Me.DtEndDate.Location = New System.Drawing.Point(88, 160)
        Me.DtEndDate.Name = "DtEndDate"
        Me.DtEndDate.ShowCheckBox = True
        Me.DtEndDate.Size = New System.Drawing.Size(136, 20)
        Me.DtEndDate.TabIndex = 7
        '
        'dtStartDate
        '
        Me.dtStartDate.Location = New System.Drawing.Point(88, 136)
        Me.dtStartDate.Name = "dtStartDate"
        Me.dtStartDate.ShowCheckBox = True
        Me.dtStartDate.Size = New System.Drawing.Size(136, 20)
        Me.dtStartDate.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(14, 160)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 23)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "End Date"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(56, 23)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Start Date"
        '
        'txtValue
        '
        Me.txtValue.Location = New System.Drawing.Point(88, 104)
        Me.txtValue.Name = "txtValue"
        Me.txtValue.Size = New System.Drawing.Size(104, 20)
        Me.txtValue.TabIndex = 3
        '
        'lblValue
        '
        Me.lblValue.Location = New System.Drawing.Point(40, 104)
        Me.lblValue.Name = "lblValue"
        Me.lblValue.Size = New System.Drawing.Size(48, 23)
        Me.lblValue.TabIndex = 2
        Me.lblValue.Text = "Value"
        '
        'CKBool
        '
        Me.CKBool.Location = New System.Drawing.Point(40, 64)
        Me.CKBool.Name = "CKBool"
        Me.CKBool.Size = New System.Drawing.Size(360, 24)
        Me.CKBool.TabIndex = 1
        '
        'lblOptionID
        '
        Me.lblOptionID.Location = New System.Drawing.Point(40, 16)
        Me.lblOptionID.Name = "lblOptionID"
        Me.lblOptionID.Size = New System.Drawing.Size(360, 23)
        Me.lblOptionID.TabIndex = 0
        '
        'cklbMulti
        '
        Me.cklbMulti.CheckOnClick = True
        Me.cklbMulti.FormattingEnabled = True
        Me.cklbMulti.Location = New System.Drawing.Point(88, 16)
        Me.cklbMulti.Name = "cklbMulti"
        Me.cklbMulti.Size = New System.Drawing.Size(384, 109)
        Me.cklbMulti.TabIndex = 12
        '
        'frmCustomerOption
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 273)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmCustomerOption"
        Me.Text = "Customer Option"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private objOpt As Intranet.intranet.customerns.SmClientOption
    Private objOption As MedscreenLib.Glossary.Options
    Private objPhraseColl As MedscreenLib.Glossary.PhraseCollection
    '<Added code modified 26-Apr-2007 07:05 by LOGOS\Taylor> 
    Public Enum FormType
        CustomerOption
        CustomerDiscount
    End Enum
    Private myFormType As FormType = FormType.CustomerOption
    '</Added code modified 26-Apr-2007 07:05 by LOGOS\Taylor> 

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sets the way the form will be displayed 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Type() As FormType
        Get
            Return Me.myFormType
        End Get
        Set(ByVal Value As FormType)
            myFormType = Value
            'If myFormType Is Nothing Then Exit Property
            If myFormType = FormType.CustomerDiscount Then
                Me.cbPhrase.Hide()
                Me.CKBool.Hide()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' set start date 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property StartDate() As Date
        Get
            Return Me.dtStartDate.Value()
        End Get
        Set(ByVal Value As Date)
            Me.dtStartDate.Value() = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the end date 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property EndDate() As Date
        Get
            Return Me.DtEndDate.Value
        End Get
        Set(ByVal Value As Date)
            Me.DtEndDate.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Contract ammendment 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/04/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ContractAmmedment() As String
        Get
            Return Me.txtContractAmmend.Text
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Supplied customer option 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CustomerOption() As Intranet.intranet.customerns.SmClientOption
        Get
            If Not objOpt Is Nothing Then
                Select Case objOption.OptionType        'Action based on Option Type
                    Case OptionCollection.GCST_OPTION_RANGE, OptionCollection.GCST_OPTION_TABLE       'Range should have phrase 
                        If Me.cbPhrase.SelectedIndex > -1 Then
                            objOpt.Value = Me.objPhraseColl.Item(Me.cbPhrase.SelectedIndex).PhraseID
                        End If
                    Case OptionCollection.GCST_OPTION_BOOL              'Boolean 
                        objOpt.optionBool = Me.CKBool.Checked
                    Case OptionCollection.GCST_OPTION_VALUE             'Value
                        objOpt.Value = Me.txtValue.Text
                    Case OptionCollection.GCST_OPTION_BOOLVALUE         'Boolean and Value
                        objOpt.optionBool = Me.CKBool.Checked
                        objOpt.Value = Me.txtValue.Text
                    Case OptionCollection.GCST_OPTION_BOOLRANGE         'Boolean and Range 
                        If Me.cbPhrase.SelectedIndex > -1 Then
                            objOpt.Value = Me.objPhraseColl.Item(Me.cbPhrase.SelectedIndex).PhraseID
                        End If
                        objOpt.optionBool = Me.CKBool.Checked
                    Case OptionCollection.GCST_OPTION_MULTI             'Multi Select items
                        Dim strStatus As String = ""
                        If cklbMulti.CheckedItems.Count <> 0 Then
                            ' If so, loop through all checked items and print results.
                            Dim x As Integer
                            Dim s As String = ""
                            For x = 0 To cklbMulti.CheckedItems.Count - 1
                                s = cklbMulti.CheckedItems(x).ToString
                                If strStatus.Trim.Length > 0 Then strStatus += ";"
                                strStatus += Mid(s, 1, InStr(s, " -")).Trim
                            Next x
                            objOpt.Value = strStatus
                            'MessageBox.Show(strStatus)
                        End If
                End Select

                'See if we can expire the option
                DtEndDate.Enabled = objOption.Expireable
                If objOption.Expireable Then

                    If Not Me.DtEndDate.Checked Then
                        'objOpt.optionEndDate = DateField.ZeroDate
                        objOpt.SetEndDateNull()
                    Else
                        objOpt.optionEndDate = Me.DtEndDate.Value
                    End If
                Else            'Not expireable
                    objOpt.SetEndDateNull()
                End If
                If Not Me.dtStartDate.Checked Then
                    objOpt.optionStartDate = DateField.ZeroDate
                Else
                    objOpt.optionStartDate = Me.dtStartDate.Value
                End If
                objOpt.ContractAmendment = Me.txtContractAmmend.Text
            End If
            Return objOpt
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.SmClientOption)
            objOpt = Value

            If objOpt Is Nothing Then Exit Property

            objOption = MedscreenLib.Glossary.Glossary.Options.Item(objOpt.OptType)
            If objOption Is Nothing Then Exit Property

            'Ensure all labels are visible 
            Me.lblOptionID.Show()
            Me.Text = "Customer Option - " & objOption.OptionDescription
            Me.lblOptionID.Text = objOption.OptionDescription
            Me.lblValue.Show()
            Me.cklbMulti.Hide()

            Select Case objOption.OptionType
                Case OptionCollection.GCST_OPTION_RANGE, OptionCollection.GCST_OPTION_TABLE, OptionCollection.GCST_OPTION_MULTI
                    Me.cbPhrase.Show()
                    Me.cbPhrase.SelectedIndex = -1
                    Me.txtValue.Hide()
                    Me.CKBool.Hide()

                    objPhraseColl = objOption.Phrases 'MedscreenLib.Glossary.Glossary.PhraseList.Item(objOption.OptionI)
                    If objPhraseColl Is Nothing Then
                        If objOption.OptionType = OptionCollection.GCST_OPTION_TABLE Then
                            objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(objOption.Reference, MedscreenLib.Glossary.PhraseCollection.BuildBy.Table)
                        ElseIf objOption.OptionType = OptionCollection.GCST_OPTION_MULTI Then
                            objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(objOption.Reference)
                            objPhraseColl.Load()
                        Else
                            objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(objOption.Reference)
                            objPhraseColl.Load()
                        End If

                        MedscreenLib.Glossary.Glossary.PhraseList.Add(objPhraseColl) 'This is a collection of phrases
                    End If
                    If Not objPhraseColl Is Nothing Then
                        Me.cbPhrase.DataSource = objPhraseColl
                        Me.cbPhrase.DisplayMember = "PhraseText"
                        Me.cbPhrase.ValueMember = "PhraseID"
                        Dim objPhrase As MedscreenLib.Glossary.Phrase = objPhraseColl.Item(objOpt.Value)
                        If Not objPhrase Is Nothing Then

                            Dim intOpt As Integer = objPhraseColl.IndexOf(objOpt.Value)
                            Me.cbPhrase.SelectedIndex = intOpt
                        End If
                    End If
                    If objOption.OptionType = OptionCollection.GCST_OPTION_MULTI Then
                        Me.cbPhrase.Hide()
                        Me.cklbMulti.Show()
                        Me.CKBool.Hide()
                        Me.CKBool.Hide()
                        Me.lblValue.Hide()
                        Me.lblOptionID.Hide()
                        For Each phr As Phrase In objPhraseColl
                            If phr.PhraseID.Trim.Length > 0 Then  'skip blank ones added to list 
                                If InStr(objOpt.Value, phr.PhraseID) <> 0 Then
                                    Me.cklbMulti.Items.Add(phr.PhraseID & " - " & phr.PhraseText, True)
                                Else
                                    Me.cklbMulti.Items.Add(phr.PhraseID & " - " & phr.PhraseText, False)
                                End If
                            End If
                        Next
                    End If
                    ' Case OptionCollection.GCST_OPTION_TABLE

                    'Me.cklbMulti.Hide()
                Case OptionCollection.GCST_OPTION_BOOL
                    Me.cbPhrase.Hide()
                    Me.txtValue.Hide()
                    Me.CKBool.Show()
                    Me.CKBool.Text = objOption.OptionDescription
                    Me.CKBool.Checked = objOpt.optionBool
                    Me.lblValue.Hide()
                Case OptionCollection.GCST_OPTION_VALUE
                    Me.cbPhrase.Hide()
                    Me.txtValue.Show()
                    Me.CKBool.Hide()
                    Me.txtValue.Text = objOpt.Value
                Case OptionCollection.GCST_OPTION_BOOLVALUE
                    Me.cbPhrase.Hide()
                    Me.txtValue.Show()
                    Me.lblValue.Show()
                    Me.CKBool.Show()
                    Me.CKBool.Text = objOption.OptionDescription
                    Me.CKBool.Checked = objOpt.optionBool
                    Me.txtValue.Text = objOpt.Value
                Case OptionCollection.GCST_OPTION_BOOLRANGE
                    Me.cbPhrase.Show()
                    Me.cbPhrase.SelectedIndex = -1
                    Me.txtValue.Hide()
                    Me.CKBool.Show()
                    Me.CKBool.Text = objOption.OptionDescription
                    Me.CKBool.Checked = objOpt.optionBool

                    objPhraseColl = MedscreenLib.Glossary.Glossary.PhraseList.Item(objOption.OptionID)
                    If objPhraseColl Is Nothing Then
                        objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(objOption.OptionID)
                        objPhraseColl.Load()
                        MedscreenLib.Glossary.Glossary.PhraseList.Add(objPhraseColl) 'This is a collection of phrases
                    End If
                    If Not objPhraseColl Is Nothing Then
                        Me.cbPhrase.DataSource = objPhraseColl
                        Me.cbPhrase.DisplayMember = "PhraseText"
                        Me.cbPhrase.ValueMember = "PhraseID"
                        Dim objPhrase As MedscreenLib.Glossary.Phrase = objPhraseColl.Item(objOpt.Value)
                        If Not objPhrase Is Nothing Then

                            Dim intOpt As Integer = objPhraseColl.IndexOf(objOpt.Value)
                            Me.cbPhrase.SelectedIndex = intOpt
                        End If
                    End If
                Case OptionCollection.GCST_OPTION_MULTI
                    Me.cbPhrase.Hide()
                    Me.cbPhrase.SelectedIndex = -1
                    Me.txtValue.Hide()
                    Me.CKBool.Hide()
                    For Each phr As Phrase In objPhraseColl
                        If phr.PhraseID.Trim.Length > 0 Then  'skip blank ones added to list 
                            If InStr(objOpt.Value, phr.PhraseID) <> 0 Then
                                Me.cklbMulti.Items.Add(phr.PhraseID & " - " & phr.PhraseText, True)
                            Else
                                Me.cklbMulti.Items.Add(phr.PhraseID & " - " & phr.PhraseText, False)
                            End If
                        End If
                    Next
            End Select
            If objOpt.optionEndDate <> MedscreenLib.DateField.ZeroDate Then
                Me.DtEndDate.Value = objOpt.optionEndDate
                Me.DtEndDate.Checked = True
            Else
                'Me.dtStartDate.Value = Now
                Me.DtEndDate.Checked = False

            End If
            If objOpt.optionStartDate <> MedscreenLib.DateField.ZeroDate Then
                Me.dtStartDate.Value = objOpt.optionStartDate
                Me.dtStartDate.Checked = True
            Else
                'Me.dtStartDate.Value = Now
                Me.dtStartDate.Checked = False


            End If
            Me.txtContractAmmend.Text = ""
        End Set
    End Property
    ''' <developer>Taylor</developer>
    ''' <summary>
    ''' Allow nulling of end date
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><Author>TAYLOR</Author><created>24-Feb-2011</created></revision></revisionHistory>
    Private Sub cmdNullEndDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNullEndDate.Click
        If objOpt IsNot Nothing Then
            objOpt.SetEndDateNull()
            DtEndDate.Value = DateField.ZeroDate
            Me.DtEndDate.Checked = False
        Else
            DtEndDate.Value = DateField.ZeroDate
            Me.DtEndDate.Checked = False
        End If
    End Sub
End Class
