'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-11-29 09:29:21+00 $
'$Log: frmCustomerOptions.vb,v $
'Revision 1.0  2005-11-29 09:29:21+00  taylor
'Checked in after Commenting
'
'Revision 1.2  2005-06-17 07:03:45+01  taylor
'Milestone check in
'
'Revision 1.1  2004-05-04 10:08:30+01  taylor
'Added header check in
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>
'

Imports Intranet.intranet.customerns
Imports MedscreenLib.Glossary
Imports System.Windows.Forms

''' -----------------------------------------------------------------------------
''' Project	 : CCTool
''' Class	 : frmCustomerOptions
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Form to list all options/services for a customer
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class frmCustomerOptions
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        'AddHandler DataSet11.Customer_options.RowDeleted, New DataRowChangeEventHandler(AddressOf Row_Changing)

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
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents cbOptionList As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents CHNAme As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHBoolean As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHStartDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHEndDate As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents CHAmendment As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHOperator As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCustomerOptions))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdDelete = New System.Windows.Forms.Button
        Me.cbOptionList = New System.Windows.Forms.ComboBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.cmdOk = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.CHNAme = New System.Windows.Forms.ColumnHeader
        Me.CHBoolean = New System.Windows.Forms.ColumnHeader
        Me.CHValue = New System.Windows.Forms.ColumnHeader
        Me.CHStartDate = New System.Windows.Forms.ColumnHeader
        Me.CHEndDate = New System.Windows.Forms.ColumnHeader
        Me.CHAmendment = New System.Windows.Forms.ColumnHeader
        Me.CHOperator = New System.Windows.Forms.ColumnHeader
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmdDelete)
        Me.Panel1.Controls.Add(Me.cbOptionList)
        Me.Panel1.Controls.Add(Me.cmdAdd)
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.cmdOk)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 347)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(860, 64)
        Me.Panel1.TabIndex = 1
        '
        'cmdDelete
        '
        Me.cmdDelete.Image = CType(resources.GetObject("cmdDelete.Image"), System.Drawing.Image)
        Me.cmdDelete.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDelete.Location = New System.Drawing.Point(8, 37)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(88, 23)
        Me.cmdDelete.TabIndex = 10
        Me.cmdDelete.Text = "Del Option"
        Me.cmdDelete.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbOptionList
        '
        Me.cbOptionList.Location = New System.Drawing.Point(112, 10)
        Me.cbOptionList.Name = "cbOptionList"
        Me.cbOptionList.Size = New System.Drawing.Size(328, 21)
        Me.cbOptionList.Sorted = True
        Me.cbOptionList.TabIndex = 9
        Me.cbOptionList.Visible = False
        '
        'cmdAdd
        '
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(8, 7)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(88, 23)
        Me.cmdAdd.TabIndex = 8
        Me.cmdAdd.Text = "Add Option"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(772, 32)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 7
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Image)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(692, 32)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 6
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(260, 40)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(199, 21)
        Me.ComboBox1.TabIndex = 2
        Me.ComboBox1.Text = "ComboBox1"
        Me.ComboBox1.Visible = False
        '
        'CheckBox1
        '
        Me.CheckBox1.Location = New System.Drawing.Point(240, 80)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(104, 24)
        Me.CheckBox1.TabIndex = 4
        Me.CheckBox1.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(272, 72)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(100, 20)
        Me.TextBox1.TabIndex = 5
        Me.TextBox1.Text = "TextBox1"
        Me.TextBox1.Visible = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.ListView1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(860, 347)
        Me.Panel2.TabIndex = 6
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CHNAme, Me.CHBoolean, Me.CHValue, Me.CHStartDate, Me.CHEndDate, Me.CHAmendment, Me.CHOperator})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.Location = New System.Drawing.Point(0, 0)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(860, 347)
        Me.ListView1.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'CHNAme
        '
        Me.CHNAme.Text = "Option"
        Me.CHNAme.Width = 200
        '
        'CHBoolean
        '
        Me.CHBoolean.Text = "Boolean Value"
        '
        'CHValue
        '
        Me.CHValue.Text = "Value"
        Me.CHValue.Width = 300
        '
        'CHStartDate
        '
        Me.CHStartDate.Text = "Start Date"
        Me.CHStartDate.Width = 100
        '
        'CHEndDate
        '
        Me.CHEndDate.Text = "End Date"
        Me.CHEndDate.Width = 100
        '
        'CHAmendment
        '
        Me.CHAmendment.Text = "Amendment"
        Me.CHAmendment.Width = 80
        '
        'CHOperator
        '
        Me.CHOperator.Text = "Changed By"
        Me.CHOperator.Width = 100
        '
        'frmCustomerOptions
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(860, 411)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmCustomerOptions"
        Me.Text = "frmCustomerOptions"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region
    Private myClient As Intranet.intranet.customerns.Client
    Private iRow As Integer
    Private objPhraseColl As MedscreenLib.Glossary.PhraseCollection
    Private objPhrase As MedscreenLib.Glossary.Phrase
    Private objOpt As Intranet.intranet.customerns.SmClientOption
    Private WithEvents objService As Intranet.intranet.customerns.CustomerService
    Private objOption As MedscreenLib.Glossary.Options
    Private blnLoading As Boolean = True
    Private tmpCol As MedscreenLib.Glossary.OptionCollection


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible modes for form 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum COFormType
        ''' <summary>Option mode</summary>
        Options
        ''' <summary>Services mode</summary>
        Services
    End Enum

    Private myFormType As COFormType = COFormType.Options
    Public Event NotifyServiceChange(ByVal Service As String, ByVal Change As String)


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' What mode is the form running in 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property FormType() As COFormType
        Get
            Return Me.myFormType
        End Get
        Set(ByVal Value As COFormType)
            myFormType = Value
            If Me.FormType = COFormType.Options Then
                Me.CHBoolean.Width = 80
                Me.CHValue.Width = 80
                Me.Text = "Maintain customer Options"
                Me.cmdDelete.Text = "&Del Option"
                Me.cmdAdd.Text = "&Add Option"
            Else
                Me.CHBoolean.Width = 1
                Me.CHValue.Width = 1
                Me.Text = "Maintain Customer Services"
                Me.cmdDelete.Text = "&Del Service"
                Me.cmdAdd.Text = "&Add Service"
            End If

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Client whose options/services will be displayed
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Client() As Intranet.intranet.customerns.Client
        Get
            Return myClient
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.Client)
            Dim i As Integer
            Dim objOpt As Intranet.intranet.customerns.SmClientOption
            Dim objOption As MedscreenLib.Glossary.Options
            Dim objPhraseColl As MedscreenLib.Glossary.PhraseCollection
            Dim objPhrase As MedscreenLib.Glossary.Phrase
            'Dim Dr As DataSet1.Customer_optionsRow

            myClient = Value
            If myClient Is Nothing Then Exit Property

            'Okay we now have to set up the datagrid, 
            'this is done using a datatable filled with our data 
            'Me.DataSet11.Clear() 'First clear it out 


            'Me.DataSet11.AcceptChanges() 'Commit the changes (no writing to database)
            If Me.FormType = COFormType.Options Then
                FillOptionGrid()
                Me.Text = "Customer Options " & myClient.Identity & " - " & myClient.SMIDProfile
            Else
                Me.Text = "Customer Services " & myClient.Identity & " - " & myClient.SMIDProfile
                FillServiceGrid()
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the services display 
    ''' </summary>
    ''' <remarks>Should be changed to use Services Item 
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillServiceGrid()
        Dim I As Integer

        Dim LVItem As ListViewItems.lvServiceItem

        Me.ListView1.BeginUpdate()
        Me.ListView1.Items.Clear()
        For I = 0 To myClient.Services.Count - 1 'For each option that we have 
            Dim objServ As Intranet.intranet.customerns.CustomerService
            objServ = myClient.Services.Item(I) 'Get the customer option 
            If Not objServ Is Nothing AndAlso objServ.Removed = False Then
                LVItem = New ListViewItems.lvServiceItem(objServ)
                Me.ListView1.Items.Add(LVItem)
            End If

        Next
        Me.ListView1.EndUpdate()

    End Sub

    Private Sub SetupRange(ByVal LVItem As ListViewItem)


        Dim lvSubBool As ListViewItem.ListViewSubItem
        Dim lvSubValue As ListViewItem.ListViewSubItem

        objPhrase = Nothing
        objPhraseColl = MedscreenLib.Glossary.Glossary.PhraseList.Item(objOption.OptionID)
        If objPhraseColl Is Nothing Then
            objPhraseColl = New MedscreenLib.Glossary.PhraseCollection(objOption.OptionID)
            objPhraseColl.Load()
            MedscreenLib.Glossary.Glossary.PhraseList.Add(objPhraseColl) 'This is a collection of phrases
        End If
        If Not objPhraseColl Is Nothing Then
            objPhrase = objPhraseColl.Item(objOpt.Value) 'Get the phrase item
        End If
        If Not objPhrase Is Nothing Then
            lvSubValue = New ListViewItem.ListViewSubItem(LVItem, objPhrase.PhraseText)
            LVItem.SubItems.Add(lvSubValue)
            'We can use the phrase text 
            'Dr = Me.DataSet11.Customer_options.AddCustomer_optionsRow(objOption.OptionDescription, objOpt.optionBool, objPhrase.PhraseText, objOpt.OptionId)
        Else
            lvSubValue = New ListViewItem.ListViewSubItem(LVItem, objOpt.Value)
            LVItem.SubItems.Add(lvSubValue)
            'We'll have to make do with a value of the phrase element 
            'Dr = Me.DataSet11.Customer_options.AddCustomer_optionsRow(objOption.OptionDescription, objOpt.optionBool, objOpt.Value, objOpt.OptionId)
        End If

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the options 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub FillOptionGrid()
        Dim I As Integer

        Dim LVItem As ListViewItems.LvOptionItem
        Dim lvSubBool As ListViewItem.ListViewSubItem
        Dim lvSubValue As ListViewItem.ListViewSubItem

        Me.ListView1.BeginUpdate()
        Me.ListView1.Items.Clear()
        myClient.Options.Refresh()

        For I = 0 To myClient.Options.Count - 1 'For each option that we have 
            objOpt = myClient.Options.Item(I) 'Get the customer option 
            LVItem = New ListViewItems.LvOptionItem(objOpt)
            Me.ListView1.Items.Add(LVItem)
        Next
        Me.ListView1.EndUpdate()
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exit the form 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
        Dim i As Integer
        Dim objOpt1 As Intranet.intranet.customerns.SmClientOption

        For i = Me.myClient.Options.Count - 1 To 0 Step -1
            objOpt = Me.myClient.Options.Item(i)
            If objOpt.Deleted Then myClient.Options.DeleteOption(objOpt.OptionId)
            If objOpt.Updated Then objOpt.Update()
        Next

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an option for this customer
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Dim i As Integer
        Me.cbOptionList.Items.Clear()                  'Clear option 
        If Me.myFormType = COFormType.Options Then
            Dim objToption As MedscreenLib.Glossary.Options
            Dim unUsedOptionList As String = MedscreenLib.CConnection.PackageStringList("LIB_CCTool.GetUnusedOptions", myClient.Identity)
            If Not unUsedOptionList Is Nothing AndAlso unUsedOptionList.Trim.Length > 0 Then
                Dim unUsedOptions As String() = unUsedOptionList.Split(New Char() {","})
                Dim aString As String
                For Each aString In unUsedOptions
                    Me.cbOptionList.Items.Add(MedscreenLib.CConnection.PackageStringList("lib_customer.GetOptionDescription", aString))
                Next
            End If
            'For i = 0 To MedscreenLib.Glossary.Glossary.Options.Count - 1
            '    objToption = MedscreenLib.Glossary.Glossary.Options.Item(i)  'Get option and if not already used add it to combobox 
            '    If Not objToption.RemoveFlag AndAlso Client.Options.Item(objToption.OptionID) Is Nothing Then
            '        Me.cbOptionList.Items.Add(objToption.OptionDescription)
            '    End If
            'Next
        Else
            Dim objSer As MedscreenLib.Glossary.Phrase
            For Each objSer In MedscreenLib.Glossary.Glossary.ServicesSold
                Dim objCSer As Intranet.intranet.customerns.CustomerService
                objCSer = Client.Services.Item(objSer.PhraseID)
                If objCSer Is Nothing Then
                    If objSer.PhraseText.Trim.Length > 0 Then
                        Me.cbOptionList.Items.Add(objSer.PhraseText)
                    End If
                Else
                    If objSer.PhraseText.Trim.Length > 0 Then
                        If objCSer.Removed Then
                            Me.cbOptionList.Items.Add(objSer.PhraseText & " - Removed")

                        Else
                            Me.cbOptionList.Items.Add(objSer.PhraseText & " - Sold")

                        End If
                    End If
                End If
            Next
        End If
        Me.cbOptionList.Show()
    End Sub

    Private selected As Boolean = True
    Private Sub cbOptionList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbOptionList.KeyDown
        If e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
            selected = False
        ElseIf e.KeyCode = Keys.Space Then
            selected = True
            'e.Handled = True
        End If
    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOptionList.SelectedIndexChanged
        If selected Then
            If Me.myFormType = COFormType.Options Then
                Dim strOpt As String
                Dim objTOption As MedscreenLib.Glossary.Options
                'Dim dr As DataSet1.Customer_optionsRow

                strOpt = Me.cbOptionList.Items(Me.cbOptionList.SelectedIndex) 'Get the description
                Dim optID As String = MedscreenLib.CConnection.PackageStringList("lib_customer.getoptionbydescription", strOpt)
                objTOption = New MedscreenLib.Glossary.Options(optID)         'Identify option 

                If Not objTOption Is Nothing Then                       'If valid option then create an entry 
                    objOpt = myClient.Options.AddOption(myClient.Identity, objTOption.OptionID, True, "-")
                    If objOpt Is Nothing Then Exit Sub 'If we fail exit 

                    'Add new row into grid 
                    'dr = Me.DataSet11.Customer_options.AddCustomer_optionsRow(strOpt, objOpt.optionBool, objOpt.Value, objOpt.OptionId)

                    'Null out columns depending on option type 
                    If objTOption.OptionType = "BOOL" Then
                        objOpt.optionBool = (objTOption.DefaultValue = "T")
                    ElseIf objTOption.OptionType = "BLNVL" Then
                        objOpt.optionBool = True
                        objOpt.Value = objTOption.DefaultValue
                    Else
                        objOpt.Value = objTOption.DefaultValue
                        'dr.Set_BooleanNull()
                    End If
                    'Me.DataSet11.Customer_options.AcceptChanges()
                End If
                objOpt.Update()
                Me.FillOptionGrid()
            Else
                Dim strOpt As String
                Dim strChange As String = ""                            'Record the change made if any
                strOpt = Me.cbOptionList.Items(Me.cbOptionList.SelectedIndex) 'Get the description 
                Dim objPhrase As MedscreenLib.Glossary.Phrase
                For Each objPhrase In MedscreenLib.Glossary.Glossary.ServicesSold
                    If InStr(strOpt, objPhrase.PhraseText) > 0 AndAlso objPhrase.PhraseText.Trim.Length > 0 Then
                        Exit For
                    End If
                    objPhrase = Nothing
                Next
                If Not objPhrase Is Nothing Then                        'We have a service 
                    Dim objSer As Intranet.intranet.customerns.CustomerService
                    objSer = myClient.Services.Item(objPhrase.PhraseID) 'Get the service
                    If objSer Is Nothing Then                           'Customer never had the service
                        objSer = New Intranet.intranet.customerns.CustomerService()
                        objSer.CustomerId = Me.myClient.Identity
                        objSer.Service = objPhrase.PhraseID
                        strChange = "Added"
                    Else
                        objSer.Removed = False
                        strChange = "Reinstated"
                    End If
                    objSer.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    Me.myClient.Services.Add(objSer)
                    Dim frmSer As New frmCustomerService()
                    frmSer.CustomerService = objSer
                    If frmSer.ShowDialog = DialogResult.OK Then
                        objSer = frmSer.CustomerService
                        objSer.ModifiedOn = Now
                        objSer.update()
                        objSer.NotifyChange(strChange) 'Okay making the change need to notify caller
                        RaiseEvent NotifyServiceChange(objSer.Service, strChange)

                    End If
                    Me.FillServiceGrid()
                End If
            End If
        End If
        selected = True
    End Sub

    'Value (boolean value) data retreival 



    Private Sub ListView1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Dim lvItem As ListViewItem

        'If Me.myFormType = COFormType.Options Then

        If Me.ListView1.SelectedItems.Count = 1 Then                            'Have we selected an item
            lvItem = Me.ListView1.SelectedItems.Item(0)
            If TypeOf lvItem Is ListViewItems.LvOptionItem Then                 'Is it an option item
                Dim lvOption As ListViewItems.LvOptionItem = lvItem
                'Me.objOpt = Me.myClient.Options.Item(lvItem.Tag)
                If Not lvOption._Option Is Nothing Then                         'Does it have an option?
                    Me.objOption = lvOption._Option                             'Set the value
                    Me.objOpt = lvOption.CustOption
                End If
            ElseIf TypeOf lvItem Is ListViewItems.lvServiceItem Then            'Is it a service item?
                Dim lvService As ListViewItems.lvServiceItem = lvItem
                If Not lvService.Service Is Nothing Then                        'Does it have a service?
                    Me.objService = lvService.Service
                End If
            End If
        End If
        ' Else
        'If Me.ListView1.SelectedItems.Count = 1 Then
        '    lvItem = Me.ListView1.SelectedItems.Item(0)
        '    Me.objService = Me.myClient.Services.Item(lvItem.Tag)
        'End If
        'End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Edit the item 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick

        If Me.myFormType = COFormType.Options Then
            Dim frmCO As frmCustomerOption = New frmCustomerOption()

            If objOpt Is Nothing Then Exit Sub
            If objOption Is Nothing Then Exit Sub
            If objOption.RemoveFlag Then
                MsgBox("This option is no longer Active and so editing it would be pointless", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            If Not objOption.Editable Then
                MsgBox("This option (" & objOption.OptionDescription & ") is not available for editing", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
                Exit Sub

            End If

            'Check user has sufficient rights to edit 
            'If CInt(MedscreenLib.Glossary.Glossary.CurrentSMUser.OptionAuthority) >= CInt(objOption.OptionAuthority) Then
            With frmCO
                .CustomerOption = Me.objOpt
                If .ShowDialog = DialogResult.OK Then
                    Me.objOpt = .CustomerOption
                    Me.objOpt.Updated = True
                    Me.objOpt.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    objOpt.Update()
                End If
            End With
            Me.FillOptionGrid()
            'Else
            'MsgBox("Please contact support, you don't have sufficient rights to modify this option.", MsgBoxStyle.Information Or MsgBoxStyle.OKOnly)
            'End If
        Else
            If Me.objService Is Nothing Then Exit Sub

            Dim frmser As New frmCustomerService()
            frmser.CustomerService = objService

            If frmser.ShowDialog = DialogResult.OK Then
                objService = frmser.CustomerService
                objService.update()
                Me.FillServiceGrid()
            End If
        End If
    End Sub


    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Delete service or option 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        If Me.myFormType = COFormType.Options Then
            If objOpt Is Nothing Then Exit Sub

            Me.objOpt.Deleted = True
            FillOptionGrid()
        Else
            If Me.objService Is Nothing Then Exit Sub
            Me.objService.Removed = True
            Me.objService.update()
            objService.NotifyChange("Removed")
            RaiseEvent NotifyServiceChange(objService.Service, "Removed")
            FillServiceGrid()
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Cancel the form
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Dim i As Integer
        Dim objOpt1 As Intranet.intranet.customerns.SmClientOption

        For i = Me.myClient.Options.Count - 1 To 0 Step -1
            objOpt = Me.myClient.Options.Item(i)
            objOpt.RollBack()
        Next

    End Sub
End Class
