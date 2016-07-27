Imports System.Windows.Forms
Imports System.Drawing
Imports System.ComponentModel

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : OptionsListView
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' List view control for options 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [17/04/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class OptionsListView
    Inherits ListView

    Private myOptionClass As String
    Private myClientOptions As Intranet.intranet.customerns.SMClientOptionsCollection
    Public MYclientID As String
    Private blnRefresh As Boolean = True
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColStartDate As ColumnHeader
    Friend WithEvents ColEndDate As ColumnHeader
    Friend WithEvents ColAmmend As ColumnHeader
    Friend WithEvents ColAccess As ColumnHeader
    Private myAuthorityLevel As Integer = 0
    Private strUSerId As String = ""

    Public Event OptionAdded(ByVal Sender As OptionsListView, ByVal option_id As String)
    Public Event OptionRemoved(ByVal Sender As OptionsListView, ByVal Option_id As String)
    Public Event OptionChanged(ByVal Sender As OptionsListView, ByVal Option_id As String)

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Authority level of user 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AuthorityLevel() As Integer
        Get
            Return Me.myAuthorityLevel
        End Get
        Set(ByVal Value As Integer)
            myAuthorityLevel = Value
        End Set
    End Property

    Public Property UserID() As String
        Get
            Return Me.strUSerId
        End Get
        Set(ByVal Value As String)
            strUSerId = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fill the contents of the list view 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawOptionListView()
        blnRefresh = True
        Me.BeginUpdate()
        Me.Items.Clear()
        'Dim aSerrRow As MedscreenLib.Groups._OptionRow
        Dim strOptionList As String = MedscreenLib.CConnection.PackageStringList("Lib_cctool.GetOptionInClass", myOptionClass)
        MedscreenLib.Medscreen.LogTiming("  - Draw Options - Get option list")
        Dim strOptions As String() = strOptionList.Split(New Char() {","})
        Dim i As Integer

        For i = 0 To strOptions.Length - 1

            Dim optStr As String = strOptions.GetValue(i)
            If optStr.Trim.Length > 0 Then
                Dim oOpt As MedscreenLib.Glossary.Options = New MedscreenLib.Glossary.Options(optStr.Trim)
                'Check to see if option editable
                If oOpt.Editable Then
                    MedscreenLib.Medscreen.LogTiming("  - Draw Options - Get option ")
                    If Not oOpt Is Nothing Then

                        Dim ocOpt As Intranet.intranet.customerns.SmClientOption = myClientOptions.Item(optStr)
                        'MedscreenLib.Medscreen.LogTiming("  - Draw Options - Get Client option ")

                        Dim lvItem As LvOptItem = New LvOptItem(oOpt, ocOpt, Me.UserID)
                        'MedscreenLib.Medscreen.LogTiming("  - Draw Options - Create LV Item")

                        If Not Me.Enabled Then
                            lvItem.BackColor = Color.LightGray
                            lvItem.ForeColor = Color.Yellow

                        ElseIf oOpt.OptionType = "BOOL" Or oOpt.OptionType = "BLNVAL" Then
                            If Not ocOpt Is Nothing Then
                                If ocOpt.IsActiveOption(Now) Then   'Use normal colouring the option is active
                                    If ocOpt.optionBool = True Then
                                        lvItem.BackColor = Color.LightGreen
                                    Else
                                        lvItem.BackColor = Color.Cyan
                                    End If
                                Else
                                    lvItem.BackColor = Color.Red
                                    lvItem.ForeColor = Color.Yellow
                                End If
                            End If
                        ElseIf oOpt.OptionType = "VALUE" Or oOpt.OptionType = "RANGE" Then
                            If Not ocOpt Is Nothing AndAlso CStr(ocOpt.Value).Trim.Length > 0 Then
                                lvItem.BackColor = Color.LightGreen
                            End If
                        End If
                        Me.Items.Add(lvItem)
                        If Not oOpt.UserCanEdit(Me.UserID) Then
                            lvItem.BackColor = Color.LightGray
                        End If
                    End If
                End If
            End If                                      'Option editable 
        Next
        blnRefresh = False
        Me.EndUpdate()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' List of the client options provided is datasource
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <Bindable(True), Category("Data"), TypeConverter("System.Windows.Forms.Design.DataSourceConverter,System.Design")> _
    Public Property ClientOptions() As Intranet.intranet.customerns.SMClientOptionsCollection
        Set(ByVal Value As Intranet.intranet.customerns.SMClientOptionsCollection)
            myClientOptions = Value
        End Set
        Get
            Return myClientOptions
        End Get
    End Property

    <Bindable(True), Description("Class of option that will be displayed")> _
    Public Property OptionClass() As String
        Get
            Return Me.myOptionClass
        End Get
        Set(ByVal Value As String)
            myOptionClass = Value
        End Set
    End Property




    Public Sub New()
        MyBase.New()

        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader()
        ColStartDate = New ColumnHeader()
        Me.ColAccess = New ColumnHeader()
        Me.ColAmmend = New ColumnHeader()
        Me.ColEndDate = New ColumnHeader()

        '
        ' column start date 
        '
        ColStartDate.Text = "Start Date"
        ColStartDate.Width = 100
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Option Name"
        Me.ColumnHeader1.Width = 400
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Set"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Value"
        Me.ColumnHeader3.Width = 200

        '
        '  column end date 
        '
        Me.ColEndDate.Text = "End Date"
        '
        '
        '
        Me.ColAmmend.Text = "Amendment ID"
        '
        '
        '
        Me.ColAccess.Text = "Role Required"

        MyBase.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, ColStartDate, ColEndDate, ColAmmend, ColAccess})
        Me.FullRowSelect = True
        Me.View = View.Details

    End Sub

    Protected Overrides Sub OnDoubleClick(ByVal e As System.EventArgs)
        If Me.SelectedItems.Count = 0 Then Exit Sub ' None Selected Bog off 


        Dim myItem As LvOptItem = Me.SelectedItems.Item(0) 'Get item 
        If myItem Is Nothing Then Exit Sub 'Don't have one bog off

        Dim Copt As Intranet.intranet.customerns.SmClientOption

        If myItem.CustOption Is Nothing Then ' If we don't have an option create it and initialise

            If myItem._Option.OptionType = "BOOL" Then
                Copt = Me.myClientOptions.AddOption(Me.MYclientID, myItem._Option.OptionID, True, DBNull.Value)
            Else
                Copt = Me.myClientOptions.AddOption(Me.MYclientID, myItem._Option.OptionID, DBNull.Value, myItem._Option.DefaultValue)
            End If
            myItem.CustOption = Copt
        Else
            Copt = myItem.CustOption
        End If

        Dim aFrm As New frmCustomerOption()

        If myItem._Option.UserCanEdit(Me.UserID) AndAlso myItem._Option.Editable Then
            With aFrm
                .CustomerOption = myItem.CustOption
                Dim dlgR As DialogResult
                dlgR = aFrm.ShowDialog
                If dlgR = System.Windows.Forms.DialogResult.OK Then
                    Copt = .CustomerOption
                    Copt.[Operator] = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                    Copt.Update()
                    If Not Copt Is Nothing Then
                        RaiseEvent OptionAdded(Me, myItem._Option.OptionID)
                    End If
                End If
            End With
        Else
            MsgBox("You don't have the rights to edit this option", MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation)
        End If
    End Sub



    Private Sub OptionsListView_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.Delete And e.Control = True Then
            e.Handled = True
            If Me.SelectedItems.Count = 0 Then Exit Sub ' None Selected Bog off 

            Dim myItem As LvOptItem = Me.SelectedItems.Item(0) 'Get item 
            If myItem Is Nothing Then Exit Sub 'Don't have one bog off

            Dim Copt As Intranet.intranet.customerns.SmClientOption = myItem.CustOption
            If Not Copt Is Nothing Then
                Dim intPos As Integer
                If Me.ClientOptions.DeleteOption(Copt.OptionId) Then
                    RaiseEvent OptionRemoved(Me, Copt.OptionId)
                End If
            End If

        Else

        End If
    End Sub


End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : LvOptItem
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' List view Option Item
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class LvOptItem
    Inherits ListViewItem

    Private myCustopt As Intranet.intranet.customerns.SmClientOption
    Private myOpt As MedscreenLib.Glossary.Options
    Private myAuthorityLevel As Integer = 10
    Private myUser As String
    Private Const cstFalseString = "No"
    Private Const cstTrueString = "Yes"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new item for an option
    ''' </summary>
    ''' <param name="Opt">Option </param>
    ''' <param name="CustOpt">Customer Option if present</param>
    ''' <param name="AuthorityLevel">Authority Level</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Opt As MedscreenLib.Glossary.Options, _
        ByVal CustOpt As Intranet.intranet.customerns.SmClientOption, _
    Optional ByVal User As String = "")
        MyBase.new()
        Me.myUser = user
        myOpt = Opt
        myCustopt = CustOpt
        Refresh()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new instance 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert boolean to string 
    ''' </summary>
    ''' <param name="Input"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Function BoolString(ByVal Input As Boolean) As String
        If Input Then
            Return cstTrueString
        Else
            Return cstFalseString
        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refresh the list view item
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Refresh()
        If Not myOpt Is Nothing Then

            Me.Text = myOpt.OptionDescription
            Dim subItem As ListViewItem.ListViewSubItem
            If Not myCustopt Is Nothing Then
                If myOpt.UserCanEdit(Me.myUser) Then  'Set colours 
                    Me.BackColor = Color.LightGreen
                    Me.ForeColor = Color.Black
                Else
                    Me.BackColor = Color.PeachPuff
                    Me.ForeColor = Color.Black
                End If
                Me.Checked = True
                If myOpt.OptionType = MedscreenLib.Glossary.Glossary.Options.GCST_OPTION_BOOL Or _
                    myOpt.OptionType = MedscreenLib.Glossary.Glossary.Options.GCST_OPTION_BOOLVALUE Then
                    subItem = New ListViewItem.ListViewSubItem(Me, BoolString(myCustopt.optionBool))
                Else
                    subItem = New ListViewItem.ListViewSubItem(Me, "--")
                End If
                Me.SubItems.Add(subItem)
                If myOpt.OptionType = MedscreenLib.Glossary.Glossary.Options.GCST_OPTION_RANGE _
                    Or myOpt.OptionType = MedscreenLib.Glossary.Glossary.Options.GCST_OPTION_BOOLVALUE _
                    Or myOpt.OptionType = MedscreenLib.Glossary.Glossary.Options.GCST_OPTION_VALUE Then
                    subItem = New ListViewItem.ListViewSubItem(Me, myCustopt.Value)
                Else
                    subItem = New ListViewItem.ListViewSubItem(Me, "--")
                End If
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, myCustopt.StartDateText)
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, myCustopt.EndDateText)
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, myCustopt.ContractAmendment)
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, myOpt.RoleToEdit)
                Me.SubItems.Add(subItem)
            Else
                If myOpt.UserCanEdit(Me.myUser) Then 'Set colours 
                    Me.BackColor = Color.LightGreen
                    Me.ForeColor = Color.Black
                Else
                    Me.BackColor = Color.PeachPuff
                    Me.ForeColor = Color.Black
                End If
                subItem = New ListViewItem.ListViewSubItem(Me, "--")
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, "--")
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, "--")
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, "--")
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, "--")
                Me.SubItems.Add(subItem)
                subItem = New ListViewItem.ListViewSubItem(Me, myOpt.RoleToEdit)
                Me.SubItems.Add(subItem)
            End If

        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the option 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property _Option() As MedscreenLib.Glossary.Options
        Get
            Return Me.myOpt
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the customer option 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [23/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property CustOption() As Intranet.intranet.customerns.SmClientOption
        Get
            Return Me.myCustopt
        End Get
        Set(ByVal Value As Intranet.intranet.customerns.SmClientOption)
            myCustopt = Value
            Refresh()
        End Set
    End Property


End Class
