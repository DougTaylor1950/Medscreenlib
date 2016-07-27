Imports System.Windows.Forms
Imports MedscreenLib
Imports MedscreenLib.Glossary.Glossary
Imports MedscreenLib.CConnection

'''<summary>
''' Form to enable users to find an address
''' </summary>
''' <remarks>
''' Generic form at the moment mainly use din CCTool to find an address.
''' </remarks>
Public Class FindAddress
    Inherits System.Windows.Forms.Form
    Private strQuery As String = ""
    Private MyAddresses As MedscreenLib.Address.AddressCollection
    Private myAddressId As Integer
    Private FillingForm As Boolean = True

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.cbCountries.DataSource = MedscreenLib.Glossary.Glossary.Countries
        Me.cbCountries.ValueMember = "CountryId"
        Me.cbCountries.DisplayMember = "CountryName"

        'Add any initialization after the InitializeComponent() call
        Dim I As Integer

        For I = 0 To 10
            ColumnDirection(I) = 1
        Next

        'LoadAddresses()
        FillingForm = False


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
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents DataSet1 As DataSet
    'Friend WithEvents Addresses1 As Intranet.Addresses
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPostCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents txtLine1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtDistict As System.Windows.Forms.TextBox
    Friend WithEvents txtContact As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdOk As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cbCountries As System.Windows.Forms.ComboBox
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents CHAddrId As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHLine1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHLine2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHLine4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHCity As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHDistrict As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHPostcode As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHContact As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHEmail As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHPhone As System.Windows.Forms.ColumnHeader
    Friend WithEvents CHFax As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CmdFind As System.Windows.Forms.Button
    Friend WithEvents CHLine3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtAddressID As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FindAddress))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdOk = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.CmdFind = New System.Windows.Forms.Button()
        Me.cbCountries = New System.Windows.Forms.ComboBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtContact = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtDistict = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtLine1 = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtCity = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtPostCode = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.CHAddrId = New System.Windows.Forms.ColumnHeader()
        Me.CHLine1 = New System.Windows.Forms.ColumnHeader()
        Me.CHLine2 = New System.Windows.Forms.ColumnHeader()
        Me.CHLine3 = New System.Windows.Forms.ColumnHeader()
        Me.CHLine4 = New System.Windows.Forms.ColumnHeader()
        Me.CHCity = New System.Windows.Forms.ColumnHeader()
        Me.CHDistrict = New System.Windows.Forms.ColumnHeader()
        Me.CHPostcode = New System.Windows.Forms.ColumnHeader()
        Me.CHContact = New System.Windows.Forms.ColumnHeader()
        Me.CHEmail = New System.Windows.Forms.ColumnHeader()
        Me.CHPhone = New System.Windows.Forms.ColumnHeader()
        Me.CHFax = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.DataSet1 = New System.Data.DataSet()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtAddressID = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.cmdCancel, Me.cmdOk})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 357)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(856, 40)
        Me.Panel1.TabIndex = 0
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Bitmap)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(768, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(72, 24)
        Me.cmdCancel.TabIndex = 13
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmdOk
        '
        Me.cmdOk.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdOk.Image = CType(resources.GetObject("cmdOk.Image"), System.Drawing.Bitmap)
        Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdOk.Location = New System.Drawing.Point(688, 8)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(72, 24)
        Me.cmdOk.TabIndex = 12
        Me.cmdOk.Text = "Ok"
        Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Panel2
        '
        Me.Panel2.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtAddressID, Me.Label8, Me.CmdFind, Me.cbCountries, Me.Label7, Me.Label6, Me.txtContact, Me.Label5, Me.txtDistict, Me.Label4, Me.txtLine1, Me.Label3, Me.txtCity, Me.Label2, Me.txtPostCode, Me.Label1})
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(856, 100)
        Me.Panel2.TabIndex = 1
        '
        'CmdFind
        '
        Me.CmdFind.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.CmdFind.Image = CType(resources.GetObject("CmdFind.Image"), System.Drawing.Bitmap)
        Me.CmdFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.CmdFind.Location = New System.Drawing.Point(728, 56)
        Me.CmdFind.Name = "CmdFind"
        Me.CmdFind.Size = New System.Drawing.Size(72, 24)
        Me.CmdFind.TabIndex = 13
        Me.CmdFind.Text = "Find"
        Me.CmdFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cbCountries
        '
        Me.cbCountries.Location = New System.Drawing.Point(520, 39)
        Me.cbCountries.Name = "cbCountries"
        Me.cbCountries.Size = New System.Drawing.Size(121, 21)
        Me.cbCountries.TabIndex = 12
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(464, 41)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 23)
        Me.Label7.TabIndex = 11
        Me.Label7.Text = "Country"
        Me.Label7.Visible = False
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, (System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.HotTrack
        Me.Label6.Location = New System.Drawing.Point(8, 72)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(656, 23)
        Me.Label6.TabIndex = 10
        Me.Label6.Text = "All searches are case insensitive."
        '
        'txtContact
        '
        Me.txtContact.Location = New System.Drawing.Point(264, 40)
        Me.txtContact.Name = "txtContact"
        Me.txtContact.Size = New System.Drawing.Size(168, 20)
        Me.txtContact.TabIndex = 9
        Me.txtContact.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtContact, "Search in contact field")
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(200, 40)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(56, 23)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "Contact"
        '
        'txtDistict
        '
        Me.txtDistict.Location = New System.Drawing.Point(403, 5)
        Me.txtDistict.Name = "txtDistict"
        Me.txtDistict.Size = New System.Drawing.Size(96, 20)
        Me.txtDistict.TabIndex = 7
        Me.txtDistict.Text = ""
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(352, 5)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(40, 23)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "District"
        '
        'txtLine1
        '
        Me.txtLine1.Location = New System.Drawing.Point(72, 40)
        Me.txtLine1.Name = "txtLine1"
        Me.txtLine1.Size = New System.Drawing.Size(96, 20)
        Me.txtLine1.TabIndex = 5
        Me.txtLine1.Text = ""
        Me.ToolTip1.SetToolTip(Me.txtLine1, "Search for value in lines 1, 2 or 3 of the address")
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(24, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Lines"
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(240, 4)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(96, 20)
        Me.txtCity.TabIndex = 3
        Me.txtCity.Text = ""
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(200, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "City"
        '
        'txtPostCode
        '
        Me.txtPostCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtPostCode.Location = New System.Drawing.Point(104, 4)
        Me.txtPostCode.Name = "txtPostCode"
        Me.txtPostCode.Size = New System.Drawing.Size(56, 20)
        Me.txtPostCode.TabIndex = 1
        Me.txtPostCode.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Post Code :"
        '
        'Panel3
        '
        Me.Panel3.Controls.AddRange(New System.Windows.Forms.Control() {Me.ListView1})
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 100)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(856, 257)
        Me.Panel3.TabIndex = 2
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.CHAddrId, Me.CHLine1, Me.CHLine2, Me.CHLine3, Me.CHLine4, Me.CHCity, Me.CHDistrict, Me.CHPostcode, Me.CHContact, Me.CHEmail, Me.CHPhone, Me.CHFax, Me.ColumnHeader1})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.FullRowSelect = True
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(856, 257)
        Me.ListView1.TabIndex = 0
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'CHAddrId
        '
        Me.CHAddrId.Text = "Address Id"
        '
        'CHLine1
        '
        Me.CHLine1.Text = "Line 1"
        Me.CHLine1.Width = 120
        '
        'CHLine2
        '
        Me.CHLine2.Text = "Line 2"
        Me.CHLine2.Width = 90
        '
        'CHLine3
        '
        Me.CHLine3.Text = "Line 3"
        '
        'CHLine4
        '
        Me.CHLine4.Text = "Line 4"
        Me.CHLine4.Width = 90
        '
        'CHCity
        '
        Me.CHCity.Text = "City"
        '
        'CHDistrict
        '
        Me.CHDistrict.Text = "District"
        '
        'CHPostcode
        '
        Me.CHPostcode.Text = "Post Code"
        '
        'CHContact
        '
        Me.CHContact.Text = "Contact"
        '
        'CHEmail
        '
        Me.CHEmail.Text = "Email"
        Me.CHEmail.Width = 90
        '
        'CHPhone
        '
        Me.CHPhone.Text = "Phone"
        '
        'CHFax
        '
        Me.CHFax.Text = "Fax"
        Me.CHFax.Width = 90
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Country"
        '
        'DataSet1
        '
        Me.DataSet1.DataSetName = "Addresses"
        Me.DataSet1.EnforceConstraints = False
        Me.DataSet1.Locale = New System.Globalization.CultureInfo("en-GB")
        Me.DataSet1.Namespace = "http://tempuri.org/tempaddr.xsd"
        '
        'txtAddressID
        '
        Me.txtAddressID.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper
        Me.txtAddressID.Location = New System.Drawing.Point(688, 8)
        Me.txtAddressID.Name = "txtAddressID"
        Me.txtAddressID.Size = New System.Drawing.Size(56, 20)
        Me.txtAddressID.TabIndex = 15
        Me.txtAddressID.Text = ""
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(600, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 23)
        Me.Label8.TabIndex = 14
        Me.Label8.Text = "Address ID :"
        '
        'FindAddress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(856, 397)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel3, Me.Panel2, Me.Panel1})
        Me.Name = "FindAddress"
        Me.Text = "FindAddress"
        Me.Panel1.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        CType(Me.DataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    'Friend WithEvents AddressData As Intranet.Addresses
    Private ColumnDirection(10) As Integer


    Private Sub LoadAddresses()
        If SMAddresses.Count = 0 Then
            SMAddresses.Load()
        End If
        MyAddresses = SMAddresses
        'Me.Addresses1 = MyAddresses.ToDataSet
        'Me.DGAddress.DataSource = Me.Addresses1.Address
    End Sub

    Private Sub txtPostCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPostCode.TextChanged
        SetUpQuery()
    End Sub

    Private Sub SetUpQuery()
        Dim blnWhere As Boolean = False
        strQuery = "Select Address_id from address "
        If Me.txtAddressID.Text.Trim.Length > 0 Then
            strQuery += " where address_id = " & Me.txtAddressID.Text
        Else
            If txtPostCode.Text.Trim.Length > 0 Then
                If Not blnWhere Then
                    strQuery += " where postcode like '" & txtPostCode.Text.Trim.ToUpper & "%'"
                    blnWhere = True
                Else
                    strQuery += " and postcode like '" & txtPostCode.Text.Trim.ToUpper & "%'"
                End If
            End If

            If txtCity.Text.Trim.Length > 0 Then
                If Not blnWhere Then
                    strQuery += " where upper(City) like '" & txtCity.Text.Trim.ToUpper & "%'"
                    blnWhere = True
                Else
                    strQuery += " and upper(City) like '" & txtCity.Text.Trim.ToUpper & "%'"
                End If
            End If

            If Me.txtDistict.Text.Trim.Length > 0 Then
                If Not blnWhere Then
                    strQuery += " where upper(District) like '" & txtDistict.Text.Trim.ToUpper & "%'"
                    blnWhere = True
                Else
                    strQuery += " and upper(District) like '" & txtDistict.Text.Trim.ToUpper & "%'"
                End If
            End If

            If Me.txtContact.Text.Trim.Length > 0 Then
                If Not blnWhere Then
                    strQuery += " where upper(contact) like '%" & txtContact.Text.Trim.ToUpper & "%'"
                    blnWhere = True
                Else
                    strQuery += " and upper(contact) like '%" & txtContact.Text.Trim.ToUpper & "%'"
                End If
            End If

            Try

                If Me.cbCountries.SelectedIndex > 0 Then
                    Dim oCountry As MedscreenLib.Address.Country = Countries.Item(Me.cbCountries.SelectedIndex)
                    If Not oCountry Is Nothing Then         ' Is it a valid country 
                        Dim intCnt As Integer = oCountry.CountryId
                        If Not blnWhere Then
                            strQuery += " where country = " & CStr(intCnt)
                            blnWhere = True
                        Else
                            strQuery += " and country = " & CStr(intCnt)
                        End If
                    End If
                End If
            Catch ex As Exception
            End Try



            If Me.txtLine1.Text.Trim.Length > 0 Then
                If Not blnWhere Then
                    strQuery += " where ((upper(ADDRLINE1) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE2) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE3) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE4) like '%" & txtLine1.Text.Trim.ToUpper & "%'))"

                    blnWhere = True
                Else
                    strQuery += " and ((upper(ADDRLINE1) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE2) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE3) like '%" & txtLine1.Text.Trim.ToUpper & "%') or " & _
                        "(upper(ADDRLINE4) like '%" & txtLine1.Text.Trim.ToUpper & "%'))"
                End If
            End If

        End If


        'Me.Addresses1 = MyAddresses.ToDataSet()

    End Sub



    Private Sub txtCity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCity.TextChanged
        SetUpQuery()
    End Sub

    Private Sub txtLine1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtLine1.TextChanged
        SetUpQuery()
    End Sub


    Private Sub txtDistict_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDistict.TextChanged
        SetUpQuery()
    End Sub

    Private Sub txtContact_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtContact.TextChanged
        SetUpQuery()
    End Sub

    '''<summary>
    ''' Id of the address selected by the user
    ''' </summary>
    Public Property AddressID() As Integer
        Get
            Return myAddressId
        End Get
        Set(ByVal Value As Integer)
            myAddressId = Value
        End Set
    End Property

    '''<summary>
    ''' Force a requery of the database
    ''' </summary>
    Public Sub Requery()
        Me.txtCity.Clear()
        Me.txtContact.Clear()
        Me.txtDistict.Clear()
        Me.txtLine1.Clear()
        Me.txtPostCode.Clear()
        SetUpQuery()

    End Sub

    Private Sub DGAddress_Navigate(ByVal sender As System.Object, ByVal ne As System.Windows.Forms.NavigateEventArgs)

    End Sub

    Private Sub DGAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.Click
        Dim oAddr As MedscreenLib.Address.Caddress

        If Me.ListView1.SelectedItems.Count < 1 Then Exit Sub

        Dim lvItem As AddressItem = Me.ListView1.SelectedItems.Item(0)
        oAddr = lvItem.Address
        If Not oAddr Is Nothing Then
            Me.myAddressId = oAddr.AddressId
        End If
    End Sub

    Private Sub DGAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        Me.DialogResult = DialogResult.OK

    End Sub

    Private Sub cbCountries_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbCountries.SelectedIndexChanged
        ' If Not FillingForm Then Me.SetUpQuery()
    End Sub


    Private Class AddressItem
        Inherits System.Windows.Forms.ListViewItem

        Private myAddress As MedscreenLib.Address.Caddress

        Public Sub New(ByVal inAddress As MedscreenLib.Address.Caddress)
            MyBase.new(inAddress.AddressId)
            myAddress = inAddress

            Dim lvItem As ListViewItem.ListViewSubItem

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.AdrLine1)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.AdrLine2)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.AdrLine3)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.AdrLine4)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.City)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.District)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.PostCode)
            Me.SubItems.Add(lvItem)

            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.Contact)
            Me.SubItems.Add(lvItem)


            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.Email)
            Me.SubItems.Add(lvItem)
            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.Phone)
            Me.SubItems.Add(lvItem)
            lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.Fax)
            Me.SubItems.Add(lvItem)
            Dim oCountry As MedscreenLib.Address.Country = Countries.Item(myAddress.Country, 0)
            If oCountry Is Nothing Then
                lvItem = New ListViewItem.ListViewSubItem(Me, myAddress.CountryName)
            Else
                lvItem = New ListViewItem.ListViewSubItem(Me, oCountry.CountryName)
            End If
            Me.SubItems.Add(lvItem)

        End Sub

        Public Property Address() As MedscreenLib.Address.Caddress
            Get
                Return Me.myAddress
            End Get
            Set(ByVal Value As MedscreenLib.Address.Caddress)
                myAddress = Value
            End Set
        End Property
    End Class

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView1.ColumnClick
        Dim intdirection As Integer

        intdirection = Me.ColumnDirection(e.Column)
        'myUserOptions.DiaryColumn = e.Column
        'myUserOptions.DiarySortDirection = intdirection
        Me.ColumnDirection(e.Column) *= -1

        Me.ListView1.ListViewItemSorter = New Comparers.ListViewItemComparer(e.Column, intdirection)

    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        If Me.ListView1.Items.Count = 0 Then Me.SetUpQuery()
    End Sub

    Private Sub CmdFind_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdFind.Click
        MyAddresses = New MedscreenLib.Address.AddressCollection()

        Me.SetUpQuery()
        Dim oAddr As MedscreenLib.Address.Caddress

        Dim OCmd As New OleDb.OleDbCommand()
        Dim oRead As OleDb.OleDbDataReader
        Dim intId As Integer

        Try
            OCmd.Connection = DBConnection
            SetConnOpen()
            OCmd.CommandText = strQuery
            oRead = OCmd.ExecuteReader()
            Me.ListView1.BeginUpdate()
            Me.ListView1.Items.Clear()
            Dim i As Integer
            Dim AColl As New Collection()
            While oRead.Read
                intId = oRead.GetValue(0)
                AColl.Add(intId)
            End While
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If

            For i = 1 To AColl.Count
                intId = AColl.Item(i)
                oAddr = SMAddresses.item(intId, MedscreenLib.Address.AddressCollection.ItemType.AddressId)
                MyAddresses.Add(oAddr)
                Dim lvAddress As AddressItem
                lvAddress = New AddressItem(oAddr)
                Me.ListView1.Items.Add(lvAddress)

                If i > 20 Then Me.ListView1.EndUpdate()
            Next

        Catch ex As Exception
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            SetConnClosed()
            Me.ListView1.EndUpdate()
        End Try

    End Sub
End Class