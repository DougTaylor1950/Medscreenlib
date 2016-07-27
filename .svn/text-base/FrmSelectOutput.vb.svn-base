'$Revision: 1.0 $
'$Author: taylor $
'$Date: 2005-09-02 07:41:45+01 $
'$Log: Class1.vb,v $
'<>

Imports System.Drawing.Printing
Imports System.Drawing

'''<summary>
''' Output method and if print which printer
''' </summary>
Public Class frmOutputSelect
    Inherits System.Windows.Forms.Form

    Public Enum PrinterOutput
        DeviceName
        LogicalName
    End Enum
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.cbOption.Items.Add(cstEMAILOption) ', cstPDFOption, cstFaxOption, cstPrinterOption})
        Me.cbOption.Items.Add(cstPDFOption)
        Me.cbOption.Items.Add(cstFaxOption)
        Me.cbOption.Items.Add(cstPrinterOption)
        'Add any initialization after the InitializeComponent() call
        PopulateInstalledPrintersCombo()
        'TT2.Style = CBalloonToolTip.ttStyleEnum.TTBalloon
        'TT2.CreateToolTip(Me.CBInstalledPrinters.Handle.ToInt32, "Printers available")

        'TT2.Icon(Me.CBInstalledPrinters.Handle.ToInt32) = CBalloonToolTip.ttIconType.TTIconInfo
        'TT2.Title(Me.CBInstalledPrinters.Handle.ToInt32) = "Information"
        'TT2.PopupOnDemand = False
        'TT2.VisibleTime(Me.CBInstalledPrinters.Handle.ToInt32) = 6000 'After 6 Seconds tooltip will go away
        'Dim myFont As New Font("Tahoma", 10, FontStyle.Italic Or FontStyle.Underline)
        'TT2.TipFont(Me.CBInstalledPrinters.Handle.ToInt32) = myFont

        'TT2.CreateToolTip(Me.cbOption.Handle.ToInt32, "Output Methods Available")
        'TT2.CreateToolTip(Me.txtDestination.Handle.ToInt32, "Where document will be printed Emailed")
        'TT2.CreateToolTip(Me.cmdCancel.Handle.ToInt32, "Cancel to stop message going out")
        'TT2.CreateToolTip(Me.Button1.Handle.ToInt32, "Email/Print document to selected destination")
        'TT2.CreateToolTip(Me.cbOption.Handle.ToInt32, "Output Methods Available")

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbOption As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtDestination As System.Windows.Forms.TextBox
    Friend WithEvents PrintDialog1 As System.Windows.Forms.PrintDialog
    Friend WithEvents CBInstalledPrinters As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents HelpProvider1 As System.Windows.Forms.HelpProvider
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOutputSelect))
        Me.Label1 = New System.Windows.Forms.Label
        Me.cbOption = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtDestination = New System.Windows.Forms.TextBox
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog
        Me.CBInstalledPrinters = New System.Windows.Forms.ComboBox
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider(Me.components)
        Me.HelpProvider1 = New System.Windows.Forms.HelpProvider
        Me.Panel1.SuspendLayout()
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Output Method"
        '
        'cbOption
        '
        Me.cbOption.Location = New System.Drawing.Point(120, 24)
        Me.cbOption.Name = "cbOption"
        Me.cbOption.Size = New System.Drawing.Size(240, 21)
        Me.cbOption.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cbOption, "The way in which the document will be sent")
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "To"
        '
        'txtDestination
        '
        Me.txtDestination.Location = New System.Drawing.Point(120, 56)
        Me.txtDestination.Name = "txtDestination"
        Me.txtDestination.Size = New System.Drawing.Size(160, 20)
        Me.txtDestination.TabIndex = 3
        '
        'CBInstalledPrinters
        '
        Me.CBInstalledPrinters.Location = New System.Drawing.Point(120, 56)
        Me.CBInstalledPrinters.Name = "CBInstalledPrinters"
        Me.CBInstalledPrinters.Size = New System.Drawing.Size(280, 21)
        Me.CBInstalledPrinters.TabIndex = 4
        Me.CBInstalledPrinters.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Controls.Add(Me.cmdCancel)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 101)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(408, 40)
        Me.Panel1.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Image = CType(resources.GetObject("cmdCancel.Image"), System.Drawing.Image)
        Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCancel.Location = New System.Drawing.Point(320, 8)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button1.Location = New System.Drawing.Point(232, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Ok"
        Me.Button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ToolTip1
        '
        Me.ToolTip1.IsBalloon = True
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        '
        'frmOutputSelect
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(408, 141)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.CBInstalledPrinters)
        Me.Controls.Add(Me.txtDestination)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbOption)
        Me.Controls.Add(Me.Label1)
        Me.Name = "frmOutputSelect"
        Me.Text = "Output Select "
        Me.Panel1.ResumeLayout(False)
        CType(Me.ErrorProvider1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private strDefaultOutput As String = ""
    Private myPrinterOutput As PrinterOutput = PrinterOutput.DeviceName
    'Dim TT2 As New CBalloonToolTip() '//Demo for mouse over tooltip

    Public Const cstPrinterOption As String = "PRINTER"
    Public Const cstFaxOption As String = "FAX"
    Public Const cstPDFOption As String = "PDF"
    Public Const cstEMAILOption As String = "EMAIL"

    
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Output the following info
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property PrinterOutputTo() As PrinterOutput
        Get
            Return myPrinterOutput
        End Get
        Set(ByVal Value As PrinterOutput)
            myPrinterOutput = Value
        End Set
    End Property



    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 10:28:12
    '''     Entering destination text box
    ''' </summary>
    ''' <param name="sender" type="Object">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <param name="e" type="EventArgs">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub txtDestination_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDestination.Enter
        If Me.cbOption.Text.ToUpper = cstPrinterOption Then

        End If
    End Sub

    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 10:31:16
    '''     Populate list of installed Printers
    ''' </summary>
    ''' <remarks>
    ''' Gets information from the printers table
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub PopulateInstalledPrintersCombo()
        ' Add list of installed printers found to the combo box.
        ' The pkInstalledPrinters string will be used to provide the display string.
        'Dim pkInstalledPrinters As String
        Dim StrDef As String = ""


        'Dim inF As IniFile.IniFiles = New IniFile.IniFiles()

        'inF.FileName = MedscreenLib.Medscreen.LiveRoot & MedscreenLib.Medscreen.IniFiles & "reporterprinters.ini"
        'Dim PColl As Collection = New Collection()
        'PColl = inF.ReadSection("Printers")

        'Dim prColl As New IniFile.IniCollection()
        'Dim pStr As String
        'Dim i As Integer
        'Dim iItem As IniFile.IniElement
        ' Dim inIndex As Integer

        'For i = 1 To PColl.Count
        '    pStr = PColl.Item(i)
        '    inIndex = InStr(pStr, "=")
        '    If inIndex > 0 Then
        '        iItem = New IniFile.IniElement()
        '        iItem.Header = Mid(pStr, 1, inIndex - 1)
        '        iItem.Item = Mid(pStr, inIndex + 1)
        '        prColl.Add(iItem)
        '        Me.CBInstalledPrinters.Items.Add(iItem.Item)
        '    End If
        'Next
        'If prColl.Count > 0 Then
        Me.CBInstalledPrinters.DataSource = MedscreenLib.Glossary.Glossary.Printers
        Me.CBInstalledPrinters.ValueMember = "DeviceName"
        Me.CBInstalledPrinters.DisplayMember = "Info"
        Me.CBInstalledPrinters.SelectedIndex = 0
        'End If
    End Sub

    '''<summary>
    ''' Output method
    ''' </summary>
    Public Property OutputMethod() As String
        Get
            Return Me.cbOption.Text
        End Get
        Set(ByVal Value As String)
            Dim i As Integer
            For i = 0 To Me.cbOption.Items.Count - 1
                If Me.cbOption.Items(i) = Value Then
                    Me.cbOption.SelectedIndex = i
                    Exit For
                End If
            Next
        End Set
    End Property

    '''<summary>
    ''' Output destination represented by <see cref="MedscreenLib.Constants.SendMethod" />
    ''' </summary>
    Public ReadOnly Property OutputMethodByType() As MedscreenLib.Constants.SendMethod
        Get
            '"EMAIL", "EMAIL (PDF)", "FAX", "PRINTER"
            If Me.OutputMethod = cstEMAILOption Then
                Return Constants.SendMethod.Email
            ElseIf Me.OutputMethod = cstPDFOption Then
                Return Constants.SendMethod.PDF
            ElseIf Me.OutputMethod = cstFaxOption Then
                Return Constants.SendMethod.Fax
            ElseIf Me.OutputMethod = cstPrinterOption Then
                Return Constants.SendMethod.Printer
            End If
        End Get
    End Property

    '''<summary>
    ''' Where to print to
    ''' </summary>
    Public Property OutputDestination() As String
        Get
            If Me.cbOption.Text.ToUpper <> cstPrinterOption Then
                Return Me.txtDestination.Text
            Else
                'Dim i As Integer
                Dim strRet As String = ""
                strRet = Me.CBInstalledPrinters.SelectedValue
                If myPrinterOutput = PrinterOutput.LogicalName Then
                    strRet = CType(Me.CBInstalledPrinters.SelectedItem, MedscreenLib.Glossary.Printer).LogicalName
                End If
                'For i = 0 To Me.CBInstalledPrinters.Text.Length - 1
                '    If Me.CBInstalledPrinters.Text.Chars(i) = "/" Then
                '        Exit For
                '    End If
                '    strRet += Me.CBInstalledPrinters.Text.Chars(i)
                'Next
                Return strRet
            End If
        End Get
        Set(ByVal Value As String)
            Try

                strDefaultOutput = Value
                Dim i As Integer
                For i = 0 To Me.CBInstalledPrinters.Items.Count - 1
                    'See if printer matches required device type
                    If InStr(CType(Me.CBInstalledPrinters.Items(i), MedscreenLib.Glossary.Printer).DeviceName.ToUpper, Value.ToUpper) > 0 Then
                        Me.CBInstalledPrinters.SelectedIndex = i
                        Exit For
                    End If
                Next
            Catch ex As System.Exception
                ' Handle exception here
                MedscreenLib.Medscreen.LogError(ex, , "Set printer in output selection ")
            End Try
        End Set
    End Property


    ''' <summary>
    ''' Created by taylor on ANDREW at 19/01/2007 10:12:37
    '''     Deal with selected index changing
    ''' </summary>
    ''' <param name="sender" type="Object">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <param name="e" type="EventArgs">
    '''     <para>
    '''         
    '''     </para>
    ''' </param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 19/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Private Sub cbOption_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbOption.SelectedIndexChanged
        If Me.cbOption.Text.ToUpper = cstPrinterOption Then
            Me.CBInstalledPrinters.Show()
            Me.txtDestination.Hide()
            'TT2.TipText(Me.cbOption.Handle.ToInt32) = "Output to chosen printer"
            Me.ToolTip1.SetToolTip(Me.cbOption, "Output to chosen printer")

        Else
            Me.CBInstalledPrinters.Hide()
            Me.txtDestination.Show()
            ' If Not Me.m_objAddress Is Nothing Then Exit Sub

            ' 
            If Me.cbOption.Text = cstEMAILOption Then
                If Not Me.m_objAddress Is Nothing Then
                    Me.txtDestination.Text = Me.m_objAddress.Email
                Else
                    Me.txtDestination.Text = strDefaultOutput
                End If
                'TT2.TipText(Me.cbOption.Handle.ToInt32) = "Output to Email address"
                Me.ToolTip1.SetToolTip(Me.cbOption, "Output to Email Address")


            ElseIf Me.cbOption.Text = cstFaxOption Then
                'Sort out destination address
                Me.ToolTip1.SetToolTip(Me.cbOption, "Output to Fax Address")
                'TT2.TipText(Me.cbOption.Handle.ToInt32) = "Output to Fax address"

                If Me.m_objAddress Is Nothing Then Exit Sub
                Dim strDest As String = Me.m_objAddress.Fax.Replace("(", "")
                strDest = strDest.Replace(")", "")
                strDest = strDest.Replace("+", "00")
                Me.txtDestination.Text = strDest & "@starfax.co.uk"
            ElseIf Me.cbOption.Text = cstPDFOption Then
                'TT2.TipText(Me.cbOption.Handle.ToInt32) = "Output to Email address as PDF"
                Me.ToolTip1.SetToolTip(Me.cbOption, "Output to Email Address as PDF")

                If Not Me.m_objAddress Is Nothing Then
                    Me.txtDestination.Text = Me.m_objAddress.Email
                Else
                    Me.txtDestination.Text = strDefaultOutput
                End If
            End If
        End If
    End Sub
#Region "Property 'AddressID'"
    ''' <summary>
    '''     Private address ID variable
    ''' </summary>
    Private m_intAddressID As Integer = -1
    ''' <summary>
    '''     Private address object variable
    ''' </summary>
    
    Private m_objAddress As MedscreenLib.Address.Caddress
    ''' <summary>
    '''     Property to return the address id.
    ''' </summary>
    Public Property AddressID() As Integer
        Get
            Return m_intAddressID
        End Get
        Set(ByVal Value As Integer)
            m_intAddressID = Value
            If m_intAddressID > 0 Then      'Get address property 
                m_objAddress = New Address.Caddress(Me.m_intAddressID)
                m_objAddress.Fields.Load(CConnection.DbConnection, "address_id = " & m_objAddress.AddressId)
            End If
        End Set
    End Property

#End Region 'Property 'AddressID'
End Class
