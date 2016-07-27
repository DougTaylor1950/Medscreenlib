Imports System.Windows.Forms
Imports System.Drawing
Imports MedscreenLib

Namespace Panels

    Public Class DiscountPanelSet
#Region "Declarations"
        Private strPanel As String = ""         'Define the containing panel
        Private strHandler As String = ""       'Define the action handler
        Private myPanels As DiscountPanelCollection
        Private myPanel As Panel
#End Region

#Region "Public Instance"

#Region "Functions"
        Public Function BuildControls(ByVal inForm As Form)
            ' get containing panel from form
            Dim aControl As Control
            Dim cControl As Control
            For Each aControl In inForm.Controls
                cControl = FindControls(aControl)
                If Not cControl Is Nothing Then
                    myPanel = cControl
                    Exit For
                End If
            Next
            'see if we have the panel
            If Not myPanel Is Nothing Then
                Dim anpanel As New DiscountPanel()
                Dim SenderEvents() As System.Reflection.EventInfo = anpanel.GetType().GetEvents()
                Dim ReceiverType As Type = inForm.GetType()
                Dim Method As System.Reflection.MethodInfo
                Dim E As System.Reflection.EventInfo
                For Each E In SenderEvents
                    If E.Name = "PriceChanged" Then
                        Exit For
                    End If
                Next
                Dim D As System.Delegate
                Method = ReceiverType.GetMethod( _
  Me.ActionHandler, _
  System.Reflection.BindingFlags.IgnoreCase Or _
  System.Reflection.BindingFlags.Instance Or _
  System.Reflection.BindingFlags.NonPublic)
                If Not Method Is Nothing Then
                    D = System.Delegate.CreateDelegate(E.EventHandlerType, inForm, Method.Name)
                End If
                'Next

                'we can now attempt to add each item in the panel collection 
                Dim aPanel As DiscountPanelBase
                For Each aPanel In Panels
                    Dim myDiscPanel As DiscountPanel = New DiscountPanel()
                    myDiscPanel.AllowLeave = True
                    myDiscPanel.BackColor = System.Drawing.Color.Beige
                    myDiscPanel.Discount = ""
                    myDiscPanel.DiscountScheme = ""
                    myDiscPanel.Dock = System.Windows.Forms.DockStyle.Top
                    myDiscPanel.LineItem = aPanel.LineItem
                    myDiscPanel.ListPrice = ""
                    myDiscPanel.Location = New System.Drawing.Point(aPanel.Left, aPanel.Top)
                    myDiscPanel.Name = aPanel.Name
                    myDiscPanel.PanelLabel = aPanel.Label
                    myDiscPanel.PanelService = Nothing
                    myDiscPanel.PanelType = aPanel.PanelType
                    myDiscPanel.PriceEnabled = True
                    myDiscPanel.Scheme = ""
                    myDiscPanel.Service = aPanel.Service
                    myDiscPanel.size = New System.Drawing.Size(aPanel.Width, aPanel.Height)
                    myPanel.Controls.Add(myDiscPanel)
                    aPanel.Panel = myDiscPanel
                    E.AddEventHandler(myDiscPanel, D)
                    'addhandler(mydiscpanel.PriceChanged )
                Next
            End If
        End Function

        Private Function FindControls(ByVal bControl As Control) As Control
            Dim cControl As Control
            If bControl.Name.ToUpper = strPanel.ToUpper Then
                Return bControl
            Else
                Dim aControl As Control
                For Each aControl In bControl.Controls
                    cControl = FindControls(aControl)
                    If Not cControl Is Nothing Then

                        Exit For
                    End If
                    cControl = Nothing
                Next
            End If
            Return cControl
        End Function
#End Region

#Region "Procedures"
        Public Sub New()
            MyBase.new()
            myPanels = New DiscountPanelCollection()
        End Sub

#End Region

#Region "Properties"
        Public Property Panels() As DiscountPanelCollection
            Get
                Return myPanels
            End Get
            Set(ByVal Value As DiscountPanelCollection)
                myPanels = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Parent control for panels 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ParentPanel() As String
            Get
                Return Me.strPanel
            End Get
            Set(ByVal Value As String)
                strPanel = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Event that will be fired
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------

        Public Property ActionHandler() As String
            Get
                Return Me.strHandler
            End Get
            Set(ByVal Value As String)
                strHandler = Value
            End Set
        End Property

#End Region
#End Region

  

 
    End Class
    Public Class DiscPanelCollCollection
        Inherits CollectionBase
        Public Sub New()
            MyBase.new()
        End Sub

        Default Public Property Item(ByVal index As Integer) As DiscountPanelSet
            Get
                Return CType(MyBase.List.Item(index), DiscountPanelSet)
            End Get
            Set(ByVal Value As DiscountPanelSet)
                MyBase.List(index) = Value
            End Set
        End Property

        Public Function Add(ByVal item As DiscountPanelSet)
            Return MyBase.List.Add(item)
        End Function

        Public Function Filename(Optional ByVal Opt As Integer = 0) As String
            Dim strFilename As String = Application.ExecutablePath

            Dim i As Integer = strFilename.Length

            While i > 0 And strFilename.Chars(i - 1) <> "\"
                i -= 1
            End While

            strFilename = Mid(strFilename, 1, i) & "PanelDetails.xml"


            'IO.Directory.SetCurrentDirectory(Application.ExecutablePath)

            Return strFilename
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Save contents of collection to XML file
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	17/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Save() As Boolean
            Dim myWriter As IO.StreamWriter
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(DiscPanelCollCollection))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = Filename()
                myWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
            Catch ex As Exception
            Finally
                If Not myWriter Is Nothing Then
                    myWriter.Flush()
                    myWriter.Close()
                End If
            End Try
        End Function

        Private Function DoLoad(ByVal strFilename As String) As Boolean
            Dim blnReturn As Boolean = True
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(DiscPanelCollCollection))
                If IO.File.Exists(strFilename) Then
                    Try
                        Dim myFileStream As IO.FileStream = _
                        New IO.FileStream(strFilename, IO.FileMode.Open)
                        Dim Myobject As DiscPanelCollCollection = CType( _
                        mySerializer.Deserialize(myFileStream), DiscPanelCollCollection)
                        Me.Clear()
                        Dim objPanel As DiscountPanelSet
                        For Each objPanel In Myobject.List
                            Me.Add(objPanel)
                        Next
                        myFileStream.Close()
                    Catch ex As Exception
                        blnReturn = False
                    End Try
                Else
                    MsgBox("No user options found")
                    blnReturn = False
                End If
            Catch ex As Exception
            Finally

            End Try
            Return blnReturn
        End Function

        Public Function Load(Optional ByVal NormalLoad As Boolean = True) As Boolean
            ' To read the file, create a FileStream object.
            Dim strFileName As String
            'If NormalLoad Then
            strFileName = Filename()
            Return DoLoad(strFileName)
        End Function


    End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Panels.DiscountPanelCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks>
    ''' Will be used to serialise the discount panels so a dynamic system can be set up.
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	16/12/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class DiscountPanelCollection
        Inherits CollectionBase


#Region "Public Instance"

#Region "Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an item to the collection 
        ''' </summary>
        ''' <param name="item"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As DiscountPanelBase) As Integer
            Return MyBase.List.Add(item)

        End Function

        Public Function Filename(Optional ByVal Opt As Integer = 0) As String
            Dim strFilename As String = Application.ExecutablePath

            Dim i As Integer = strFilename.Length

            While i > 0 And strFilename.Chars(i - 1) <> "\"
                i -= 1
            End While

            strFilename = Mid(strFilename, 1, i) & "PanelDetail.xml"


            'IO.Directory.SetCurrentDirectory(Application.ExecutablePath)

            Return strFilename
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Save() As Boolean
            Dim myWriter As IO.StreamWriter
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(DiscountPanelCollection))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = Filename()
                myWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
            Catch ex As Exception
            Finally
                If Not myWriter Is Nothing Then
                    myWriter.Flush()
                    myWriter.Close()
                End If
            End Try
        End Function

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new collection 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub


#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Item in collection 
        ''' </summary>
        ''' <param name="index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	16/12/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As Integer) As DiscountPanelBase
            Get
                Return CType(MyBase.List.Item(index), DiscountPanelBase)
            End Get
            Set(ByVal Value As DiscountPanelBase)
                MyBase.List.Item(index) = Value
            End Set
        End Property

#End Region
#End Region



    End Class

    Public Class DiscountPanelBase
#Region "Declarations"
        Private myService As String = ""
        Private myLineItem As String = ""
        Private myDiscounts As String = ""
        Private myName As String
        Private myLabel As String = ""
        Private myHeight As Integer = 0
        Private myWidth As Integer = 0
        Private myLeft As Integer = 0
        Private myTop As Integer = 0
        Private myPanel As Object
        Private myType As Integer
        Private myPriceDate As Date = today

        Private redrawing As Boolean = False
#End Region

#Region "Public Instance"
#Region "Properties"
        Public Property PanelType() As Integer
            Get
                Return myType
            End Get
            Set(ByVal Value As Integer)
                myType = Value
            End Set
        End Property

        Public Property PriceDate() As Date
            Get
                Return mypricedate
            End Get
            Set(ByVal Value As Date)
                myPriceDate = Value
            End Set
        End Property

        Public Property Panel() As Object
            Get
                Return myPanel
            End Get
            Set(ByVal Value As Object)
                myPanel = Value
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return myWidth
            End Get
            Set(ByVal Value As Integer)
                myWidth = Value
            End Set
        End Property

        Public Property Height() As Integer
            Get
                Return myHeight
            End Get
            Set(ByVal Value As Integer)
                myHeight = Value
            End Set
        End Property

        Public Property Left() As Integer
            Get
                Return myLeft
            End Get
            Set(ByVal Value As Integer)
                myLeft = Value
            End Set
        End Property

        Public Property Top() As Integer
            Get
                Return myTop
            End Get
            Set(ByVal Value As Integer)
                myTop = Value
            End Set
        End Property
        Public Property Label() As String
            Get
                Return myLabel
            End Get
            Set(ByVal Value As String)
                myLabel = Value
            End Set
        End Property

        Public Property Name() As String
            Get
                Return myName
            End Get
            Set(ByVal Value As String)
                myName = Value
            End Set
        End Property
        Public Property DiscountScheme() As String
            Get
                Return myDiscounts
            End Get
            Set(ByVal Value As String)
                myDiscounts = Value
            End Set
        End Property

        Public Property LineItem() As String
            Get
                If Me.myLineItem Is Nothing Then
                    myLineItem = ""
                End If
                Return myLineItem
            End Get
            Set(ByVal Value As String)
                myLineItem = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Allow event to fire
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/07/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property AllowLeave() As Boolean
            Get
                Return Not Me.redrawing
            End Get
            Set(ByVal Value As Boolean)
                redrawing = Not Value
            End Set
        End Property

        Public Property Service() As String
            Get
                Return myService
            End Get
            Set(ByVal Value As String)
                myService = Value
            End Set
        End Property


#End Region
#End Region

        Public Sub New()
            MyBase.new()
        End Sub

        Public Function Save() As Boolean
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(DiscountPanelBase))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = "x:\doug\panel.xl"
                Dim myWriter As IO.StreamWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
                myWriter.Flush()
                myWriter.Close()
            Catch ex As Exception
            End Try
        End Function
    End Class
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : Panels.DiscountPanel
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Panel to handle discounts
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class DiscountPanel




        Inherits Panel


        'Private components As System.ComponentModel.Container = Nothing

#Region "Declarations"
        'Private myLineItem As String = ""
        'Private myDiscounts As String = ""
        Private myClient As Intranet.intranet.customerns.Client
        Private objService As Intranet.intranet.customerns.CustomerService
        Private myPanelType As TypesOfPanel

        Private blnDisableAll As Boolean = False
        Private oldPrice As Double
        'Private redrawing As Boolean = False
        Private myBasePanel As DiscountPanelBase

#Region "enumerations"
        Public Enum TypesOfPanel
            Sample
            MinimumCharge
            AdditionalTime
            MRO
            Cancellation
            CollectionCharge
        End Enum

        Public Enum ServiceTypes
            ROUTINE
            CALLOUT
            ROUTINEOUT
            CALLOUTOUT
        End Enum


#End Region

#Region "Events"
        Public Event PriceChanged(ByVal Sender As DiscountPanel, ByVal NewPrice As String)

        Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
        Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox

#End Region
#End Region

#Region "Public Instance"

#Region "Functions"
        Public Function save()
            Me.myBasePanel.Save()
        End Function
#End Region

#Region "Procedures"

#End Region

#Region "Properties"

#End Region
#End Region
#Region " Component Designer generated code "

        Public Sub New(ByVal Container As System.ComponentModel.IContainer)
            MyClass.New()

            'Required for Windows.Forms Class Composition Designer support
            Container.Add(Me)
        End Sub


        Public Shadows Property Name() As String
            Get
                Return MyBase.Name
            End Get
            Set(ByVal Value As String)
                MyBase.Name = Value
                myBasePanel.Name = Value
            End Set
        End Property
        Public Sub New()
            MyBase.New()

            'This call is required by the Component Designer.
            InitializeComponent()

            'Add any initialization after the InitializeComponent() call
            Me.myBasePanel = New DiscountPanelBase()
            reDraw(True)
        End Sub

        'Component overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        Public Property Service() As String
            Get
                Return Me.myBasePanel.Service
            End Get
            Set(ByVal Value As String)
                myBasePanel.Service = Value
            End Set
        End Property

        Public Property PriceDate() As Date
            Get
                Return myBasePanel.PriceDate

            End Get
            Set(ByVal Value As Date)
                myBasePanel.PriceDate = Value
            End Set
        End Property

        Public Property PanelLabel() As String
            Get
                Return Me.ActionLabel.Text
            End Get
            Set(ByVal Value As String)
                Me.ActionLabel.Text = Value
                Me.BasePanel.Label = Value
            End Set
        End Property

        Public Shadows Property size() As System.Drawing.Size
            Get
                Return MyBase.Size
            End Get
            Set(ByVal Value As System.Drawing.Size)
                MyBase.Size = Value
                myBasePanel.Height = Value.Height
                myBasePanel.Width = Value.Width
            End Set
        End Property

        Public Shadows Property Location() As System.Drawing.Point
            Get
                Return MyBase.Location
            End Get
            Set(ByVal Value As System.Drawing.Point)
                MyBase.Location = Value
                myBasePanel.Top = Value.Y
                myBasePanel.Left = Value.X
            End Set
        End Property

        Public WriteOnly Property PictureBoxBitmap() As Bitmap
            Set(ByVal Value As Bitmap)
                Me.PictureBox1.Image = Value
            End Set
        End Property

        Public WriteOnly Property InfoBitmap() As Bitmap
            Set(ByVal Value As Bitmap)
                Me.PictureBox2.Image = Value
                Me.PictureBox2.Show()
            End Set
        End Property

        'Required by the Component Designer
        Private components As System.ComponentModel.IContainer
        Friend WithEvents ActionLabel As System.Windows.Forms.Label
        Friend WithEvents txtListPrice As System.Windows.Forms.TextBox
        Friend WithEvents txtScheme As System.Windows.Forms.TextBox
        Friend WithEvents txtDiscount As System.Windows.Forms.TextBox
        Friend WithEvents txtPrice As System.Windows.Forms.TextBox
        Friend WithEvents ToolTip1 As MedscreenCommonGui.EnhancedToolTip
        Friend WithEvents SchemeLabel As System.Windows.Forms.Label
        Friend WithEvents lblPricephr As System.Windows.Forms.Label

        'NOTE: The following procedure is required by the Component Designer
        'It can be modified using the Component Designer.
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            components = New System.ComponentModel.Container()
            Me.ActionLabel = New System.Windows.Forms.Label()
            Me.txtListPrice = New System.Windows.Forms.TextBox()
            Me.txtDiscount = New System.Windows.Forms.TextBox()
            Me.txtScheme = New System.Windows.Forms.TextBox()
            Me.txtPrice = New System.Windows.Forms.TextBox()
            Me.ToolTip1 = New MedscreenCommonGui.EnhancedToolTip(Me.components)
            Me.SchemeLabel = New System.Windows.Forms.Label()
            Me.PictureBox1 = New System.Windows.Forms.PictureBox()
            Me.PictureBox2 = New System.Windows.Forms.PictureBox()

            Me.lblPricephr = New System.Windows.Forms.Label()

            Me.Controls.Add(Me.ActionLabel)
            Me.Controls.Add(Me.txtListPrice)
            Me.Controls.Add(Me.txtDiscount)
            Me.Controls.Add(Me.txtScheme)
            Me.Controls.Add(Me.txtPrice)
            Me.Controls.Add(Me.SchemeLabel)
            Me.Controls.Add(Me.lblPricephr)
            Me.Controls.Add(Me.PictureBox1)
            Me.Controls.Add(Me.PictureBox2)
            'Me.Controls.Add(Me.ToolTip1)
            '
            'Label
            '
            Me.ActionLabel.Location = New System.Drawing.Point(32, 9)
            Me.ActionLabel.Name = "Label4"
            Me.ActionLabel.TabIndex = 36
            Me.ActionLabel.Text = ""
            Me.ActionLabel.Size = New System.Drawing.Size(195, 40)
            '
            'txtListPrice
            '
            Me.txtListPrice.Enabled = False
            Me.txtListPrice.Location = New System.Drawing.Point(231, 9)
            Me.txtListPrice.Name = "txtLp"
            Me.txtListPrice.Size = New System.Drawing.Size(72, 20)
            Me.txtListPrice.TabIndex = 37
            Me.txtListPrice.Text = ""
            '
            'txtDiscount
            '
            Me.txtDiscount.BackColor = System.Drawing.SystemColors.Info
            Me.txtDiscount.Location = New System.Drawing.Point(343, 9)
            Me.txtDiscount.Name = "txtDis"
            Me.txtDiscount.Size = New System.Drawing.Size(72, 20)
            Me.txtDiscount.TabIndex = 38
            Me.txtDiscount.Text = ""
            Me.txtDiscount.Enabled = False
            '
            'txtscheme
            '
            Me.txtScheme.Enabled = False
            Me.txtScheme.Location = New System.Drawing.Point(578, 9)
            Me.txtScheme.Name = "txtsch"
            Me.txtScheme.Size = New System.Drawing.Size(72, 20)
            Me.txtScheme.TabIndex = 40
            Me.txtScheme.Text = ""
            Me.txtScheme.Visible = False
            '
            'Label
            '
            Me.SchemeLabel.Location = New System.Drawing.Point(578, 9)
            Me.SchemeLabel.Name = "Label5"
            Me.SchemeLabel.TabIndex = 36
            Me.SchemeLabel.Text = ""
            Me.SchemeLabel.Size = New System.Drawing.Size(195, 20)
            Me.SchemeLabel.Anchor = ((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
            '
            'txtPrice
            '
            Me.txtPrice.Location = New System.Drawing.Point(463, 9)
            Me.txtPrice.Name = "txtPrice"
            Me.txtPrice.Size = New System.Drawing.Size(72, 20)
            Me.txtPrice.TabIndex = 39
            Me.txtPrice.Text = ""
            '
            'lblpricephr
            '
            Me.lblPricephr.Location = New System.Drawing.Point(528, 8)
            Me.lblPricephr.Name = "lblpricephr"
            Me.lblPricephr.Size = New System.Drawing.Size(40, 23)
            Me.lblPricephr.TabIndex = 42
            '
            'PictureBox1
            '

            Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form))


            'Me.PictureBox1.Image = CType(Resources.GetObject("PictureBox1.Image"), System.Drawing.Bitmap)
            Me.PictureBox1.Location = New System.Drawing.Point(10, 10)
            Me.PictureBox1.Name = "PictureBox1"
            Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
            Me.PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            Me.PictureBox1.TabIndex = 61
            Me.PictureBox1.TabStop = False
            '
            'Picturebox2
            '

            'Me.Picturebox2.Image = CType(Resources.GetObject("Picturebox2.Image"), System.Drawing.Bitmap)
            Me.PictureBox2.Location = New System.Drawing.Point(36, 10)
            Me.PictureBox2.Name = "Picturebox2"
            Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
            Me.PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
            Me.PictureBox2.TabIndex = 61
            Me.PictureBox2.TabStop = False
            Me.PictureBox2.Visible = True

            ToolTip1.SetToolTipWhenDisabled(txtPrice, "when the list price is 0 or less, the actual price can't be changed")

        End Sub

#End Region
        'Public Sub New()
        '    MyBase.new()
        '    'Me.InitializeComponent()



        'End Sub

        'Private Sub InitializeComponent()
        '    components = New System.ComponentModel.Container()
        'End Sub
        Public Property PanelType() As TypesOfPanel
            Get
                Return Me.myPanelType
            End Get
            Set(ByVal Value As TypesOfPanel)
                Me.myPanelType = Value
                Me.txtPrice.Width = 72
                If Me.myPanelType = TypesOfPanel.AdditionalTime Then Me.txtPrice.Width = 45
                Me.reDraw(Me.blnDisableAll)
            End Set
        End Property

        Public Property ListPrice() As String
            Get
                Return Me.txtListPrice.Text
            End Get
            Set(ByVal Value As String)
                Me.txtListPrice.Text = Value
            End Set
        End Property

        Public Property Discount() As String
            Get
                Return Me.txtDiscount.Text
            End Get
            Set(ByVal Value As String)
                Me.txtDiscount.Text = Value
            End Set
        End Property

        Public Property BasePanel() As DiscountPanelBase
            Get
                Return Me.myBasePanel
            End Get
            Set(ByVal Value As DiscountPanelBase)
                myBasePanel = Value
            End Set
        End Property

        Public Property DiscountScheme() As String
            Get
                Return Me.myBasePanel.DiscountScheme
            End Get
            Set(ByVal Value As String)
                myBasePanel.DiscountScheme = Value
            End Set
        End Property


        Public Property Scheme() As String
            Get
                Return Me.txtScheme.Text
            End Get
            Set(ByVal Value As String)
                Me.txtScheme.Text = Value
            End Set
        End Property


        Public Property Client() As Intranet.intranet.customerns.Client
            Get
                Return myClient
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.Client)
                myClient = Value
                'reDraw(True)
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Redraw the contents of the discount panel
        ''' </summary>
        ''' <param name="DisableAll"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action>Handle enabled controls</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function reDraw(ByVal DisableAll As Boolean) As Boolean
            objService = Nothing
            blnDisableAll = DisableAll
            'redrawing = True
            If Not myClient Is Nothing Then
                Dim blnSchemeFound As Boolean = False
                'Check for volume discounts 
                'Medscreen.LogTiming("           Fill Panel - Client - Vol disc")

                If Me.myClient.InvoiceInfo.VolumeDiscountSchemes.Count > 0 Then
                    Dim objID As Intranet.intranet.customerns.CustomerVolumeDiscount
                    For Each objID In Me.myClient.InvoiceInfo.VolumeDiscountSchemes
                        If Me.PanelType = TypesOfPanel.Sample Then
                            If objID.DiscountScheme.TypeInScheme("SAMPLE") Then
                                blnSchemeFound = True
                                Exit For
                            End If
                        ElseIf Me.PanelType = TypesOfPanel.MinimumCharge Then
                            If objID.DiscountScheme.TypeInScheme("MINCHARGE") Then
                                blnSchemeFound = True
                                Exit For
                            End If
                        ElseIf Me.PanelType = TypesOfPanel.CollectionCharge Then
                            If objID.DiscountScheme.TypeInScheme("COLLECTION") Then
                                blnSchemeFound = True
                                Exit For
                            End If

                        End If

                    Next
                    'if there are volume discount schemes show an alert
                    If blnSchemeFound Then
                        Me.PictureBox1.Show()
                    Else
                        Me.PictureBox1.Hide()
                    End If
                Else
                    Me.PictureBox1.Hide()
                End If
                Dim services As String() = Service.Split(New Char() {","})
                Dim k As Integer
                'Medscreen.LogTiming("           Fill Panel - Client - Services - check")
                For k = 0 To services.Length - 1
                    objService = Me.Client.Services.Item(services(k))
                    If Not objService Is Nothing Then Exit For
                Next
                'Medscreen.LogTiming("           Fill Panel - Client - Services - check done")

                'Sort out list price of line item; check that the item is valid
                Dim dblPrice As Double
                If Me.LineItem.Trim.Length > 0 Then
                    dblPrice = Client.InvoiceInfo.Price(LineItem)
                    Me.txtListPrice.Text = dblPrice
                    'We do have a few zero prices, if there are any the discount can't be set as it would mean an infinite discount.

                Else
                    Me.txtListPrice.Text = ""
                End If
                'Medscreen.LogTiming("           Fill Panel - Client - get discount")

                Dim myDiscount As Double
                myDiscount = CDbl(Client.InvoiceInfo.Discount(LineItem, Me.PriceDate))
                'Medscreen.LogTiming("           Fill Panel - Client - get discount price")
                oldPrice = Client.InvoiceInfo.DiscountedPrice(LineItem, Me.PriceDate)
                Me.txtDiscount.Text = myDiscount.ToString("0.00")
                txtDiscount.Enabled = False '(Not (Client.InvoiceInfo.HasStdPricing _
                'Or blnDisableAll)) AndAlso (Not objService Is Nothing)
                'Medscreen.LogTiming("           Fill Panel - Client - get scheme")
                Me.txtScheme.Text = ""
                Me.txtPrice.Text = oldPrice
                'if the list price is zero or negative! disable the new price box
                If (dblPrice > 0) Then
                    Me.txtPrice.Enabled = True
                    Me.ToolTip1.SetToolTip(txtPrice, "Set new Price")

                Else
                    Me.txtPrice.Enabled = False
                    Me.ToolTip1.SetToolTip(txtPrice, "List Price <= 0, can't be changed")
                    Me.txtListPrice.Text = "-"
                    Me.ToolTip1.SetToolTipWhenDisabled(txtListPrice, "List Price is 0 this will have derived from price book")
                End If
                If Math.Abs(myDiscount) <> 0 Then 'We wlll only have text if we have a discount
                    Me.txtScheme.Text = Client.InvoiceInfo.DiscountScheme(LineItem)
                    Me.DiscountScheme = Client.InvoiceInfo.DiscountSchemeList(LineItem)
                End If
                'objTextBox.Tag = LineItem
                Me.txtPrice.Enabled = (Not (Client.InvoiceInfo.HasStdPricing Or Me.blnDisableAll)) AndAlso (Not objService Is Nothing) AndAlso dblPrice > 0
                'Me.ToolTip1.SetToolTip(Me.txtPrice, Client.InvoiceInfo.DiscountSchemeDescription(LineItem))
                'Medscreen.LogTiming("           Fill Panel - Client - get description")

                Me.SchemeLabel.Text = Client.InvoiceInfo.DiscountSchemeDescription(LineItem)

                If Me.PanelType = TypesOfPanel.AdditionalTime Then
                    Me.lblPricephr.Text = (Client.InvoiceInfo.DiscountedPrice(Me.LineItem, Me.PriceDate) * 60).ToString("0.00")

                End If
            End If
                BackColor = SystemColors.Control
                'Set panel color dependent on whether we have the service or not but not for minimum charge types

                If Me.Enabled Then
                    If Me.PanelType = TypesOfPanel.Sample Or _
                    Me.PanelType = TypesOfPanel.AdditionalTime Or _
                    Me.PanelType = TypesOfPanel.Cancellation Or _
                    Me.PanelType = TypesOfPanel.CollectionCharge Or _
                    Me.PanelType = TypesOfPanel.MinimumCharge Then
                        If objService Is Nothing Then
                            BackColor = Color.Beige
                        Else
                            If objService.Removed Then
                                BackColor = Color.CornflowerBlue
                            ElseIf (objService.EndDate > DateField.ZeroDate And objService.EndDate < Now) Then
                                'Expired
                                BackColor = Color.CadetBlue
                            ElseIf (objService.StartDate > DateField.ZeroDate And objService.StartDate > Now) Then
                                'in future
                                BackColor = Color.Orange
                            Else
                                BackColor = Color.LightGreen
                            End If

                        End If
                    Else
                        BackColor = SystemColors.Control
                    End If
                Else
                    BackColor = SystemColors.Control

                End If
                'redrawing = False
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Allow event to fire
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [14/07/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property AllowLeave() As Boolean
            Get
                Return Not myBasePanel.AllowLeave
            End Get
            Set(ByVal Value As Boolean)
                myBasePanel.AllowLeave = Not Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Service supplied
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [16/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property PanelService() As Intranet.intranet.customerns.CustomerService
            Get
                Return objService
            End Get
            Set(ByVal Value As Intranet.intranet.customerns.CustomerService)
                objService = Value
            End Set
        End Property

        Public Property Tooltip() As System.Windows.Forms.ToolTip
            Get
                Return Me.ToolTip1
            End Get
            Set(ByVal value As System.Windows.Forms.ToolTip)
                ToolTip1 = value
            End Set
        End Property

        Public Property LineItem() As String
            Get
                Return Me.myBasePanel.LineItem
            End Get
            Set(ByVal Value As String)
                myBasePanel.LineItem = Value
            End Set
        End Property



        Public Property PriceEnabled() As Boolean
            Get
                Return Me.txtPrice.Enabled
            End Get
            Set(ByVal Value As Boolean)
                Me.txtPrice.Enabled = Value
            End Set
        End Property

        Private Sub txtPrice_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPrice.Leave
            Dim newPrice As Double
            Try
                newPrice = CDbl(Me.txtPrice.Text)
                If newPrice <> oldPrice Then
                    If Not myBasePanel.AllowLeave Then
                        myBasePanel.AllowLeave = True
                        oldPrice = newPrice
                        RaiseEvent PriceChanged(Me, Me.txtPrice.Text)
                    End If
                End If
                Me.reDraw(False)
            Catch ex As Exception
                Me.txtPrice.Text = oldPrice.ToString("0.00")
            End Try

        End Sub

        Private Sub DiscountPanel_EnabledChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.EnabledChanged
            If Me.Enabled = False Then
                BackColor = SystemColors.Control
            Else
                If Me.PanelType = TypesOfPanel.Sample Or _
Me.PanelType = TypesOfPanel.AdditionalTime Or _
Me.PanelType = TypesOfPanel.Cancellation Or _
Me.PanelType = TypesOfPanel.CollectionCharge Or _
Me.PanelType = TypesOfPanel.MinimumCharge Then
                    If objService Is Nothing Then
                        BackColor = Color.Beige
                    Else
                        If objService.Removed Then
                            BackColor = Color.CornflowerBlue
                        Else
                            BackColor = Color.LightGreen
                        End If

                    End If
                Else
                    BackColor = SystemColors.Control
                End If

            End If
        End Sub
    End Class
End Namespace