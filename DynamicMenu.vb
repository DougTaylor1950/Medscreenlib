Imports System.Windows.Forms
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : DynamicMenu
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A class that provides menus built from an XML file 
''' </summary>
''' <remarks>
''' &lt;MenuItem id="Set a&amp;greed costs" OnClick="mnuCollAgreedCosts_" Tag="M"&gt;
''' is a typical entry the id will be the menu text, the onclick is the menu handler, note the trailing _, 
''' the actual signature in the code should be of the form 'onclick_Click'
''' </remarks>
''' <history>
''' 	[Taylor]	26/11/2008	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class MenuHandle

    Public Sub New(ByVal aHandle As Integer, ByVal anId As String)
        MyBase.New()
        Handle = aHandle
        ID = anId
    End Sub

    Private myHandle As Integer
    Public Property Handle() As Integer
        Get
            Return myHandle
        End Get
        Set(ByVal value As Integer)
            myHandle = value
        End Set
    End Property


    Private myId As String
    Public Property ID() As String
        Get
            Return myId
        End Get
        Set(ByVal value As String)
            myId = value
        End Set
    End Property



End Class

Public MustInherit Class DynamicMenuBase

#Region "Declarations"
    'Create a main menu object.
    Protected mainMenu As Object
    'Object for loading XML File
    Protected objXML As Xml.XmlDocument
    ' Create menu item objects.
    Protected mItem As New TaggedMenuItem()
    'Menu handle that should be returned
    Protected objMenu As Menu
    'Path of the XML Menu Configuration File 
    Protected XMLMenuFile As String
    'Form Object in which Menu has to be build
    Protected objForm As Object
    Protected HandleArray As New Generic.List(Of TaggedMenuItem)
#End Region

#Region "Public Instance"

#Region "Functions"
    Public Function LoadDynamicMenu() As Menu
        Dim oXmlElement As Xml.XmlElement
        Dim objNode As Xml.XmlNode

        objXML = New Xml.XmlDocument()
        'load the XML File
        objXML.Load(XMLMenuFile)
        'Get the documentelement of the XML file.
        oXmlElement = CType(objXML.DocumentElement, Xml.XmlElement)
        'loop through the each Top level nodes
        'For ex., File & Edit becomes Top Level nodes
        'And File -> Open , File ->Save will be treated as 
        'child for the Top Level Nodes
        For Each objNode In objXML.FirstChild.ChildNodes
            'Create a New MenuItem for Top Level Nodes
            mItem = New TaggedMenuItem()
            Dim sType As String = objNode.Attributes("type").Value
            If sType Is Nothing OrElse sType.Trim.Length = 0 Then
                ' Set the caption of the menu items
                mItem.Text = objNode.Attributes("id").Value

            ElseIf sType.ToUpper = "BAR" Then
                mItem.Text = "-"                                    'Only going to deal with bar types at the moment
            Else
                mItem.Text = objNode.Attributes("id").Value
            End If
            If Not objNode.Attributes("Name") Is Nothing Then
                mItem.ID = objNode.Attributes("Name").Value
                mItem.Name = mItem.ID
            End If

            If Not objNode.Attributes("OnClick") Is Nothing Then
                Dim sMenu As MenuItem = New MenuItem()
                ' Set the caption of the menu items.
                sMenu.Text = objNode.Attributes("id").Value

                FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNode.Attributes("OnClick").Value)
            End If
            If Not objNode.Attributes("OnSelect") Is Nothing Then
                Dim sMenu As MenuItem = New MenuItem()
                ' Set the caption of the menu items.
                sMenu.Text = objNode.Attributes("id").Value

                FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNode.Attributes("OnSelect").Value)
            End If
            If Not objNode.Attributes("OnPopup") Is Nothing Then
                Dim sMenu As MenuItem = New MenuItem()
                ' Set the caption of the menu items.
                sMenu.Text = objNode.Attributes("id").Value

                FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNode.Attributes("OnPopup").Value)
            End If

            HandleArray.Add(mItem)


            ' Add the menu items to the main menu.
            mainMenu.MenuItems.Add(mItem)
            'Dim idString As String = objNode.Attributes("OnClick").Value
            'HandleArray.Add(New MenuHandle(mItem.Handle, idString))
            'Call this Method to generate child nodes for
            'the top level node which was added now(mItem in the above Add statement)
            GenerateMenusFromXML(objNode, mainMenu.MenuItems(mainMenu.MenuItems.Count - 1))
        Next
        'return this Menu handle to the parent Form so that
        'generated menu gets displayed in the Form
        Return mainMenu
    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'This method takes care of loading Menus based on XML file contents. 
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Protected Sub GenerateMenusFromXML(ByVal objNode As Xml.XmlNode, ByVal mItm As Object)
        'This method will be invoked in an recursive fashion
        'till all the child nodes are generated. This method
        'drills up to N-levels to generate all the Child nodes
        Dim objNod As Xml.XmlNode
        'Dim sMenu As New MenuItem()
        Dim sMenu As TaggedMenuItem

        'loop for child nodes
        For Each objNod In objNode.ChildNodes
            sMenu = New TaggedMenuItem()
            sMenu = New TaggedMenuItem()
            ' Set the caption of the menu items.
            Dim sType As String = ""
            If Not objNod.Attributes("type") Is Nothing Then
                sType = objNod.Attributes("type").Value()
            End If
            sMenu.MenuType = sType
            If sType Is Nothing OrElse CStr(sType).Trim.Length = 0 Then
                ' Set the caption of the menu items
                sMenu.Text = objNod.Attributes("id").Value
            ElseIf sType.ToUpper = "REPORT" OrElse sType.ToUpper = "REPORTHTML" OrElse sType.ToUpper = "REPORTCRYSTAL" Then
                sMenu.Text = objNod.Attributes("id").Value                                    'Only going to deal with bar types at the moment
            ElseIf sType.ToUpper = "BAR" Then
                sMenu.Text = "-"                                    'Only going to deal with bar types at the moment
            Else
                sMenu.Text = objNod.Attributes("id").Value
            End If
            ' Set the caption of the menu items.
            mItm.MenuItems.Add(sMenu)
            'Add a Event handler to the menu item added
            'this method takes care of Binding Event Name(based on the parameter from
            'from xml file) to newly added menu item.
            'for ex., Your Form Code should have a Private sub MenuItemOnClick_New even to handle
            'the click of New Menu Item
            If Not objNod.Attributes("OnClick") Is Nothing Then
                FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNod.Attributes("OnClick").Value)
            End If
            If Not objNod.Attributes("Tag") Is Nothing Then
                sMenu.MenuTag = objNod.Attributes("Tag").Value
            End If
            If Not objNod.Attributes("Name") Is Nothing Then
                sMenu.ID = objNod.Attributes("Name").Value
                sMenu.Name = sMenu.ID
                HandleArray.Add(sMenu)
            End If
            If Not objNod.Attributes("Group") Is Nothing Then
                sMenu.Group = objNod.Attributes("Group").Value
            End If
            If Not objNod.Attributes("Roles") Is Nothing Then
                sMenu.Roles = objNod.Attributes("Roles").Value
            End If
            If Not objNod.Attributes("Clients") Is Nothing Then
                sMenu.Clients = objNod.Attributes("Clients").Value
            End If
            If Not objNod.Attributes("AltText") Is Nothing Then
                sMenu.NewText = objNod.Attributes("AltText").Value
            End If
            If Not objNod.Attributes("AltStatus") Is Nothing Then
                sMenu.ChangeStatus = objNod.Attributes("AltStatus").Value
            End If
            If Not objNod.Attributes("enabled") Is Nothing Then
                sMenu.Enabled = (objNod.Attributes("enabled").Value <> "False")
            End If
            If Not objNod.Attributes("shortcut") Is Nothing Then
                sMenu.Shortcut = (objNod.Attributes("shortcut").Value)
            End If
            'call the same method to see you have any child nodes
            If Not objNod.Attributes("NeedProfile") Is Nothing Then
                sMenu.NeedProfile = (objNod.Attributes("NeedProfile").Value = "True")
            End If
            'for the particular node you have added now(above mItm)
            GenerateMenusFromXML(objNod, mItm.MenuItems(mItm.MenuItems.Count - 1))
        Next
        'assign the generated mainMenu object to objMenu - public object
        'which is to be used in the Main Form
        objMenu = mainMenu
    End Sub
#End Region

#Region "Procedures"
    Protected Sub FindEventsByName(ByVal sender As Object, _
   ByVal receiver As Object, ByVal bind As Boolean, _
   ByVal handlerPrefix As String, ByVal handlerSufix As String)
        ' Get the sender's public events.
        Dim SenderEvents() As System.Reflection.EventInfo = sender.GetType().GetEvents()
        ' Get the receiver's type and lookup its public
        ' methods matching the naming convention:
        '  handlerPrefix+Click+handlerSufix
        Dim ReceiverType As Type = receiver.GetType()
        Dim E As System.Reflection.EventInfo
        Dim Method As System.Reflection.MethodInfo
        For Each E In SenderEvents
            Method = ReceiverType.GetMethod( _
              handlerSufix & E.Name, _
              System.Reflection.BindingFlags.IgnoreCase Or _
              System.Reflection.BindingFlags.Instance Or _
              System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Public)

            If Not Method Is Nothing Then
                Dim D As System.Delegate = System.Delegate.CreateDelegate(E.EventHandlerType, receiver, Method.Name)
                If bind Then
                    'add the event handler
                    E.AddEventHandler(sender, D)
                Else
                    'you can also remove the event handler if you pass bind variable as false
                    E.RemoveEventHandler(sender, D)
                End If
            End If
        Next
    End Sub
#End Region

#Region "Properties"

#End Region
#End Region


End Class

Public Class DynamicMenu
    Inherits DynamicMenuBase
    ''''''''''''''''''''''variable declarations begins''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''variable declarations ends '''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'This method will get invoked by a parent Form.
    'And it returns Menu Object.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Sub New(ByVal MenuFile As String, ByVal aForm As Object)
        MyBase.new()
        MyBase.mainMenu = New MainMenu()
        XMLMenuFile = MenuFile
        Me.objForm = aForm
    End Sub

    

    Public Function GetMenuItem(ByVal MenuId As String) As TaggedMenuItem
        Dim iRet As Integer = 0
        Dim oHandle As TaggedMenuItem = Nothing
        For Each oHandle In HandleArray
            If oHandle.ID IsNot Nothing AndAlso oHandle.ID.ToUpper = MenuId.ToUpper Then
                Exit For
            End If
            oHandle = Nothing
        Next
        Return oHandle
    End Function

    Public Property Menu() As Menu
        Get
            Return objMenu
        End Get
        Set(ByVal value As Menu)
            objMenu = value
        End Set
    End Property

    

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'objective of this method is to find out the private event present in Form 
    'and attach the newly added menuitem to this event, this was achieved using
    'Reflection technique
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  
End Class

Public Class DynamicContextMenuStrip
    Private mainToolStrip As ContextMenuStrip
    Private XMLMenuFile As String
    'Form Object in which Menu has to be build
    Private objForm As Object
    Private objXML As Xml.XmlDocument
    Private objToolStrip As ToolStrip




    Public Sub New(ByVal MenuFile As String, ByVal aForm As Object)
        MyBase.new()
        mainToolStrip = New ContextMenuStrip
        XMLMenuFile = MenuFile
        Me.objForm = aForm
    End Sub

    Public Function LoadDynamicToolstrip() As ContextMenuStrip
        Dim oXmlElement As Xml.XmlElement
        Dim objNode As Xml.XmlNode

        objXML = New Xml.XmlDocument()
        'load the XML File
        objXML.Load(XMLMenuFile)
        'Get the documentelement of the XML file.
        oXmlElement = CType(objXML.DocumentElement, Xml.XmlElement)
        'loop through the each Top level nodes
        'For ex., File & Edit becomes Top Level nodes
        'And File -> Open , File ->Save will be treated as 
        'child for the Top Level Nodes
        For Each objNode In objXML.FirstChild.ChildNodes
            'Create a New MenuItem for Top Level Nodes
            'mItem = New MenuItem()
            ' Set the caption of the menu items.
            'mItem.Text = objNode.Attributes("id").Value
            ' Add the menu items to the main menu.
            'mainMenu.MenuItems.Add(mItem)
            'Call this Method to generate child nodes for
            'the top level node which was added now(mItem in the above Add statement)

            If Not objNode.Attributes("OnClick") Is Nothing Then
                Dim sMenu As ToolStripMenuItem = New ToolStripMenuItem()
                ' Set the caption of the menu items.
                sMenu.Text = objNode.Attributes("id").Value
                mainToolStrip.Items.Add(sMenu)
                FindEventsByName(sMenu, objForm, True, "Click", objNode.Attributes("OnClick").Value)
            End If


            GenerateToolStripFromXML(objNode, Me.mainToolStrip)
        Next
        'return this Menu handle to the parent Form so that
        'generated menu gets displayed in the Form
        Return objToolStrip
    End Function

    Private Sub GenerateToolStripFromXML(ByVal objNode As Xml.XmlNode, ByVal mItm As Object)
        'This method will be invoked in an recursive fashion
        'till all the child nodes are generated. This method
        'drills up to N-levels to generate all the Child nodes
        Dim objNod As Xml.XmlNode
        Dim sMenu As TaggedToolStripItem
        'loop for child nodes

        For Each objNod In objNode.ChildNodes
            sMenu = New TaggedToolStripItem()
            ' Set the caption of the menu items.
            Dim sType As String = ""
            If Not objNod.Attributes("type") Is Nothing Then
                sType = objNod.Attributes("type").Value()
            End If
            sMenu.MenuType = sType
            If sType Is Nothing OrElse CStr(sType).Trim.Length = 0 Then
                ' Set the caption of the menu items
                sMenu.Text = objNod.Attributes("id").Value
            ElseIf sType.ToUpper = "REPORT" OrElse sType.ToUpper = "REPORTHTML" OrElse sType.ToUpper = "REPORTCRYSTAL" Then
                sMenu.Text = objNod.Attributes("id").Value                                    'Only going to deal with bar types at the moment
            ElseIf sType.ToUpper = "BAR" Then
                sMenu.Text = "-"                                    'Only going to deal with bar types at the moment
            Else
                sMenu.Text = objNod.Attributes("id").Value
            End If
            If TypeOf mItm Is ContextMenuStrip Then
                mItm.items.add(sMenu)
            Else
                mItm.DropDownItems.Add(sMenu)
            End If

            'Add a Event handler to the menu item added
            'this method takes care of Binding Event Name(based on the parameter from
            'from xml file) to newly added menu item.
            'for ex., Your Form Code should have a Private sub MenuItemOnClick_New even to handle
            'the click of New Menu Item
            If Not objNod.Attributes("OnClick") Is Nothing Then
                If Not sType Is Nothing AndAlso (sType.ToUpper = "REPORT" OrElse sType.ToUpper = "REPORTHTML" OrElse sType.ToUpper = "REPORTCRYSTAL") Then
                    FindEventsByName(sMenu, objForm, True, "MenuItemOn", "Report_handler_")
                Else
                    FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNod.Attributes("OnClick").Value)
                End If
            End If
            If Not objNod.Attributes("Tag") Is Nothing Then
                sMenu.MenuTag = objNod.Attributes("Tag").Value
            End If
            If Not objNod.Attributes("Group") Is Nothing Then
                sMenu.Group = objNod.Attributes("Group").Value
            End If
            If Not objNod.Attributes("Clients") Is Nothing Then
                sMenu.Clients = objNod.Attributes("Clients").Value
            End If
            If Not objNod.Attributes("Roles") Is Nothing Then
                sMenu.Roles = objNod.Attributes("Roles").Value
            End If
            If Not objNod.Attributes("AltText") Is Nothing Then
                sMenu.NewText = objNod.Attributes("AltText").Value
            End If
            If Not objNod.Attributes("AltStatus") Is Nothing Then
                sMenu.ChangeStatus = objNod.Attributes("AltStatus").Value
            End If
            If Not objNod.Attributes("enabled") Is Nothing Then
                sMenu.Enabled = (objNod.Attributes("enabled").Value <> "False")
            End If
            'call the same method to see you have any child nodes
            'for the particular node you have added now(above mItm)
            If TypeOf mItm Is ContextMenuStrip Then
                GenerateToolStripFromXML(objNod, mItm.Items(mItm.Items.Count - 1))
            Else
                GenerateToolStripFromXML(objNod, mItm.DropDownItems(mItm.DropDownItems.Count - 1))
            End If
        Next
        'assign the generated mainMenu object to objMenu - public object
        'which is to be used in the Main Form
        Me.objToolStrip = mainToolStrip
    End Sub



    Private Sub FindEventsByName(ByVal sender As Object, _
 ByVal receiver As Object, ByVal bind As Boolean, _
 ByVal handlerPrefix As String, ByVal handlerSufix As String)
        ' Get the sender's public events.
        Dim SenderEvents() As System.Reflection.EventInfo = sender.GetType().GetEvents()
        ' Get the receiver's type and lookup its public
        ' methods matching the naming convention:
        '  handlerPrefix+Click+handlerSufix
        Dim ReceiverType As Type = receiver.GetType()
        Dim E As System.Reflection.EventInfo
        Dim Method As System.Reflection.MethodInfo
        For Each E In SenderEvents
            Method = ReceiverType.GetMethod( _
              handlerSufix & E.Name, _
              System.Reflection.BindingFlags.IgnoreCase Or _
              System.Reflection.BindingFlags.Instance Or _
              System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Public)
            Debug.Print(E.Name)
            If Not Method Is Nothing Then
                Dim D As System.Delegate = System.Delegate.CreateDelegate(E.EventHandlerType, receiver, Method.Name)
                If bind Then
                    'add the event handler
                    E.AddEventHandler(sender, D)
                Else
                    'you can also remove the event handler if you pass bind variable as false
                    E.RemoveEventHandler(sender, D)
                End If
            End If
        Next
    End Sub


End Class

Public Class DynamicContextMenu
    Inherits DynamicMenuBase
    ''''''''''''''''''''''variable declarations begins''''''''''''''''''''''''''''
    'Create a main menu object.
    'Private Shadows mainMenu As ContextMenu
    Private mainToolStrip As ContextMenuStrip


    ''''''''''''''''''''''variable declarations ends '''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'This method will get invoked by a parent Form.
    'And it returns Menu Object.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Public Sub New(ByVal MenuFile As String, ByVal aForm As Object)
        MyBase.new()
        mainMenu = New ContextMenu()
        XMLMenuFile = MenuFile
        Me.objForm = aForm
    End Sub


    ''' <developer>CONCATENO\Taylor</developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="menuitems"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>20-Dec-2011 13:25</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
    Public Shared Sub SetMenuRoleState(ByVal menuitems As MenuItem.MenuItemCollection)
        For Each mnuitem As TaggedMenuItem In menuitems
            If mnuitem.Roles.Trim.Length > 0 Then
                Debug.Print(mnuitem.Roles)
                Dim RoleArray As String() = mnuitem.Roles.Split(New Char() {","})
                Dim hasRole As Boolean = False
                For Each role As String In RoleArray
                    If role.Trim.Length > 0 AndAlso MedscreenLib.Glossary.Glossary.UserHasRole(role) Then
                        hasRole = True
                        Exit For
                    End If
                Next
                mnuitem.Enabled = hasRole
            End If
            SetMenuRoleState(mnuitem.MenuItems)
        Next

    End Sub



    Public Overloads Function LoadDynamicMenu() As ContextMenu
        Dim oXmlElement As Xml.XmlElement
        Dim objNode As Xml.XmlNode

        objXML = New Xml.XmlDocument()
        'load the XML File
        objXML.Load(XMLMenuFile)
        'Get the documentelement of the XML file.
        oXmlElement = CType(objXML.DocumentElement, Xml.XmlElement)
        'loop through the each Top level nodes
        'For ex., File & Edit becomes Top Level nodes
        'And File -> Open , File ->Save will be treated as 
        'child for the Top Level Nodes
        For Each objNode In objXML.FirstChild.ChildNodes
            'Create a New MenuItem for Top Level Nodes
            'mItem = New MenuItem()
            ' Set the caption of the menu items.
            'mItem.Text = objNode.Attributes("id").Value
            ' Add the menu items to the main menu.
            'mainMenu.MenuItems.Add(mItem)
            'Call this Method to generate child nodes for
            'the top level node which was added now(mItem in the above Add statement)

            If Not objNode.Attributes("OnClick") Is Nothing Then
                Dim sMenu As MenuItem = New MenuItem()
                ' Set the caption of the menu items.
                sMenu.Text = objNode.Attributes("id").Value

                MyBase.FindEventsByName(mainMenu, objForm, True, "MenuItemOn", objNode.Attributes("OnClick").Value)
            End If


            GenerateMenusFromXML(objNode, mainMenu)
        Next
        'return this Menu handle to the parent Form so that
        'generated menu gets displayed in the Form
        Return objMenu
    End Function



    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''This method takes care of loading Menus based on XML file contents. 
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private Sub GenerateMenusFromXML(ByVal objNode As Xml.XmlNode, ByVal mItm As Menu)
    '    'This method will be invoked in an recursive fashion
    '    'till all the child nodes are generated. This method
    '    'drills up to N-levels to generate all the Child nodes
    '    Dim objNod As Xml.XmlNode
    '    Dim sMenu As TaggedMenuItem
    '    'loop for child nodes

    '    For Each objNod In objNode.ChildNodes
    '        sMenu = New TaggedMenuItem()
    '        ' Set the caption of the menu items.
    '        Dim sType As String = ""
    '        If Not objNod.Attributes("type") Is Nothing Then
    '            sType = objNod.Attributes("type").Value()
    '        End If
    '        sMenu.MenuType = sType
    '        If sType Is Nothing OrElse CStr(sType).Trim.Length = 0 Then
    '            ' Set the caption of the menu items
    '            sMenu.Text = objNod.Attributes("id").Value
    '        ElseIf sType.ToUpper = "REPORT" OrElse sType.ToUpper = "REPORTHTML" OrElse sType.ToUpper = "REPORTCRYSTAL" Then
    '            sMenu.Text = objNod.Attributes("id").Value                                    'Only going to deal with bar types at the moment
    '        ElseIf sType.ToUpper = "BAR" Then
    '            sMenu.Text = "-"                                    'Only going to deal with bar types at the moment
    '        Else
    '            sMenu.Text = objNod.Attributes("id").Value
    '        End If
    '        mItm.MenuItems.Add(sMenu)
    '        'Add a Event handler to the menu item added
    '        'this method takes care of Binding Event Name(based on the parameter from
    '        'from xml file) to newly added menu item.
    '        'for ex., Your Form Code should have a Private sub MenuItemOnClick_New even to handle
    '        'the click of New Menu Item
    '        If Not objNod.Attributes("OnClick") Is Nothing Then
    '            If Not sType Is Nothing AndAlso (sType.ToUpper = "REPORT" OrElse sType.ToUpper = "REPORTHTML" OrElse sType.ToUpper = "REPORTCRYSTAL") Then
    '                FindEventsByName(sMenu, objForm, True, "MenuItemOn", "Report_handler_")
    '            Else
    '                FindEventsByName(sMenu, objForm, True, "MenuItemOn", objNod.Attributes("OnClick").Value)
    '            End If
    '        End If
    '        If Not objNod.Attributes("Tag") Is Nothing Then
    '            sMenu.MenuTag = objNod.Attributes("Tag").Value
    '        End If
    '        If Not objNod.Attributes("Group") Is Nothing Then
    '            sMenu.Group = objNod.Attributes("Group").Value
    '        End If
    '        If Not objNod.Attributes("Roles") Is Nothing Then
    '            sMenu.Roles = objNod.Attributes("Roles").Value
    '        End If
    '        If Not objNod.Attributes("AltText") Is Nothing Then
    '            sMenu.NewText = objNod.Attributes("AltText").Value
    '        End If
    '        If Not objNod.Attributes("AltStatus") Is Nothing Then
    '            sMenu.ChangeStatus = objNod.Attributes("AltStatus").Value
    '        End If
    '        'call the same method to see you have any child nodes
    '        'for the particular node you have added now(above mItm)
    '        GenerateMenusFromXML(objNod, mItm.MenuItems(mItm.MenuItems.Count - 1))
    '    Next
    '    'assign the generated mainMenu object to objMenu - public object
    '    'which is to be used in the Main Form
    '    objMenu = mainMenu
    'End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'objective of this method is to find out the private event present in Form 
    'and attach the newly added menuitem to this event, this was achieved using
    'Reflection technique
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Private Sub FindEventsByName(ByVal sender As Object, _
    ' ByVal receiver As Object, ByVal bind As Boolean, _
    ' ByVal handlerPrefix As String, ByVal handlerSufix As String)
    '    ' Get the sender's public events.
    '    Dim SenderEvents() As System.Reflection.EventInfo = sender.GetType().GetEvents()
    '    ' Get the receiver's type and lookup its public
    '    ' methods matching the naming convention:
    '    '  handlerPrefix+Click+handlerSufix
    '    Dim ReceiverType As Type = receiver.GetType()
    '    Dim E As System.Reflection.EventInfo
    '    Dim Method As System.Reflection.MethodInfo
    '    For Each E In SenderEvents
    '        Method = ReceiverType.GetMethod( _
    '          handlerSufix & E.Name, _
    '          System.Reflection.BindingFlags.IgnoreCase Or _
    '          System.Reflection.BindingFlags.Instance Or _
    '          System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Public)

    '        If Not Method Is Nothing Then
    '            Dim D As System.Delegate = System.Delegate.CreateDelegate(E.EventHandlerType, receiver, Method.Name)
    '            If bind Then
    '                'add the event handler
    '                E.AddEventHandler(sender, D)
    '            Else
    '                'you can also remove the event handler if you pass bind variable as false
    '                E.RemoveEventHandler(sender, D)
    '            End If
    '        End If
    '    Next
    'End Sub
End Class


Public Class TaggedToolStripItem
    Inherits ToolStripMenuItem

    Private myRoles As String = ""
#Region "Declarations"
    Private myTag As String
    Private myNewText As String
    Private Changed As Boolean = False
    Private myChangeStatus As String
    Private myOldText As String
    Private myGroup As String = ""
    Private myType As String = ""

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

#End Region

#Region "Properties"
    Public Property Roles() As String
        Get
            Return myRoles
        End Get
        Set(ByVal Value As String)
            myRoles = Value
        End Set
    End Property

    Public Property ChangeStatus() As String
        Get
            Return myChangeStatus
        End Get
        Set(ByVal Value As String)
            myChangeStatus = Value
        End Set
    End Property

    Public Property MenuType() As String
        Get
            Return myType
        End Get
        Set(ByVal Value As String)
            myType = Value
        End Set
    End Property

    Public Property MenuTag() As String
        Get
            Return myTag
        End Get
        Set(ByVal Value As String)
            myTag = Value
        End Set
    End Property


    Private myClients As String
    Public Property Clients() As String
        Get
            Return myClients
        End Get
        Set(ByVal value As String)
            myClients = value
        End Set
    End Property


    Public Property Group() As String
        Get
            Return myGroup
        End Get
        Set(ByVal Value As String)
            myGroup = Value
        End Set
    End Property

    Public Property NewText() As String
        Get
            Return myNewText
        End Get
        Set(ByVal Value As String)
            myNewText = Value
        End Set
    End Property
#End Region
#End Region


End Class

Public Class TaggedMenuItem
    Inherits MenuItem
    Private myTag As String
    Private myNewText As String
    Private Changed As Boolean = False
    Private myChangeStatus As String
    Private myOldText As String
    Private myGroup As String = ""
    Private myType As String = ""
    Private myRoles As String = ""


    Private myNeedProfile As Boolean = False
    Public Property NeedProfile() As Boolean
        Get
            Return myNeedProfile
        End Get
        Set(ByVal value As Boolean)
            myNeedProfile = value
        End Set
    End Property

    Public Sub ActionRoles()
        Dim Roles As String = Me.Roles
        If Roles.Trim.Length > 0 Then ' We have roles 
            Dim blnInRole As Boolean = False
            Dim RoleList As String() = Roles.Split(New Char() {","})
            Dim aRole As String
            For Each aRole In RoleList
                blnInRole = MedscreenLib.Glossary.Glossary.UserHasRole(aRole)
                If blnInRole Then Exit For
            Next
            Me.Visible = blnInRole
        End If
        Me.Enabled = True
        'do any children
        For Each tm As TaggedMenuItem In Me.MenuItems
            tm.ActionRoles()
        Next
    End Sub


    Private myId As String
    Public Property ID() As String
        Get
            Return myId
        End Get
        Set(ByVal value As String)
            myId = value
        End Set
    End Property


    Private myClient As String = ""
    Public Property Clients() As String
        Get
            Return myClient
        End Get
        Set(ByVal value As String)
            myClient = value
        End Set
    End Property


    Public Property Roles() As String
        Get
            Return myRoles
        End Get
        Set(ByVal Value As String)
            myRoles = Value
        End Set
    End Property

    Public Property MenuType() As String
        Get
            Return myType
        End Get
        Set(ByVal Value As String)
            myType = Value
        End Set
    End Property

    Public Property Group() As String
        Get
            Return myGroup
        End Get
        Set(ByVal Value As String)
            myGroup = Value
        End Set
    End Property

    Public Property NewText() As String
        Get
            Return myNewText
        End Get
        Set(ByVal Value As String)
            myNewText = Value
        End Set
    End Property


    Public Property ChangeStatus() As String
        Get
            Return myChangeStatus
        End Get
        Set(ByVal Value As String)
            myChangeStatus = Value
        End Set
    End Property

    Public Property MenuTag() As String
        Get
            Return myTag
        End Get
        Set(ByVal Value As String)
            myTag = Value
        End Set
    End Property

    Public Function ResetText() As Boolean
        If Me.Changed Then
            Me.Text = Me.myOldText
            Changed = False
        End If
    End Function

    Public Function ChangeText() As Boolean
        Me.myOldText = Me.Text
        Me.Text = Me.myNewText
        Me.Changed = True
    End Function
End Class