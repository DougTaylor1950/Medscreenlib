Imports System.Xml.XPath
Namespace RSSFEEDS

    Public Class RSS

        Private rssSelect As String = ""
        Private rssTitle As String = "title"
        Private rssLink As String = "link"
        Private rssDescription As String = "description"
        Private rssItemDesc As String = "item"
        Private webNewsSourceID As Integer = 0
        Private rssURL As String = ""
        Private myList As New RSSCollection()

        Public Sub New(ByVal URL As String)
            rssURL = URL
            rssSelect = "/"
        End Sub

        Public Property List() As RSSCollection
            Get
                Return Me.myList
            End Get
            Set(ByVal Value As RSSCollection)
                myList = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get feed from Website
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub GetFeed()
            Try
                myList.Clear()
                ' Bring back the Feed
                Dim doc As New XPathDocument(rssURL)
                Dim nav As XPathNavigator = doc.CreateNavigator()

                Dim iter As XPathNodeIterator = nav.Select(rssSelect)
                ' Loop through the nodes
                While iter.MoveNext()
                    ' Get the data we need from the node
                    ProcessNode(iter.Current)
                End While
            Catch ex As Exception
                Console.WriteLine(ex.Message.ToString())
            End Try
        End Sub 'GetFeed

        Sub ProcessNode(ByVal lstNav As XPathNavigator)
            Dim title As String = ""
            Dim link As String = ""
            Dim description As String = ""
            Dim Item As String = ""
            ' Get the child nodes
            Dim iterNews As XPathNodeIterator = lstNav.SelectDescendants(XPathNodeType.Element, False)
            ' Loop through the child nodes
            While iterNews.MoveNext()
                ' Save the current Name
                Dim rssName As String = iterNews.Current.Name
                ' Is this the title?
                If rssName.ToUpper() = rssTitle.ToUpper() Then
                    title = iterNews.Current.Value

                End If
                ' Is this the Link?
                If rssName.ToUpper() = rssLink.ToUpper() Then
                    link = iterNews.Current.Value
                End If
                ' Is this the Description?
                If rssName.ToUpper() = rssDescription.ToUpper() Then
                    description = iterNews.Current.Value
                End If

                If rssName.ToUpper = rssItemDesc.ToUpper Then
                    Item = iterNews.Current.Value
                    Dim intPos As Integer = InStr(Item, ".")
                    Dim itemText As String
                    Dim itemURL As String
                    If intPos > 0 Then
                        itemText = Mid(Item, 1, intPos)
                        itemURL = Mid(Item, intPos + 1)
                    End If
                    Dim oRSSItem As RSSItem = New RSSItem(link, title, description)
                    Me.myList.Add(oRSSItem)

                End If
            End While
            ' Make sure the link and title are at least 10 characters long
            If link.Length > 10 And title.Length > 10 Then
                ' Update the database
                'InsertNews(webNewsSourceID, link, title, description)
            End If
        End Sub 'ProcessNode

    End Class

    Public Class RSSCollection
        Inherits CollectionBase

        Public Sub New()
            MyBase.New()
        End Sub
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an item to collection 
        ''' </summary>
        ''' <param name="Item"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal Item As RSSItem) As Integer
            Return MyBase.List.Add(Item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Item in collection 
        ''' </summary>
        ''' <param name="Index"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal Index As Integer) As RSSItem
            Get
                Return CType(MyBase.List.Item(Index), RSSItem)
            End Get
            Set(ByVal Value As RSSItem)
                MyBase.List.Item(Index) = Value
            End Set
        End Property
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : RSSFEEDS.RSSItem
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Item descripbing a RSS envelope
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	14/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class RSSItem
        Private myLink As String
        Private myDescription As String
        Private myTitle As String

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new item 
        ''' </summary>
        ''' <param name="Link"></param>
        ''' <param name="Title"></param>
        ''' <param name="Description"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Link As String, ByVal Title As String, ByVal Description As String)
            myLink = Link
            myTitle = Title
            myDescription = Description

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Link to item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Link() As String
            Get
                Return myLink
            End Get
            Set(ByVal Value As String)
                myLink = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Title of Item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Title() As String
            Get
                Return myTitle
            End Get
            Set(ByVal Value As String)
                myTitle = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description of Item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	14/02/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property Description() As String
            Get
                Return Me.myDescription
            End Get
            Set(ByVal Value As String)
                Me.myDescription = Value
            End Set
        End Property
    End Class
End Namespace