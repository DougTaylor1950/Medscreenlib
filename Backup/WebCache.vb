''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : WebCacheElement
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class to manage elements added to the web cache
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	11/02/2008	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class WebCacheElement
    Private myHyperlink As String
    Private myFileName As String

    Public Sub New(ByVal Hyperlink As String, ByVal Text As String)
        myHyperlink = Hyperlink
        Me.Save(Text)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Save the element in Cache
    ''' </summary>
    ''' <param name="Text"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Save(ByVal Text As String) As Boolean
        myFileName = Environment.GetFolderPath(Environment.SpecialFolder.InternetCache) & "\WC" & Now.ToString("yyyyMMddhhmmss") & ".HTML"
        Dim iof As IO.StreamWriter = New IO.StreamWriter(myFileName)
        iof.WriteLine(Text)
        iof.Flush()
        iof.Close()
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get this element
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Read() As String
        Dim strReturn As String
        Try

            Dim iof As IO.StreamReader = New IO.StreamReader(myFileName)
            strReturn = iof.ReadToEnd
            iof.Close()
        Catch
            strReturn = ""
        End Try
        Return strReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Filename of the cache element 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	11/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property FileName() As String
        Get
            Return Me.myFileName
        End Get
    End Property

    Public ReadOnly Property HyperLink() As String
        Get
            Return Me.myHyperlink
        End Get
    End Property

End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : WebCache
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Manage the web cache
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[Taylor]	13/02/2008	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class WebCache
    Inherits CollectionBase

    Private myPosition As Integer = 0
    Public Sub New()
        MyBase.New()
    End Sub

    Public Enum IndexBy
        Hyperlink
        Filename
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Clear the contents of the cache
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	13/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shadows Function Clear()
        Dim i As Integer
        Dim objCacheElement As WebCacheElement
        For i = 0 To Me.Count - 1
            objCacheElement = MyBase.InnerList.Item(0)
            IO.File.Delete(objCacheElement.FileName)
            MyBase.InnerList.RemoveAt(0)
        Next
        MyBase.Clear()

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find an item by its numeric index
    ''' </summary>
    ''' <param name="index"></param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	13/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal index As Integer) As WebCacheElement
        Get
            Return CType(MyBase.InnerList.Item(index), WebCacheElement)
        End Get
        Set(ByVal Value As WebCacheElement)
            MyBase.InnerList.Item(index) = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Find the item by one of its properties
    ''' </summary>
    ''' <param name="Index">What to look for </param>
    ''' <param name="indexby">How to index </param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	13/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Default Public Property Item(ByVal Index As String, ByVal indexby As IndexBy) As WebCacheElement
        Get
            Dim objWE As WebCacheElement
            Dim i As Integer = 0
            For Each objWE In Me.List
                If indexby = indexby.Hyperlink Then
                    If objWE.HyperLink = Index Then
                        Exit For
                    End If
                Else
                    If objWE.FileName = Index Then
                        Exit For
                    End If
                End If
                i += 1
                objWE = Nothing
            Next
            If Not objWE Is Nothing Then Position = i
            Return objWE
        End Get
        Set(ByVal Value As WebCacheElement)

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add an item into the cache
    ''' </summary>
    ''' <param name="Hyperlink"></param>
    ''' <param name="Text"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	13/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal Hyperlink As String, ByVal Text As String) As WebCacheElement
        Dim i As Integer
        Dim objCacheElement As WebCacheElement
        For i = myPosition + 1 To Me.Count - 1
            objCacheElement = MyBase.InnerList.Item(myPosition + 1)
            IO.File.Delete(objCacheElement.FileName)
            MyBase.InnerList.RemoveAt(myPosition + 1)
        Next

        Dim objWebCacheElement As New WebCacheElement(Hyperlink, Text)
        MyBase.InnerList.Add(objWebCacheElement)
        myPosition = Count - 1
    End Function

    Public Property Position() As Integer
        Get
            Return myPosition
        End Get
        Set(ByVal Value As Integer)
            myPosition = Value
        End Set
    End Property

    Public Function GetElement() As WebCacheElement
        Return CType(MyBase.InnerList.Item(myPosition), WebCacheElement)
    End Function
End Class

