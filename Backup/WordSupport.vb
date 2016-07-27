Imports mswordnet
Public Class WordSupport

    Private WordApp As mswordnet.ApplicationClass
    Private WordDoc As mswordnet.DocumentClass

    Public Sub New()
        WordApp = New ApplicationClass()
        WordApp.Visible = False
        WordDoc = WordApp.Documents.Add()


    End Sub

    Public Property Visible() As Boolean
        Get
            Return WordApp.Visible
        End Get
        Set(ByVal Value As Boolean)
            WordApp.Visible = Value
        End Set
    End Property

    Public Function Close()
        'WordDoc.Close()
        WordApp.Quit()
    End Function

    Public Function OpenDoc(ByVal Path As String) As mswordnet.DocumentClass
        WordDoc = Me.WordApp.Documents.Open(Path)

        Return WordDoc
    End Function

    Public Function SaveAs(ByVal Path As String)
        Me.WordDoc.SaveAs(Path, mswordnet.WdSaveFormat.wdFormatDocument)
    End Function

    Public Function NewDocument(ByVal template As String) As mswordnet.DocumentClass
        Return WordApp.Documents.Add(template)
    End Function

    Public Function GetTable(ByVal TableNo As Integer) As mswordnet.Table
        Return Me.Document.Tables.Item(TableNo)
    End Function

    Public Function InsertCellContents(ByVal wordtable As mswordnet.Table, _
        ByVal RowNo As Integer, _
        ByVal colno As Integer, _
        ByVal value As Object) As mswordnet.Range
        '*****************************************************'
        '*               Declaration                         *'
        '*****************************************************'
        Dim objRange As mswordnet.Range
        Dim lngStart As Long
        '*****************************************************'
        '                Implementation                      *'
        '*****************************************************'
        Try
            If RowNo > wordtable.Rows.Count Then wordtable.Rows.Add()
            objRange = Nothing
            objRange = wordtable.Cell(RowNo, colno).Range
            lngStart = objRange.End - 1
            objRange.InsertAfter(value)
            objRange.Start = lngStart
            InsertCellContents = objRange
        Catch ex As Exception
        End Try
        Exit Function


    End Function 'InsertCellContents


    Public ReadOnly Property Tables(ByVal index As Integer) As Table
        Get
            If WordDoc Is Nothing Then ' Document not opened 
                Return Nothing
            Else
                If index = 0 Or index > WordDoc.Tables.Count Then 'Outside range 
                    Return Nothing
                Else
                    Return WordDoc.Tables.Item(index)
                End If
            End If
        End Get
    End Property

    Public Property Selection() As mswordnet.Selection
        Get
            Return WordApp.Selection
        End Get
        Set(ByVal Value As mswordnet.Selection)
        End Set
    End Property

    Public Property Document() As mswordnet.DocumentClass
        Get
            Return WordDoc
        End Get
        Set(ByVal Value As mswordnet.DocumentClass)
            WordDoc = Value
        End Set
    End Property

    Public Function SetBookmark(ByVal BookmarkName As String, ByVal Text As String)
        Dim objRange As Range = Me.WordDoc.GoTo(mswordnet.WdGoToItem.wdGoToBookmark, , , BookmarkName)
        objRange.Text = Text
    End Function


    Public Sub New(ByVal strTemplate As String, Optional ByVal blntemplate As Boolean = True)
        WordApp = New ApplicationClass()
        WordApp.Visible = True
        If blntemplate Then
            WordDoc = WordApp.Documents.Add(strTemplate)
        Else
            WordDoc = WordApp.Documents.Open(strTemplate)
        End If
    End Sub
End Class
