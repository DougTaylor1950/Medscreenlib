Imports Microsoft.Office.Interop.Word
Public Class WordSupport

    Private WordApp As ApplicationClass
    Private WordDoc As DocumentClass

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

    Public Sub Close()
        'WordDoc.Close()
        WordApp.Quit()
    End Sub

    Public Function ExportToPDF(ByVal DocPath As String, ByVal newPath As String) As Boolean
        Try
            WordDoc.Application.ChangeFileOpenDirectory(DocPath)
            'WS.Document.Application.ChangeFileOpenDirectory("\\john\live\ResultFiles\Customer\")
            WordDoc.Application.ActiveDocument.ExportAsFixedFormat(OutputFileName:=newPath _
                , ExportFormat:=WdExportFormat.wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:= _
                WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=WdExportRange.wdExportAllDocument, From:=1, To:=1, _
                Item:=WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, _
                CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True, _
                BitmapMissingFonts:=True, UseISO19005_1:=False)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    <CLSCompliant(False)> _
    Public Function OpenDoc(ByVal Path As String) As DocumentClass
        If IO.File.Exists(Path) Then
            WordDoc = Me.WordApp.Documents.Open(Path)

            Return WordDoc
        Else
            MsgBox("The document : " & Path & " Does not exist", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical)
            Return Nothing
        End If
    End Function

    Public Sub SaveAs(ByVal Path As String)
        Me.WordDoc.SaveAs(Path, WdSaveFormat.wdFormatDocument)
    End Sub
    <CLSCompliant(False)> _
    Public Function NewDocument(ByVal template As String) As DocumentClass
        Return WordApp.Documents.Add(template)
    End Function
    <CLSCompliant(False)> _
    Public Function GetTable(ByVal TableNo As Integer) As Table
        Return Me.Document.Tables.Item(TableNo)
    End Function
    <CLSCompliant(False)> _
    Public Function InsertCellContents(ByVal wordtable As Table, _
        ByVal RowNo As Integer, _
        ByVal colno As Integer, _
        ByVal value As Object) As Range
        '*****************************************************'
        '*               Declaration                         *'
        '*****************************************************'
        Dim objRange As Range = Nothing
        Dim lngStart As Long
        '*****************************************************'
        '                Implementation                      *'
        '*****************************************************'
        Try
            If RowNo > wordtable.Rows.Count Then wordtable.Rows.Add()
            If colno > wordtable.Columns.Count Then
                wordtable.Columns.Add()
            End If
            objRange = Nothing
            objRange = wordtable.Cell(RowNo, colno).Range
            lngStart = objRange.End - 1
            objRange.InsertAfter(value)
            objRange.Start = lngStart
            InsertCellContents = objRange
        Catch ex As Exception
        End Try
        Return objRange


    End Function 'InsertCellContents

    <CLSCompliant(False)> _
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
    <CLSCompliant(False)> _
    Public Property Selection() As Selection
        Get
            Return WordApp.Selection
        End Get
        Set(ByVal Value As Selection)
        End Set
    End Property
    <CLSCompliant(False)> _
    Public Property Document() As DocumentClass
        Get
            Return WordDoc
        End Get
        Set(ByVal Value As DocumentClass)
            WordDoc = Value
        End Set
    End Property

    Public Sub SetBookmark(ByVal BookmarkName As String, ByVal Text As String)
        Dim objRange As Range = Me.WordDoc.GoTo(WdGoToItem.wdGoToBookmark, , , BookmarkName)
        objRange.Text = Text
    End Sub


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
