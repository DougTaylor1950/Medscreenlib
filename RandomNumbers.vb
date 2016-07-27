
Imports System.Collections.Generic
Imports System.Xml.Serialization

''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
''' <revisionHistory></revisionHistory>
''' <author></author>
Public Class RandomNumbers
#Region "Declarations"
    Private myTitle As String = ""
    Private myHeader As String = ""
    Private myColumnList As List(Of RandomNumberColumn)
    Private myBlobValue As BlobManagement
    Private myBlobId As String = ""
    Private myCustomerId As String = ""
    Private myListType As ListTypes = ListTypes.Random
    Private myWordTemplate As String = "\\corp.concateno.com\medscreen\common\lab programs\dbreports\templates\RandomNumbers.Dot"
    Private myColumnGapSize As Integer = 20
#End Region

#Region "Enumerations"
    Public Enum ListTypes
        Random
        Every
    End Enum
#End Region

#Region "Public Instance"
#Region "Properties"

    ''' <summary>
    ''' Template to be used for Word Document
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property WordTemplate() As String
        Get
            Return myWordTemplate
        End Get
        Set(ByVal value As String)
            myWordTemplate = value
        End Set
    End Property



    Public Property ColumnGapSize() As Integer
        Get
            Return myColumnGapSize
        End Get
        Set(ByVal value As Integer)
            myColumnGapSize = value
        End Set
    End Property



    Public Property ListType() As ListTypes
        Get
            Return myListType
        End Get
        Set(ByVal value As ListTypes)
            myListType = value
        End Set
    End Property



    ''' <summary>
    ''' Title of the Random Number List
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>20-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Property Title() As String
        Get
            Return myTitle
        End Get
        Set(ByVal value As String)
            myTitle = value
        End Set
    End Property


    ''' <summary>
    ''' Header to be used in customer report
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>29-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Property Header() As String
        Get
            Return myHeader
        End Get
        Set(ByVal value As String)
            myHeader = value
        End Set
    End Property


    ''' <summary>
    ''' description of Columns in body
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>29-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Property Columns() As List(Of RandomNumberColumn)
        Get
            If myColumnList Is Nothing Then
                myColumnList = New List(Of RandomNumberColumn)
            End If
            Return myColumnList
        End Get
        Set(ByVal value As List(Of RandomNumberColumn))
            myColumnList = value
        End Set
    End Property


    Public WriteOnly Property CustomerId() As String
        Set(ByVal value As String)
            myCustomerId = value
            If myBlobId.Trim.Length = 0 Then
                myBlobId = GetCustomerBlobID()
                GetBlob(myBlobId)
            End If
        End Set
    End Property
#End Region

#Region "Procedures"
    Public Sub New(ByVal BlobId As String)
        MyClass.new()
        GetBlob(BlobId)
    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' retreive template from DB
    ''' </summary>
    ''' <param name="BlobId"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Private Sub GetBlob(ByVal BlobId As String)
        myBlobId = BlobId
        myBlobValue = New BlobManagement(BlobId)
        myBlobValue.Load()
        If myBlobValue.RowID.Trim.Length = 0 Then
            myBlobValue.TableName = "CUST_COLL_INFO"
            myBlobValue.FieldName = "RANDOMNUMBERFORMAT"
            myBlobValue.IsBinary = False
        End If
    End Sub
    Public Sub New()
        MyBase.New()
    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' update the value in Cust_coll_info
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Private Sub UpdateCustomer()
        If myCustomerId.Trim.Length = 0 OrElse myBlobId.Trim.Length = 0 Then Exit Sub

        CConnection.SetDBValue(myBlobId, "update cust_coll_info set RANDOMNUMBERFORMAT = ? where customer_id = ? ", 10, myCustomerId)

    End Sub

#End Region

#Region "Functions"

    ''' <developer></developer>
    ''' <summary>
    ''' Serialise class to a string
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Function SerialiseToString() As String
        Dim myWriter As IO.StringWriter = Nothing
        Dim strReturn As String = ""
        Try
            Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
            New System.Xml.Serialization.XmlSerializer(GetType(RandomNumbers))
            ' To write to a file, create a StreamWriter object.
            myWriter = New IO.StringWriter
            mySerializer.Serialize(myWriter, Me)
            strReturn = myWriter.ToString
        Catch ex As Exception

        Finally
            myWriter.Close()

        End Try
        Return strReturn
    End Function

    ''' <developer></developer>
    ''' <summary>
    ''' Write template to DB
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Cust Coll info.ramdom numbers is also set</remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Function WriteToDb() As Boolean
        Dim blnReturn As Boolean = False
        If myBlobValue IsNot Nothing Then
            myBlobValue.Blob = Medscreen.StringToByteArray(Me.SerialiseToString)
            blnReturn = myBlobValue.Save
            If blnReturn Then UpdateCustomer()
        End If
        Return blnReturn
    End Function
    ''' <developer></developer>
    ''' <summary>
    ''' Read a random number template from DB and deserialise
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Public Overloads Function ReadFromDB() As Boolean
        Dim blnReturn As Boolean = False
        If myBlobValue IsNot Nothing Then
            myBlobValue.Load()
            Try
                Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(RandomNumbers))
                If myBlobValue IsNot Nothing AndAlso myBlobValue.BlobString.Trim.Length > 0 Then
                    Try
                        Dim myFileStream As IO.StringReader = _
                        New IO.StringReader(myBlobValue.BlobString)
                        Dim Myobject As RandomNumbers = CType( _
                        mySerializer.Deserialize(myFileStream), RandomNumbers)

                        With Myobject
                            Me.Columns = .Columns
                            Me.Header = .Header
                            Me.Title = .Title
                            Me.WordTemplate = .WordTemplate
                            For Each col As RandomNumberColumn In .Columns
                                For Each sc As SubColumn In col.SubColumns
                                    sc.Column = col.Index
                                Next
                            Next
                        End With
                        blnReturn = True
                    Catch ex As Exception
                        blnReturn = False
                    End Try
                Else
                    MsgBox("No user options found")
                    blnReturn = False
                End If
            Catch ex As Exception
            End Try
        End If
        Return blnReturn
    End Function
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Private Function GetCustomerBlobID() As String
        Dim retStr As String = ""
        If myCustomerId.Trim.Length = 0 Then Return retStr : Exit Function
        retStr = CConnection.GetDBValue("Select RANDOMNUMBERFORMAT from cust_coll_info where customer_id = ?", myCustomerId, 10)
        Return retStr
    End Function

    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory>
    ''' <revision><created>30-Mar-2011</created><Author>Taylor</Author></revision>
    ''' </revisionHistory>
    Function TotalHeaderColumns() As Integer
        Dim intRet As Integer = 0
        For Each acol As RandomNumberColumn In Columns
            intRet += acol.SubColumns.Count
        Next

        Return intRet
    End Function


#End Region
#End Region

End Class

''' <summary>
''' A group of columns 
''' </summary>
''' <remarks></remarks>
''' <revisionHistory></revisionHistory>
''' <author></author>
Public Class RandomNumberColumn
#Region "Declarations"
    Private myIndex As Integer
#End Region

#Region "Enumerations"
    Public Enum DataImport
        None
        Excel
        Word
    End Enum
#End Region

#Region "Public Instance"


    Private myDataImport As DataImport = DataImport.None
    Public Property DataImportMethod() As DataImport
        Get
            Return myDataImport
        End Get
        Set(ByVal value As DataImport)
            myDataImport = value
        End Set
    End Property



    Private myFilename As String
    Public Property Filename() As String
        Get
            Return myFilename
        End Get
        Set(ByVal value As String)
            myFilename = value
        End Set
    End Property




    Private mySubColumns As List(Of SubColumn)
    Public Property SubColumns() As List(Of SubColumn)
        Get
            If mySubColumns Is Nothing Then
                mySubColumns = New List(Of SubColumn)
            End If
            Return mySubColumns
        End Get
        Set(ByVal value As List(Of SubColumn))
            mySubColumns = value
        End Set
    End Property


    Private myNumberOfDonors As Integer = 20
    Public Property NumberOfDonors() As Integer
        Get
            Return myNumberOfDonors
        End Get
        Set(ByVal value As Integer)
            myNumberOfDonors = value
        End Set
    End Property


    Private myFrom As Integer = 40

    ''' <summary>
    ''' Maximum   Random number to generate
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property From() As Integer
        Get
            Return myFrom
        End Get
        Set(ByVal value As Integer)
            myFrom = value
        End Set
    End Property


    Private myDecription As String = ""
    ''' <summary>
    ''' Description for group 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property GroupDescription() As String
        Get
            If myDecription.Trim.Length = 0 Then
                Return "Group : " & Me.Index
            Else
                Return myDecription
            End If

        End Get
        Set(ByVal value As String)
            myDecription = value
        End Set

    End Property


    ''' <summary>
    ''' Position of Column Group
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property Index() As Integer
        Get
            Return myIndex
        End Get
        Set(ByVal value As Integer)
            myIndex = value
        End Set
    End Property

    Public Sub New()
        MyBase.New()
    End Sub
#End Region
End Class

Public Class SubColumn
#Region "Declarations"
    Private myheader As String
    Private myDataStart As String
    Private myColumn As Integer
    Public Enum ColumnType
        RandomNumber
        Formula
        Index
        Dept
        Name
    End Enum
#End Region

#Region "Public instance"
#Region "Properties"
    Public Property Header() As String
        Get
            Return myheader
        End Get
        Set(ByVal value As String)
            myheader = value
        End Set
    End Property

    Public Property Column() As Integer
        Get
            Return myColumn
        End Get
        Set(ByVal value As Integer)
            myColumn = value
        End Set
    End Property

    Private myType As ColumnType
    Public Property ColType() As ColumnType
        Get
            Return myType
        End Get
        Set(ByVal value As ColumnType)
            myType = value
        End Set
    End Property



    Public Property DataStart() As String
        Get
            Return myDataStart
        End Get
        Set(ByVal value As String)
            myDataStart = value
        End Set
    End Property



#End Region
#End Region

End Class

