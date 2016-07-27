Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop
Public Class ExcelSupport
    Private Shared myExcelApp As Excel.Application
    Private Shared myWorkSheet As Excel.Worksheet
    Private Shared myWorkBook As Excel.Workbook

    Public Shared Sub Quit(Optional ByVal SaveChanges As Boolean = False)
        If myExcelApp IsNot Nothing Then
            If myWorkSheet IsNot Nothing Then myWorkSheet = Nothing
            If myWorkBook IsNot Nothing Then myWorkBook.Close(SaveChanges)
            myExcelApp.Quit()
        End If
    End Sub

    Private Shared Sub GetApp()
        If myExcelApp Is Nothing Then
            myExcelApp = New Excel.Application
        End If
    End Sub

    Public Shared Property Visible() As Boolean
        Get
            If Not myExcelApp Is Nothing Then
                Return myExcelApp.Visible
            End If
        End Get
        Set(ByVal value As Boolean)
            If Not myExcelApp Is Nothing Then
                myExcelApp.Visible = value
            End If
        End Set
    End Property
    Public Shared Function Open(ByVal Document As String) As Boolean
        Dim blnRet As Boolean = True
        Try
            GetApp()
            myWorkBook = myExcelApp.Workbooks.Open(Document)
            myWorkSheet = myWorkBook.Worksheets(1)
        Catch
            blnRet = False
        End Try
        Return blnRet
    End Function

    Public Shared Function CountRows(ByVal StartRow As String) As Integer
        Dim intRet As Integer = 0
        Dim objRange As Range
        myColDisp = 0
        myRowDisp = 0
        objRange = myWorkSheet.Range(RelativeAddress(StartRow))

        While objRange.Text.trim.length > 0
            intRet += 1
            myRowDisp += 1
            objRange = myWorkSheet.Range(RelativeAddress(StartRow))
        End While
        Return intRet
    End Function

    Public Shared Function GetCell(ByVal Address As String, ByVal RowOffset As Integer, ByVal ColOffset As Integer) As String
        Dim strRet As String = ""
        Dim objRange As Range
        myColDisp = ColOffset
        myRowDisp = RowOffset
        objRange = myWorkSheet.Range(RelativeAddress(Address))
        Return objRange.Text

    End Function

#Region "Support Functions"
    Private Shared myColDisp As Integer
    Private Shared myRowDisp As Integer
    Public Shared Function RelativeAddress(ByVal Address As String) As String
        'First split the address into column and Row
        Dim strCol As String = ""
        Dim i As Integer = 0
        While Char.IsLetter(Address(i))
            strCol += Address(i)
            i += 1
        End While
        Dim strRow As String = 0
        For k As Integer = i To Address.Length - 1
            strRow += Address(k)
        Next
        Dim intCol As Integer = Asc(strCol) - 65 + myColDisp
        Dim intRow As Integer = CInt(strRow) + myRowDisp
        Dim retStr As String = ""
        If intCol > 26 Then
            Dim Remand As Integer
            Dim Mult As Integer = Math.DivRem(intCol, 26, Remand)
            retStr = Chr(Mult + 64) & Chr(Remand + 65) & CStr(intRow).Trim

        Else
            retStr = Chr(intCol + 65) & CStr(intRow).Trim
        End If

        Return retStr

    End Function

    
#End Region
End Class
