'********************************************************8
'*	Author: Andrew Deren
'*	Date: July, 2004
'*	http://www.adersoftware.com
'* 
'*	StringTokenizer class. You can use this class in any way you want
'* as long as this header remains in this file.
'* 
'**********************************************************/
Imports System
Imports System.IO
Imports System.Text

Public Enum TokenKind
    Unknown
    Word
    Number
    QuotedString
    WhiteSpace
    Symbol
    EOL
    EOF
    Command
End Enum


''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
''' <revisionHistory></revisionHistory>
''' <author></author>
Public Class StringTokenizer
#Region "Declarations"
    Private Const EOF As Char = Chr(0)

    Private line As Integer
    Private column As Integer
    Private pos As Integer '	// position within data

    Private data As String
    Private mySymbolchars As Char()
    Private myIgnoreWhiteSpace As Boolean

    Private saveLine As Integer
    Private saveCol As Integer
    Private savePos As Integer

#End Region

#Region "Public Instance"

#Region "Functions"
    Protected Function LA(ByVal count As Integer) As Char

        If (pos + count >= data.Length) Then
            Return EOF
        Else
            Return data(pos + count)
        End If
    End Function

    Protected Function Consume() As Char

        Dim ret As Char = data(pos)
        pos += 1
        column += 1

        Return ret
    End Function

    Protected Overloads Function CreateToken(ByVal kind As TokenKind, ByVal value As String) As Token

        Return New Token(kind, value, line, column)
    End Function

    Protected Overloads Function CreateToken(ByVal kind As TokenKind) As Token

        Dim tokenData As String = data.Substring(savePos, pos - savePos)
        Return New Token(kind, tokenData, saveLine, saveCol)
    End Function

    Public Function NextToken() As Token

ReadToken:

        Dim ch As Char = LA(0)
        Select ch

            Case EOF
                Return CreateToken(TokenKind.EOF, String.Empty)

            Case " ", Chr(9)

                If (IgnoreWhiteSpace) Then

                    Consume()
                    GoTo ReadToken

                Else
                    Return ReadWhitespace()
                End If
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                Return ReadNumber()

            Case Chr(13)

                StartRead()
                Consume()
                If (LA(0) = Chr(10)) Then
                    Consume()   '// on DOS/Windows we have \r\n for new line

                    line += 1
                    column = 1
                End If

                Return CreateToken(TokenKind.EOL)
            Case Chr(10)

                StartRead()
                Consume()
                line += 1
                column = 1

                Return CreateToken(TokenKind.EOL)


            Case ""

                Return ReadString()


            Case Else

                If (Char.IsLetter(ch) Or ch = "_") Then
                    Return ReadWord()
                ElseIf (IsSymbol(ch)) Then

                    StartRead()
                    Consume()
                    Return CreateToken(TokenKind.Symbol)

                Else

                    StartRead()
                    Consume()
                    Return CreateToken(TokenKind.Unknown)

                End If

        End Select
    End Function


    ''' <summary>
    ''' checks whether c is a symbol character.
    ''' </summary>
    Protected Function IsSymbol(ByVal c As Char) As Boolean
        Dim MyRet As Boolean = False
        For Each sc As Char In SymbolChars
            If c = sc Then
                MyRet = True
                Exit For
            End If
        Next
        Return MyRet
    End Function


    ''' <summary>
    ''' reads all characters until next " is found.
    ''' If "" (2 quotes) are found, then they are consumed as
    ''' part of the string
    ''' </summary>
    ''' <returns></returns>
    Protected Function ReadString() As Token

        StartRead()

        Consume() '' read "

        While (True)

            Dim ch As Char = LA(0)
            If (ch = EOF) Then
                Exit While
            ElseIf (ch = Chr(13)) Then  '' handle CR in strings

                Consume()
                If (LA(0) = Chr(10)) Then    '' for DOS & windows
                    Consume()

                    line += 1
                    column = 1
                End If
            ElseIf (ch = Chr(10)) Then  '' new line in quoted string

                Consume()

                line += 1
                column = 1

            ElseIf (ch = "") Then

                Consume()
                If (LA(0) <> "") Then
                    Exit While   '' done reading, and this quotes does not have escape character
                Else
                    Consume() '' consume second ", because first was just an escape
                End If
            Else
                Consume()
            End If
        End While

        Return CreateToken(TokenKind.QuotedString)


    End Function

    ''' <summary>
    ''' reads word. Word contains any alpha character or _
    ''' </summary>
    Protected Function ReadWord() As Token

        StartRead()

        Consume() ' consume first character of the word

        While (True)

            Dim ch As Char = LA(0)
            If (Char.IsLetter(ch) Or ch = "_") Then
                Consume()
            Else
                Exit While
            End If

        End While
        Return CreateToken(TokenKind.Word)
    End Function



    ''' <summary>
    ''' reads number. Number is: DIGIT+ ("." DIGIT*)?
    ''' </summary>
    ''' <returns></returns>
    Protected Function ReadNumber() As Token

        StartRead()

        Dim hadDot As Boolean = False

        Consume() ' read first digit

        While (True)

            Dim ch As Char = LA(0)
            If (Char.IsDigit(ch)) Then
                Consume()
            ElseIf (ch = "." And Not hadDot) Then
                hadDot = True
                Consume()
            Else
                Exit While
            End If
        End While

        Return CreateToken(TokenKind.Number)
    End Function



    ''' <summary>
    ''' reads all whitespace characters (does not include newline)
    ''' </summary>
    ''' <returns></returns>
    Protected Function ReadWhitespace() As Token

        StartRead()

        Consume() ' consume the looked-ahead whitespace char

        While (True)

            Dim ch As Char = LA(0)
            If Char.IsWhiteSpace(ch) Then
                Consume()
            Else
                Exit While
            End If
        End While

        Return CreateToken(TokenKind.WhiteSpace)
    End Function


#End Region

#Region "Procedures"
    ''' <summary>
    ''' save read point positions so that CreateToken can use those
    ''' </summary>
    Private Sub StartRead()
        saveLine = line
        saveCol = column
        savePos = pos
    End Sub


    Public Sub New(ByVal reader As TextReader)

        If (reader Is Nothing) Then
            Throw New ArgumentNullException("reader")
        End If
        data = reader.ReadToEnd()

        Reset()
    End Sub

    Public Sub New(ByVal indata As String)

        If indata Is Nothing Then
            Throw New ArgumentNullException("data")
        End If
        data = indata

        Reset()
    End Sub

    Private Sub Reset()

        myIgnoreWhiteSpace = False
        mySymbolchars = New Char() {"=", "+", "-", "/", ",", ".", "*", "~", "!", "@", "#", "$", "%", "^", "&", "(", ")", "{", "}", "[", "]", ":", ";", "<", ">", "?", "|", "\\"}

        line = 1
        column = 1
        pos = 0
    End Sub


#End Region

#Region "Properties"



    Public Property IgnoreWhiteSpace() As Boolean
        Get
            Return myIgnoreWhiteSpace
        End Get
        Set(ByVal value As Boolean)
            myIgnoreWhiteSpace = value
        End Set
    End Property


    Public Property SymbolChars() As Char()
        Get
            Return mySymbolchars
        End Get
        Set(ByVal value As Char())
            mySymbolchars = value
        End Set
    End Property

#End Region
#End Region

End Class

Public Class Token
#Region "Declarations"
    Private myline As Integer
    Private mycolumn As Integer
    Private myvalue As String
    Private mykind As TokenKind


#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
    Public Sub New(ByVal kind As TokenKind, ByVal value As String, ByVal line As Integer, ByVal column As Integer)
        MyBase.new()
        mykind = kind
        myvalue = value
        myline = line
        mycolumn = column
    End Sub


#End Region

#Region "Properties"


    Public Property Kind() As TokenKind
        Get
            Return mykind
        End Get
        Set(ByVal value As TokenKind)
            mykind = value
        End Set
    End Property


    Public Property Line() As Integer
        Get
            Return myline
        End Get
        Set(ByVal value As Integer)
            myline = value
        End Set
    End Property



    Public Property Column() As Integer
        Get
            Return mycolumn
        End Get
        Set(ByVal value As Integer)
            mycolumn = value
        End Set
    End Property


    Public Property Value() As String
        Get
            Return myvalue
        End Get
        Set(ByVal value As String)
            myvalue = value
        End Set
    End Property

#End Region
#End Region

End Class

Public Class CommandLineTokeniser
    Inherits StringTokenizer
#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"
    Public Overloads Function NextToken() As Token
        Dim AToken As Token = MyBase.NextToken
        If AToken.Kind = TokenKind.Symbol AndAlso AToken.Value = "/" Then
            Dim CommandName As String = MyBase.NextToken.Value
            Dim CommandParameter As String = ""
            Dim nxt As Token = MyBase.NextToken
            If nxt.Kind = TokenKind.Symbol Then
                nxt = MyBase.NextToken
                If nxt.Kind = TokenKind.Word Then CommandParameter = nxt.Value
            End If
            AToken = New CommandToken(CommandName, CommandParameter, AToken.Line, AToken.Column)

        End If
        Return AToken
    End Function

#End Region

#Region "Procedures"
    Public Sub New(ByVal CommandLine As String)
        MyBase.New(CommandLine)
    End Sub
#End Region

#Region "Properties"

#End Region
#End Region

End Class

Public Class CommandToken
    Inherits Token
#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
    Public Sub New(ByVal CommandName As String, ByVal CommandParameter As String, ByVal Line As Integer, ByVal Column As Integer)
        MyBase.New(TokenKind.Command, CommandName, Line, Column)
        myCommandParameter = CommandParameter
    End Sub
#End Region

#Region "Properties"

    Private myCommandParameter As String
    Public Property CommandParameter() As String
        Get
            Return myCommandParameter
        End Get
        Set(ByVal value As String)
            myCommandParameter = value
        End Set
    End Property

    Public Property CommandName() As String
        Get
            Return MyBase.Value.ToUpper
        End Get
        Set(ByVal value As String)
            MyBase.Value = value
        End Set
    End Property

#End Region
#End Region

End Class