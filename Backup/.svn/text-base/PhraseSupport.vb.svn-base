Imports System.Windows.Forms

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : PhraseSupport
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class to allow phrases to be used for fomatting controls
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Taylor]</Author><date> [15/06/2010]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class PhraseSupport

    Private myForeColor As String = ""
    Public Property ForeColour() As String
        Get
            Return myForeColor
        End Get
        Set(ByVal value As String)
            myForeColor = value
        End Set
    End Property


    Private myBackColor As String = ""
    Public Property BackColour() As String
        Get
            Return myBackColor
        End Get
        Set(ByVal value As String)
            myBackColor = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' Image index, if a image list is being used
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myImageIndex As Integer
    Public Property ImageIndex() As Integer
        Get
            Return myImageIndex
        End Get
        Set(ByVal value As Integer)
            myImageIndex = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Name of the font
    ''' </summary>
    ''' <remarks>
    ''' this property must be not null if a font is to be used
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myFontName As String = ""
    Public Property FontName() As String
        Get
            Return myFontName
        End Get
        Set(ByVal value As String)
            myFontName = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the size of the font in points
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myFontSize As Single
    Public Property FontSize() As Single
        Get
            Return myFontSize
        End Get
        Set(ByVal value As Single)
            myFontSize = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Font is Bold
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myFontBold As Boolean
    Public Property FontBold() As Boolean
        Get
            Return myFontBold
        End Get
        Set(ByVal value As Boolean)
            myFontBold = value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Font is italic
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myFontItalic As Boolean
    Public Property FontItalic() As Boolean
        Get
            Return myFontItalic
        End Get
        Set(ByVal value As Boolean)
            myFontItalic = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Font is truckthrough
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------

    Private myFontStrikeout As Boolean
    Public Property FontStrikeout() As Boolean
        Get
            Return myFontStrikeout
        End Get
        Set(ByVal value As Boolean)
            myFontStrikeout = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Font underline property
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private myFontUnderline As Boolean
    Public Property FontUnderline() As Boolean
        Get
            Return myFontUnderline
        End Get
        Set(ByVal value As Boolean)
            myFontUnderline = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert object to XML string, will normally be saved in the GenText2 field of a Phrase
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' will only return something if one of the properties is set
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function SerialiseToXML() As String
        If Me.BackColour.Trim.Length > 0 OrElse Me.ForeColour.Trim.Length > 0 OrElse Me.FontName.Trim.Length > 0 OrElse Me.ImageIndex > 0 Then
            Dim myWriter As IO.StringWriter = Nothing
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(PhraseSupport))
                myWriter = New IO.StringWriter()
                mySerializer.Serialize(myWriter, Me)
            Catch ex As Exception

            End Try
            Return myWriter.ToString
        Else
            Return ""
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fore colour expressed as  a color object
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ForeColor() As Drawing.Color
        Get
            If ForeColour.Trim.Length > 0 Then
                Return Drawing.Color.FromName(ForeColour)
            Else
                Return Drawing.SystemColors.WindowText
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Control Back Colour exposed as a colour object
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' if no color is provided system colours control will be used
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property backColor() As Drawing.Color
        Get
            If BackColour.Trim.Length > 0 Then
                Return Drawing.Color.FromName(BackColour)
            Else
                Return Drawing.SystemColors.Window
            End If
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create a FONT
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' It is assumed that a font will at least have a name, if the name field is null then no font will be created
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetFont() As Drawing.Font
        Dim aFont As System.Drawing.Font
        Dim DefaultSize As Single = 8
        Dim FontStyle As Drawing.FontStyle = Drawing.FontStyle.Regular
        If Me.FontName.Trim.Length > 0 Then                                                 'Check to see if there is a name
            'initialise style
            If Me.FontBold Then FontStyle = FontStyle Or Drawing.FontStyle.Bold
            If Me.FontItalic Then FontStyle = FontStyle Or Drawing.FontStyle.Italic
            If Me.FontStrikeout Then FontStyle = FontStyle Or Drawing.FontStyle.Strikeout
            If Me.FontUnderline Then FontStyle = FontStyle Or Drawing.FontStyle.Underline
            'set size
            If Me.FontSize > 0 Then DefaultSize = FontSize
            'If TypeOf inobject Is Control Then
            'Dim inControl As Control = inobject
            'Else
            aFont = New Drawing.Font(Me.FontName, DefaultSize, FontStyle)
            'End If
            'If aFont Is Nothing Then Exit Function
            'End 'If
        End If
        Return aFont
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' decorate controls based on what type they are, some objects are not controls such as treenodes and listviewitems, they derive from different ancestors
    ''' </summary>
    ''' <param name="inControl"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Decorate(ByVal inControl As Object)
        Dim aFont As Drawing.Font = GetFont()
        If TypeOf inControl Is Control Then                                         'All Controls
            inControl.BackColor = Me.backColor
            inControl.ForeColor = Me.ForeColor
            If Not aFont Is Nothing Then CType(inControl, Control).Font = aFont
        ElseIf TypeOf inControl Is TreeNode Then                                    'Tree nodes
            Dim anode As TreeNode = inControl
            anode.BackColor = Me.backColor
            anode.ForeColor = Me.ForeColor
            If Me.ImageIndex > 0 Then anode.ImageIndex = Me.ImageIndex
            If Not aFont Is Nothing Then anode.NodeFont = aFont
        ElseIf TypeOf inControl Is ListViewItem Then                                'Listview Items
            Dim anItem As ListViewItem = inControl
            anItem.BackColor = Me.backColor
            anItem.ForeColor = Me.ForeColor
            If Me.ImageIndex > 0 Then anItem.ImageIndex = Me.ImageIndex
            If Not aFont Is Nothing Then anItem.Font = aFont
        End If
    End Sub


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert form an XML string back to object 
    ''' </summary>
    ''' <param name="XMLString"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function DeSerialiseFromXML(ByVal XMLString As String) As Boolean
        Dim blnReturn As Boolean = True
        Try
            'Create a serialiser and a reader which is filled by the string supplied
            Dim mySerializer As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(PhraseSupport))
            Dim myReader As New IO.StringReader(XMLString)
            'deserialise object into a temporary object
            Dim tmpPhraseSupport As PhraseSupport = CType(mySerializer.Deserialize(myReader), PhraseSupport)
            With tmpPhraseSupport                       'Populate object
                Me.BackColour = .BackColour
                Me.ForeColour = .ForeColour
                Me.ImageIndex = .ImageIndex
                Me.FontBold = .FontBold
                Me.FontItalic = .FontItalic
                Me.FontName = .FontName
                Me.FontSize = .FontSize
                Me.FontStrikeout = .FontStrikeout
                Me.FontUnderline = .FontUnderline
            End With
            tmpPhraseSupport = Nothing
        Catch
        End Try

    End Function






End Class

