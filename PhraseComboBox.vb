Imports System.Drawing
Imports System.Windows.Forms

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 : PhraseComboBox
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' 
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[Taylor]</Author><date> [24/06/2010]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class PhraseComboBox
    Inherits ComboBox
    Private myPhraseType As String
    Private myPhraseList As MedscreenLib.Glossary.PhraseCollection
    Private StaticFont As Font
    Private Loading As Boolean = False

    Public Sub New()
        MyBase.New()
        MyBase.DrawMode = Windows.Forms.DrawMode.OwnerDrawVariable
        StaticFont = Me.Font
    End Sub

    Public Property PhraseType() As String
        Get
            Return myPhraseType
        End Get
        Set(ByVal value As String)
            Loading = True
            myPhraseType = value
            myPhraseList = New MedscreenLib.Glossary.PhraseCollection(myPhraseType)
            'Only try loading list when we have a connection 
            If Not MedscreenLib.CConnection.DbConnection Is Nothing Then
                'MedscreenLib.MedConnection.SetInstance("JOHN-LIVE")
                myPhraseList.Load("Phrase_text")
            End If

            Me.DataSource = myPhraseList
            Me.ValueMember = "PhraseID"
            Me.DisplayMember = "PhraseText"
            Loading = False
        End Set
    End Property

    Private Sub PhraseComboBox_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles MyBase.DrawItem
        If e.Index < 0 Then
            e.DrawBackground()
            e.DrawFocusRectangle()
            Exit Sub
        End If
        ' Get the Color object from the Items list
        Dim foreColor As Color = _
           CType(Me.Items(e.Index), MedscreenLib.Glossary.Phrase).Formatter.ForeColor
        Dim BackColor As Color = _
           CType(Me.Items(e.Index), MedscreenLib.Glossary.Phrase).Formatter.backColor
        Dim aPhrase As MedscreenLib.Glossary.Phrase = CType(Me.Items(e.Index), MedscreenLib.Glossary.Phrase)
        ' get a square using the bounds height
        Dim rect As Rectangle = New Rectangle _
           (2, e.Bounds.Top + 2, e.Bounds.Width - 2, _
           e.Bounds.Height - 4)

        Dim bbr As SolidBrush = New SolidBrush(BackColor)
        Dim fbr As New SolidBrush(foreColor)
        ' call these methods first
        e.DrawBackground()
        e.DrawFocusRectangle()

        ' change brush color if item is selected
        'If e.State = _
        '   Windows.Forms.DrawItemState.Selected Then
        '    br = Brushes.White
        'Else

        'End If

        ' draw a rectangle and fill it
        e.Graphics.DrawRectangle(New Pen(BackColor), rect)
        e.Graphics.FillRectangle(New SolidBrush(BackColor), rect)

        ' draw a border
        'rect.Inflate(1, 1)
        'e.Graphics.DrawRectangle(Pens.Black, rect)

        ' draw the Color name
        Dim afont As Font = aPhrase.Formatter.GetFont
        If afont Is Nothing Then
            e.Graphics.DrawString(aPhrase.PhraseText, Me.Font, fbr, e.Bounds.Height + 5, ((e.Bounds.Height - Me.Font.Height) \ 2) + e.Bounds.Top)
        Else
            e.Graphics.DrawString(aPhrase.PhraseText, afont, fbr, e.Bounds.Height + 5, ((e.Bounds.Height - Me.Font.Height) \ 2) + e.Bounds.Top)
        End If


    End Sub

    Private Sub PhraseComboBox_MeasureItem(ByVal sender As Object, ByVal e As System.Windows.Forms.MeasureItemEventArgs) Handles MyBase.MeasureItem
        'e.ItemHeight = 
    End Sub

    Private Sub PhraseComboBox_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SelectedValueChanged

        If Me.SelectedIndex > -1 Then
            If Not Loading Then
                Dim aPhrase As MedscreenLib.Glossary.Phrase = Me.SelectedItem
                Me.ForeColor = aPhrase.Formatter.ForeColor
                Me.BackColor = aPhrase.Formatter.backColor
                If Not aPhrase.Formatter.GetFont Is Nothing Then
                    Me.Font = aPhrase.Formatter.GetFont
                Else
                    Me.Font = StaticFont
                End If

            End If
        End If
    End Sub


    Private Sub PhraseComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.SelectedIndexChanged
        If Me.SelectedIndex > -1 Then
            If Not Loading Then
                Dim aPhrase As MedscreenLib.Glossary.Phrase = Me.SelectedItem
                Me.ForeColor = aPhrase.Formatter.ForeColor
                Me.BackColor = aPhrase.Formatter.backColor
                If Not aPhrase.Formatter.GetFont Is Nothing Then
                    Me.Font = aPhrase.Formatter.GetFont
                Else
                    Me.Font = StaticFont
                End If
                'Me.SelectionStart = 1
            End If
        End If
    End Sub
End Class
