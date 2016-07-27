Imports System.Drawing
Imports System.Windows.Forms
Imports MedscreenLib
Public Class WebBrowserControl

#Region "Declarations"
    Private MousePoint As Point
    Private WithEvents myCache As WebCache

    Public Event Navigating(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatingEventArgs)
    Public Event Refreshing(ByVal HyperLink As String)
#End Region

#Region "Public Instance"
#Region "Properties"
    Property DocumentText() As String
        Get
            Return Me.WebBrowserEx1.DocumentText
        End Get
        Set(ByVal value As String)
            Me.WebBrowserEx1.DocumentText = value
            Cursor = Cursors.Default
        End Set
    End Property

    Public ReadOnly Property Document() As System.Windows.Forms.HtmlDocument
        Get
            Return WebBrowserEx1.Document
            Cursor = Cursors.Default
        End Get
      
    End Property

    Public ReadOnly Property Browser() As System.Windows.Forms.WebBrowser
        Get
            Return WebBrowserEx1
        End Get
    End Property

    Public Property URL() As System.Uri
        Get
            Return Me.WebBrowserEx1.Url
        End Get
        Set(ByVal value As System.Uri)

            Me.WebBrowserEx1.Url = value
            If value IsNot Nothing Then
                myCache.Add(value.ToString, "")
                DrawAddressCombo()
            End If
        End Set
    End Property

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property Cache() As WebCache
        Get
            Return myCache
        End Get
        Set(ByVal value As WebCache)
            myCache = value
        End Set
    End Property

#End Region

#Region "Procedures"
    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.myCache = New WebCache()
        Me.cmdBrowserBack.Enabled = False
        Me.CmdBrowserForward.Enabled = False

    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' Add an element to cache
    ''' </summary>
    ''' <param name="CacheElement"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Sub SetCache(ByVal CacheElement As WebCacheElement)
        myCache.Add(CacheElement)
        DrawAddressCombo()
    End Sub
#End Region

#Region "Functions"

    Public Sub Print()

        WebBrowserEx1.Print()
    End Sub


#End Region
#End Region

    Private Sub WebBrowserEx1_Navigating(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles WebBrowserEx1.Navigating
        Me.Cursor = Cursors.WaitCursor
        If InStr(e.Url.ToString.ToUpper, "WEBCOLLMANAGER") + InStr(e.Url.ToString.ToUpper, "LOCAL") = 0 Then

            RaiseEvent Navigating(sender, e)
        Else
            RaiseEvent Refreshing(e.Url.ToString)
        End If
    End Sub

    Private Sub WebBrowserEx_ProgressChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserProgressChangedEventArgs) Handles WebBrowserEx1.ProgressChanged
        Me.ToolStripProgressBar1.Maximum = Convert.ToInt32(e.MaximumProgress)
        Me.ToolStripProgressBar1.Value = Convert.ToInt32(e.CurrentProgress)
    End Sub


    Private Sub DrawAddressCombo()
        If myCache.Position > 0 Then Me.cmdBrowserBack.Enabled = True Else Me.cmdBrowserBack.Enabled = False
        If myCache.Position < myCache.Count - 1 Then Me.CmdBrowserForward.Enabled = True Else Me.CmdBrowserForward.Enabled = False
        cbAddress.Items.Clear()
        For Each wa As WebCacheElement In myCache
            If wa.HyperLink IsNot Nothing Then cbAddress.Items.Add(wa.HyperLink)
        Next
        Me.Cursor = Cursors.Default
        Me.Panel1.Cursor = Cursors.Default
    End Sub

    Private Sub WebBrowserEx1_WebBrowserMouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.HtmlElementEventArgs) Handles WebBrowserEx1.WebBrowserMouseMove
        MousePoint = New Point(e.MousePosition.X, e.MousePosition.Y)
        Dim CurrentElement As HtmlElement = WebBrowserEx1.CurrentDocument.GetElementFromPoint(MousePoint)
        Debug.Print(CurrentElement.TagName)
        If CurrentElement.TagName.ToUpper = "A" Then
            ToolStripStatusLabel1.Text = Medscreen.ReplaceString(CurrentElement.GetAttribute("href"), "&", "&&")
            Debug.Print(CurrentElement.InnerHtml)
            Application.DoEvents()
        ElseIf CurrentElement.TagName.ToUpper = "LINK" Then
            ToolStripStatusLabel1.Text = Medscreen.ReplaceString(CurrentElement.GetAttribute("href"), "&", "&&")
        ElseIf CurrentElement.TagName.ToUpper = "SPAN" Then
            Debug.Print(CurrentElement.ToString)
        Else
            ToolStripStatusLabel1.Text = ""

            Application.DoEvents()
        End If

    End Sub

    Private Sub EmailTextToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EmailTextToolStripMenuItem.Click
        Dim elements As Windows.Forms.HtmlElementCollection = Me.WebBrowserEx1.Document.All
        Dim i As Integer = 0
        Dim strBody As String = ""
        While elements(i).TagName <> "HTML" And i < elements.Count
            i += 1
        End While
        If i < elements.Count Then
            strBody = "<HTML>" & elements.Item(i).InnerHtml & "</HTML>"
            Dim stremail As String = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
            If stremail.Length = 0 Then stremail = "Doug.Taylor@medscreen.com"

            'If blnEditorOther Then
            Medscreen.BlatEmail("Report for User ", strBody, stremail, , True)
            'Else
            'Medscreen.BlatEmail("Report for collection " & Me.lvItemCurrent.Collection.ID, strBody, stremail, , True)
            'End If
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Me.WebBrowserEx1.ShowPrintDialog()
    End Sub

    Private Sub PrintPreviewToolStripMenuItem_click(ByVal dender As Object, ByVal e As System.EventArgs) Handles PrintPreviewToolStripMenuItem.Click
        Me.WebBrowserEx1.ShowPrintPreviewDialog()

    End Sub



    Private Sub cbAddress_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim wb As WebCacheElement
        If cbAddress.SelectedIndex > -1 AndAlso cbAddress.SelectedIndex < myCache.Count Then
            myCache.Position = (cbAddress.SelectedIndex)
            wb = myCache.GetElement
            If wb.Read.Trim.Length = 0 Then
                WebBrowserEx1.Navigate(wb.HyperLink)
            End If
        End If
    End Sub

    Private Sub cmdBrowserBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowserBack.Click
        If Not myCache Is Nothing Then
            If myCache.Position > 0 Then
                myCache.Position -= 1
                Me.WebBrowserEx1.DocumentText = myCache.GetElement.Read
                'txtNavigate.Text = myCache.GetElement.HyperLink
                'txtNavigate.Refresh()
                Application.DoEvents()
            End If
            DrawAddressCombo()
        End If
    End Sub

    Private Sub CmdBrowserForward_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdBrowserForward.Click
        If Not myCache Is Nothing Then
            If myCache.Position < myCache.Count - 1 Then
                myCache.Position += 1
                Me.WebBrowserEx1.DocumentText = myCache.GetElement.Read
                'txtNavigate.Text = myCache.GetElement.HyperLink
                'txtNavigate.Refresh()
                Application.DoEvents()
            End If
            DrawAddressCombo()
        End If

    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        If Not myCache Is Nothing Then
            If myCache.Position > 0 Then
                Cursor = Cursors.WaitCursor
                'myCache.Position -= 1
                Dim strLink As String = myCache.GetElement.HyperLink
                If InStr(strLink.ToUpper, "WEBCOLLMANAGER") > 0 Then
                    If InStr(strLink.ToUpper, "</A>") Then                  'Deal with anchors
                        strLink = Mid(strLink, 10)
                        Dim iPos As Integer = InStr(strLink, ">")
                        If iPos > 0 Then strLink = Mid(strLink, 1, iPos - 1)

                    End If
                    RaiseEvent Refreshing(strLink)
                Else
                    Me.WebBrowserEx1.Navigate(myCache.GetElement.HyperLink)
                End If
                'Me.txtNavigate.Text = myCache.GetElement.HyperLink
                'txtNavigate.Refresh()
                Application.DoEvents()
            End If
            DrawAddressCombo()
        End If

    End Sub

   
    Private Sub myCache_CacheChanged() Handles myCache.CacheChanged
        DrawAddressCombo()
    End Sub

    Private Sub cbAddress_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cbAddress.KeyDown
        If e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Return Then
            Dim uri As Uri
            Dim navtext As String = cbAddress.Text
            If InStr(navtext.ToUpper, "WWW") = 1 Then
                navtext = "HTTP://" & navtext
            ElseIf InStr(navtext.ToUpper, "WWW") = 0 Then
                navtext = "HTTP://WWW." & navtext
            End If
            Try
                WebBrowserEx1.Navigate(New Uri(navtext))
            Catch
            End Try
        End If
    End Sub

    Private Sub WebBrowserControl_Navigating(ByVal sender As Object, ByVal e As System.Windows.Forms.WebBrowserNavigatingEventArgs) Handles Me.Navigating
        Debug.Write(e.Url)
        Dim strLink As String = e.Url.ToString
        If InStr(strLink.ToUpper, "WEBCOLLMANAGER") > 0 Then
            If InStr(strLink.ToUpper, "</A>") Then                  'Deal with anchors
                strLink = Mid(strLink, 10)
                Dim iPos As Integer = InStr(strLink, ">")
                If iPos > 0 Then strLink = Mid(strLink, 1, iPos - 1)

            End If
            RaiseEvent Refreshing(strLink)
       
        End If
    End Sub
End Class
