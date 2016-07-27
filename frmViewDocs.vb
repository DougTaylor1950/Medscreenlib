Imports System.Diagnostics
Imports System.Windows.Forms
Imports MedscreenLib
Imports Intranet.Support
Imports System.Drawing

Public Class frmViewDocs


    Private myCustomerId As String
    Private myCustomer As Intranet.intranet.customerns.Client
    Private myCollectionid As String
    Private myResultsDirectory As String = Medscreen.ResultsDirectory
    Private myDocuments As Intranet.Support.AdditionalDocumentColl
    Private myCollection As Intranet.intranet.jobs.CMJob

    Public Property CustomerId() As String
        Get
            Return myCustomerId
        End Get
        Set(ByVal value As String)
            myCustomerId = value
            If value.Trim.Length > 0 Then
                myCustomer = New Intranet.intranet.customerns.Client(myCustomerId)
                myResultsDirectory += myCustomer.SMID
                LoadTree()
            End If
        End Set
    End Property


    Public Property CollectionId() As String
        Get
            Return myCollectionid
        End Get
        Set(ByVal value As String)
            myCollectionid = value
            If value.Trim.Length > 0 Then
                myCollection = New Intranet.intranet.jobs.CMJob(myCollectionid)

                LoadTree()
            End If

        End Set
    End Property



    Private Sub frmViewDocs_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the UI
        SetUpListViewColumns()
        'LoadTree()
    End Sub

    Private Sub LoadTree()
        ' TODO: Add code to add items to the treeview


        'Dim tvRoot As TreeNode
        'Dim tvNode As TreeNode
        'SELECT docid,path,type,created_on
        'FROM(additional_document)
        'WHERE linkfield = 'Customer.Identity' AND reference = ?
        'ORDER BY created_on;

        If myCustomer IsNot Nothing Then
            ' Me.TreeView.Nodes.Clear()
            ' tvRoot = Me.TreeView.Nodes.Add(myCustomer.SMIDProfile)
            If myDocuments Is Nothing Then
                myDocuments = New Intranet.Support.AdditionalDocumentColl()
                myDocuments.LoadByCustomer(myCustomer.Identity)

            End If
            'For Each aDoc As AdditionalDocument In myDocuments
            '    Dim DocNode As New Treenodes.DocumentNodes.AdditionalDocNode(aDoc)
            '    tvRoot.Nodes.Add(DocNode)
            'Next
        ElseIf myCollection IsNot Nothing Then
            If myDocuments Is Nothing Then
                myDocuments = New Intranet.Support.AdditionalDocumentColl()
                myDocuments.LoadByCollection(myCollectionid, True)

            End If


        Else
            'tvRoot = Me.TreeView.Nodes.Add("Root")
        End If
        LoadListView()
    End Sub


    Private Sub LoadListView()
        ' TODO: Add code to add items to the listview based on the selected item in the treeview

        ListView.Items.Clear()

        If myCustomer IsNot Nothing Then
            If myDocuments Is Nothing Then
                myDocuments = New Intranet.Support.AdditionalDocumentColl()
                myDocuments.LoadByCustomer(myCustomer.Identity)

            End If
            For Each aDoc As AdditionalDocument In myDocuments
                Dim ListItem As New ListViewItems.lvAdditionalDoc(aDoc)
                ListView.Items.Add(ListItem)
            Next
        ElseIf myCollection IsNot Nothing Then
            If myDocuments Is Nothing Then
                myDocuments = New Intranet.Support.AdditionalDocumentColl()
                myDocuments.LoadByCollection(myCollectionid, True)

            End If
            For Each aDoc As AdditionalDocument In myDocuments
                Dim ListItem As New ListViewItems.lvAdditionalDoc(aDoc)
                ListView.Items.Add(ListItem)
            Next
            If MsgBox("Do you want to show COC Forms, this may take some time?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Cursor = Cursors.WaitCursor
                Dim tmpstr As String = myCollection.GetSampleList
                Dim info As String() = tmpstr.Split(New Char() {";"})  'Split table from samples
                Dim StrTable As String = info.GetValue(1)
                If StrTable.Trim.Length = 0 Then StrTable = "SAMPLE"
                Dim strBarcodes As String = info.GetValue(0)
                If strBarcodes.Trim.Length > 0 Then
                    Dim strBarcodeList As String() = strBarcodes.Split(New Char() {","})
                    Dim i As Integer
                    For i = 0 To strBarcodeList.Length - 1
                        Dim strSample As String = strBarcodeList.GetValue(i)
                        Dim strPath As String = CConnection.PackageStringList("lib_sample8.GetCOCPath", strSample)
                        Dim aSDoc As ListViewItems.lvSampleDoc
                        If strPath.Trim.Length > 0 Then
                            aSDoc = New ListViewItems.lvSampleDoc(strPath, strSample)
                            ListView.Items.Add(aSDoc)
                        Else
                            Intranet.intranet.sample.clsSamples.GetAddDoc(strSample, strPath)
                            If strPath.Trim.Length > 0 Then
                                strPath = CConnection.PackageStringList("lib_sample8.GetCOCPath", strSample)
                                If strPath.Trim.Length > 0 Then
                                    aSDoc = New ListViewItems.lvSampleDoc(strPath, strSample)
                                    ListView.Items.Add(aSDoc)
                                End If
                            End If
                        End If
                    Next
                End If
                Cursor = Cursors.Default
            End If

            Else
            End If

    End Sub

    Private Sub SetUpListViewColumns()

        ' TODO: Add code to set up listview columns
        Dim lvColumnHeader As ColumnHeader

        ' Setting Column widths applies only to the current view, so this line
        '  explicitly sets the ListView to be showing the SmallIcon view
        '  before setting the column width
        SetView(View.SmallIcon)

        lvColumnHeader = ListView.Columns.Add("Path")
        ' Set the SmallIcon view column width large enough so that the items
        '  do not overlap
        ' Note that the second and third column do not appear in SmallIcon
        '  view, so there is no need to set those while in SmallIcon view
        lvColumnHeader.Width = 60

        ' Change the view to Details and set up the appropriate column
        '  widths for the Details view differently than SmallIcon view
        SetView(View.Details)

        ' The first column needs to be slightly larger in Details view than it
        '  was for SmallIcon view, and Column2 and Column3 need explicit sizes
        '  set for Details view
        lvColumnHeader.Width = 600

        lvColumnHeader = ListView.Columns.Add("Reference")
        lvColumnHeader.Width = 70

        lvColumnHeader = ListView.Columns.Add("Created")
        lvColumnHeader.Width = 70

        lvColumnHeader = ListView.Columns.Add("Type")
        lvColumnHeader.Width = 70

        lvColumnHeader = ListView.Columns.Add("Link")
        lvColumnHeader.Width = 70

        lvColumnHeader = ListView.Columns.Add("Last Modified")
        lvColumnHeader.Width = 70


    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Close this form
        Me.Close()
    End Sub

    Private Sub ToolBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolBarToolStripMenuItem.Click
        'Toggle the visibility of the toolstrip and also the checked state of the associated menu item
        ToolBarToolStripMenuItem.Checked = Not ToolBarToolStripMenuItem.Checked
        ToolStrip.Visible = ToolBarToolStripMenuItem.Checked
    End Sub

    Private Sub StatusBarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StatusBarToolStripMenuItem.Click
        'Toggle the visibility of the statusstrip and also the checked state of the associated menu item
        StatusBarToolStripMenuItem.Checked = Not StatusBarToolStripMenuItem.Checked
        StatusStrip.Visible = StatusBarToolStripMenuItem.Checked
    End Sub

    'Change whether or not the folders pane is visible
    Private Sub ToggleFoldersVisible()
        'First toggle the checked state of the associated menu item
        FoldersToolStripMenuItem.Checked = Not FoldersToolStripMenuItem.Checked

        'Change the Folders toolbar button to be in sync
        FoldersToolStripButton.Checked = FoldersToolStripMenuItem.Checked

        ' Collapse the Panel containing the TreeView.
        Me.SplitContainer.Panel1Collapsed = Not FoldersToolStripMenuItem.Checked
    End Sub

    Private Sub FoldersToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripMenuItem.Click
        ToggleFoldersVisible()
    End Sub

    Private Sub FoldersToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FoldersToolStripButton.Click
        ToggleFoldersVisible()
    End Sub

    Private Sub SetView(ByVal View As System.Windows.Forms.View)
        'Figure out which menu item should be checked
        Dim MenuItemToCheck As ToolStripMenuItem = Nothing
        Select Case View
            Case View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
            Case View.LargeIcon
                MenuItemToCheck = LargeIconsToolStripMenuItem
            Case View.List
                MenuItemToCheck = ListToolStripMenuItem
            Case View.SmallIcon
                MenuItemToCheck = SmallIconsToolStripMenuItem
            Case View.Tile
                MenuItemToCheck = TileToolStripMenuItem
            Case Else
                Debug.Fail("Unexpected View")
                View = View.Details
                MenuItemToCheck = DetailsToolStripMenuItem
        End Select

        'Check the appropriate menu item and deselect all others under the Views menu
        For Each MenuItem As ToolStripMenuItem In ListViewToolStripButton.DropDownItems
            If MenuItem Is MenuItemToCheck Then
                MenuItem.Checked = True
            Else
                MenuItem.Checked = False
            End If
        Next

        'Finally, set the view requested
        ListView.View = View
    End Sub

    Private Sub ListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListToolStripMenuItem.Click
        SetView(View.List)
    End Sub

    Private Sub DetailsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetailsToolStripMenuItem.Click
        SetView(View.Details)
    End Sub

    Private Sub LargeIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LargeIconsToolStripMenuItem.Click
        SetView(View.LargeIcon)
    End Sub

    Private Sub SmallIconsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SmallIconsToolStripMenuItem.Click
        SetView(View.SmallIcon)
    End Sub

    Private Sub TileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TileToolStripMenuItem.Click
        SetView(View.Tile)
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.txt)|*.txt"
        OpenFileDialog.ShowDialog(Me)

        Dim FileName As String = OpenFileDialog.FileName
        ' TODO: Add code to open the file
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt"
        SaveFileDialog.ShowDialog(Me)

        Dim FileName As String = SaveFileDialog.FileName
        ' TODO: Add code here to save the current contents of the form to a file.
    End Sub

    Private Sub TreeView_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)
        ' TODO: Add code to change the listview contents based on the currently-selected node of the treeview
        LoadListView()
    End Sub
    Private myLineItem As ListViewItems.lvAdditionalDoc
    Private Sub ListView_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView.Click
        myLineItem = Nothing
        If ListView.SelectedItems.Count > 0 Then
            myLineItem = ListView.SelectedItems(0)
            SetUpBrowser()
        End If
    End Sub

    Private Sub SetUpBrowser()
        PictureBox1.Hide()
        WebBrowser1.Show()
        HScrollBar1.Hide()
        VScrollBar1.Hide()
        If myLineItem IsNot Nothing Then
            lblFilename.Text = myLineItem.AdditionalDocument.Path
            If InStr(myLineItem.AdditionalDocument.Path.ToUpper, ".PDF") > 0 Then
                Me.WebBrowser1.Navigate("file:" & myLineItem.AdditionalDocument.Path)
            ElseIf InStr(myLineItem.AdditionalDocument.Path.ToUpper, ".DOC") > 0 Then
                Dim WS As New WordSupport()
                WS.OpenDoc(myLineItem.AdditionalDocument.Path)
                WS.Visible = False
                Dim oldPath As String = myLineItem.AdditionalDocument.Path
                Dim DocPath As String = IO.Path.GetDirectoryName(oldPath) & "\"
                Dim newPath As String = Medscreen.ReplaceString(myLineItem.AdditionalDocument.Path, ".Doc", ".pdf")
                Dim DocCreated As Date = IO.File.GetCreationTime(oldPath)
                Dim exportSuccess As Boolean = False
                If Not IO.File.Exists(newPath) Then     'Attempt to save as PDF
                    exportSuccess = WS.ExportToPDF(DocPath, newPath)
                End If
                Try
                    WS.Close()
                    If exportSuccess Then
                        myLineItem.AdditionalDocument.Path = newPath
                        myLineItem.AdditionalDocument.CreatedOn = DocCreated
                        myLineItem.AdditionalDocument.ModifiedOn = Now
                        myLineItem.AdditionalDocument.ModifiedBy = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
                        myLineItem.AdditionalDocument.Update()
                        IO.File.Delete(oldPath)
                        myLineItem.Refresh()
                        Threading.Thread.Sleep(1000)
                        myLineItem.Refresh()
                    End If
                Catch ex As Exception
                End Try
                Me.WebBrowser1.Navigate("file:" & myLineItem.AdditionalDocument.Path)
            ElseIf InStr(myLineItem.AdditionalDocument.Path.ToUpper, ".TIF") > 0 Then
                VScrollBar1.Show()
                HScrollBar1.Show()
                AutosizeImage(myLineItem.AdditionalDocument.Path, PictureBox1, PictureBoxSizeMode.CenterImage)
                PictureBox1.Show()
                WebBrowser1.Hide()
            Else
                Me.WebBrowser1.Navigate("file:" & myLineItem.AdditionalDocument.Path)
                Dim strGETDOCUMENTXML As String = MedscreenCommonGUIConfig.NodePLSQL("GETDOCUMENTXML")
                Dim strXML As String = CConnection.PackageStringList(strGETDOCUMENTXML, myLineItem.AdditionalDocument.Docid)
                Dim strhtml As String = Medscreen.ResolveStyleSheet(strXML, "CustDoc.xsl", 0)
                Me.wbXML.DocumentText = strhtml
            End If

        End If

    End Sub

    Private Sub ListView_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView.DoubleClick
        SetUpBrowser()
    End Sub


    Private Function DoAutosizeImage(ByVal imgOrg As Bitmap) As Bitmap
        Dim imgShow As Bitmap
        Dim g As Graphics
        Dim divideBy, divideByH, divideByW As Double

        divideByW = imgOrg.Width / (PictureBox1.Width * ScaleFactor)
        divideByH = imgOrg.Height / (PictureBox1.Height * ScaleFactor)
        'If divideByW > 1 Or divideByH > 1 Then
        If divideByW > divideByH Then
            divideBy = divideByW
        Else
            divideBy = divideByH
        End If

        imgShow = New Bitmap(CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy))
        imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
        g = Graphics.FromImage(imgShow)
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        g.DrawImage(imgOrg, New Rectangle(0, 0, CInt(CDbl(imgOrg.Width) / divideBy), CInt(CDbl(imgOrg.Height) / divideBy)), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
        g.Dispose()
        'Else
        'imgShow = New Bitmap(imgOrg.Width, imgOrg.Height)
        'imgShow.SetResolution(imgOrg.HorizontalResolution, imgOrg.VerticalResolution)
        'g = Graphics.FromImage(imgShow)
        'g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        'g.DrawImage(imgOrg, New Rectangle(0, 0, imgOrg.Width, imgOrg.Height), 0, 0, imgOrg.Width, imgOrg.Height, GraphicsUnit.Pixel)
        'g.Dispose()
        'End If
        imgOrg.Dispose()
        Return imgShow
    End Function

    Dim ScaleFactor As Double = 1
    Dim BaseImage As Bitmap
    Private myImagePath As String
    Public Sub AutosizeImage(ByVal ImagePath As String, ByVal picBox As PictureBox, Optional ByVal pSizeMode As PictureBoxSizeMode = PictureBoxSizeMode.CenterImage)
        Try
            picBox.Image = Nothing
            picBox.SizeMode = PictureBoxSizeMode.Normal
            If System.IO.File.Exists(ImagePath) Then
                HScrollBar1.Value = 0
                VScrollBar1.Value = 0
                Dim imgOrg As Bitmap
                myImagePath = ImagePath
                BaseImage = Bitmap.FromFile(ImagePath)
                imgOrg = DirectCast(Bitmap.FromFile(ImagePath), Bitmap)
                'BaseImage = imgOrg

                

                picBox.Image = DoAutosizeImage(imgOrg)
            Else
                picBox.Image = Nothing
            End If

            If PictureBox1.Height > PictureBox1.Image.Height Then
                VScrollBar1.Visible = False
            End If

            If PictureBox1.Width > PictureBox1.Image.Width Then
                HScrollBar1.Visible = False
            End If

        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

    End Sub

    Private Sub HScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles HScrollBar1.Scroll
        Dim gphPictureBox As Graphics = PictureBox1.CreateGraphics()
        gphPictureBox.DrawImage(PictureBox1.Image, New Rectangle(0, 0, _
            PictureBox1.Width - HScrollBar1.Height, _
            PictureBox1.Height - VScrollBar1.Width), _
            New Rectangle(HScrollBar1.Value * ScaleFactor, VScrollBar1.Value * ScaleFactor, _
            PictureBox1.Width - HScrollBar1.Height, _
            PictureBox1.Height - VScrollBar1.Width), GraphicsUnit.Pixel)
    End Sub

    Private Sub VScrollBar1_Scroll(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ScrollEventArgs) Handles VScrollBar1.Scroll
        Dim gphPictureBox As Graphics = PictureBox1.CreateGraphics()
        gphPictureBox.DrawImage(PictureBox1.Image, New Rectangle(0, 0, _
            PictureBox1.Width - HScrollBar1.Height, _
            PictureBox1.Height - VScrollBar1.Width), _
            New Rectangle(HScrollBar1.Value * ScaleFactor, VScrollBar1.Value * ScaleFactor, _
            PictureBox1.Width - HScrollBar1.Height, _
            PictureBox1.Height - VScrollBar1.Width), GraphicsUnit.Pixel)

    End Sub

    Private Sub cmdRotate90_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRotate90.Click
        BaseImage = Bitmap.FromFile(myImagePath)

        Dim bm_out As New Bitmap(BaseImage)
        bm_out.RotateFlip(RotateFlipType.Rotate90FlipNone)
        PictureBox1.Image = DoAutosizeImage(bm_out)
        SetScrollBars()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim strText As String = ComboBox1.Text
        Dim intEnd As Integer = InStr(strText, "%")
        If intEnd > 0 Then
            strText = Mid(strText, 1, intEnd - 1)
            ScaleFactor = CDbl(strText) / 100
            BaseImage = Bitmap.FromFile(myImagePath)

            PictureBox1.Image = DoAutosizeImage(BaseImage)
            SetScrollBars()
        End If
    End Sub

    Private Sub PictureBox1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If myImagePath Is Nothing Then Exit Sub
        BaseImage = Bitmap.FromFile(myImagePath)

        PictureBox1.Image = DoAutosizeImage(BaseImage)
    End Sub

    Private Sub cmdRotate180_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRotate180.Click
        BaseImage = Bitmap.FromFile(myImagePath)

        Dim bm_out As New Bitmap(BaseImage)
        bm_out.RotateFlip(RotateFlipType.Rotate180FlipNone)
        PictureBox1.Image = DoAutosizeImage(bm_out)
        SetScrollBars()
    End Sub

    Private Sub cmdRotate270_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRotate270.Click
        BaseImage = Bitmap.FromFile(myImagePath)

        Dim bm_out As New Bitmap(BaseImage)
        bm_out.RotateFlip(RotateFlipType.Rotate270FlipNone)
        PictureBox1.Image = DoAutosizeImage(bm_out)
        
        SetScrollBars()
    End Sub
    Private Sub SetScrollBars()
        VScrollBar1.Value = 0
        HScrollBar1.Value = 0
        If PictureBox1.Height > PictureBox1.Image.Height Then
            VScrollBar1.Visible = False
        Else
            VScrollBar1.Show()
        End If

        If PictureBox1.Width > PictureBox1.Image.Width Then
            HScrollBar1.Visible = False
        Else
            HScrollBar1.Show()
        End If
    End Sub
    ''' <developer>CONCATENO\Taylor</developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>13-Dec-2011 10:07</created><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
    Private Sub cmdUpright_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpright.Click
        BaseImage = Bitmap.FromFile(myImagePath)

        PictureBox1.Image = DoAutosizeImage(BaseImage)
        VScrollBar1.Value = 0
        HScrollBar1.Value = 0
        SetScrollBars()
    End Sub
End Class
