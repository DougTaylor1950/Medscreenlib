Imports System.Xml.Serialization
Imports System.Drawing
Imports System.Windows.Forms

Namespace UserDefaults

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.UserOptions
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A grouping of user options for the customer centre tool
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class UserOptions

#Region "Declarations"
        Private pDiaryStatusFilter As String = ""
        Private pDiaryTypeFilter As String = ""
        Private pTreeViewOption As Integer = 0
        Private pViewJobState As Integer = 0
        Private pViewHistoryState As Integer = 0
        Private pViewIndustry As String = ""
        Private pCollDestination As Integer = 0
        Private pShowMyCollections As Boolean = False
        Private pShowMyCustomers As Boolean = False
        Private pShowTreeToolTips As Boolean = False
        Private pSMIDComments As Boolean = True
        Private pDiaryColumn As Integer = 0
        Private pDiarySortDirection As Integer = 1
        Private pCustSelectColumn As Integer = 0
        Private pCustSelectSortDirection As Integer = 1
        Private pRefreshTime As Integer = 300
        Private pPastDays As Integer = 90
        Private pDefaultPrinter As String                       'Default printer
        Private pDefaultSend As MedscreenLib.Constants.SendMethod   'Default Send method
        Private pShowRSS As Boolean = False

        Private meDiaryColours As New UserColourCollection()
        Private myDiaryColumns As New ColumnInfoCollection()
        Private myCustSelectColumns As New ColumnInfoCollection()


        Private Const aMinute As Integer = 60000
#End Region

#Region "Functions"

#End Region

#Region "Procedures"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Send a Crystal report to the user by their default method
        ''' </summary>
        ''' <param name="objReport"></param>
        ''' <param name="Subject"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	06/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub OutputCrystalReport(ByVal objReport As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal Subject As String)
            'If we are going to print it see if a default printer has been set up 
            If DefaultSendMethod = MedscreenLib.Constants.SendMethod.Printer Then
                Dim prtdoc As New System.Drawing.Printing.PrintDocument()
                Dim strDefaultPrinter As String
                strDefaultPrinter = MedscreenCommonGui.CommonForms.UserOptions.DefaultPrinter
                'If no default printer set use windows default 
                If strDefaultPrinter.Trim.Length = 0 Then
                    strDefaultPrinter = prtdoc.PrinterSettings.PrinterName
                End If
                MedscreenLib.Medscreen.PrintReport(objReport, strDefaultPrinter)
            Else  'Okay not printing so sending by email with a possible attachment
                'Set up recipient
                Dim strRecipient As String
                If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then
                    strRecipient = MedscreenLib.Glossary.Glossary.CurrentSMUser.Email
                End If

                Dim FileName As String = MedscreenLib.Medscreen.ExportReport(objReport, MedscreenCommonGui.CommonForms.UserOptions.DefaultSendMethod)
                'if we are sending as an email then it is part of the message
                If DefaultSendMethod = MedscreenLib.Constants.SendMethod.HTML Then
                    Dim iof As IO.StreamReader
                    Try
                        iof = New IO.StreamReader(FileName)


                        Dim strHTML As String = iof.ReadToEnd
                        MedscreenLib.Medscreen.BlatEmail(Subject, strHTML, strRecipient, , True)

                    Catch ex As Exception
                    Finally
                        iof.Close()
                    End Try
                Else  ' Not email send file as an attachment
                    MedscreenLib.Medscreen.BlatEmail(Subject, "Please find your report enclosed", strRecipient, , True, , FileName)
                End If
            End If

        End Sub

#End Region

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Indiacte whether RSS info will 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	24/07/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ShowRSS() As Boolean
            Get
                Return Me.pShowRSS
            End Get
            Set(ByVal Value As Boolean)
                pShowRSS = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Collection Diary currently sorted column
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiaryColumn() As Integer
            Get
                Return pDiaryColumn
            End Get
            Set(ByVal Value As Integer)
                pDiaryColumn = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Saved industry value 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	21/10/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property CurrentIndustry() As String
            Get
                Return Me.pViewIndustry
            End Get
            Set(ByVal Value As String)
                Me.pViewIndustry = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Columns to be displayed in the Diary
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiaryColumns() As ColumnInfoCollection
            Get
                Return myDiaryColumns
            End Get
            Set(ByVal Value As ColumnInfoCollection)
                myDiaryColumns = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Columns for Customer Selection Screen
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CustSelectColumns() As ColumnInfoCollection
            Get
                Return myCustSelectColumns
            End Get
            Set(ByVal Value As ColumnInfoCollection)
                myCustSelectColumns = Value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Direction that the column in the Diary should be sorted 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiarySortDirection() As Integer
            Get
                Return pDiarySortDirection
            End Get
            Set(ByVal Value As Integer)
                pDiarySortDirection = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Filters set on the collection diary
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiaryTypeFilter() As String
            Get
                Return Me.pDiaryTypeFilter
            End Get
            Set(ByVal Value As String)
                pDiaryTypeFilter = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Customer Select dialogue, sorted column
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CustSelectColumn() As Integer
            Get
                Return pCustSelectColumn
            End Get
            Set(ByVal Value As Integer)
                pCustSelectColumn = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Direction of sorting in the customer selection dialogue
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property CustSelectSortDirection() As Integer
            Get
                Return pCustSelectSortDirection
            End Get
            Set(ByVal Value As Integer)
                pCustSelectSortDirection = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Show users collections
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ShowMyCollections() As Boolean
            Get
                Return pShowMyCollections
            End Get
            Set(ByVal Value As Boolean)
                pShowMyCollections = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Show users customers
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ShowMyCustomers() As Boolean
            Get
                Return pShowMyCustomers
            End Get
            Set(ByVal Value As Boolean)
                pShowMyCustomers = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Diary status filter (SQL in statement)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiaryStatusFilter() As String
            Get
                Return pDiaryStatusFilter
            End Get
            Set(ByVal Value As String)
                pDiaryStatusFilter = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Option for display of customers in the tree view 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property TreeViewOption() As Integer
            Get
                Return pTreeViewOption
            End Get
            Set(ByVal Value As Integer)
                pTreeViewOption = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' How to display Jobs in CCTool Tree View
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property TreeViewJobOption() As Integer
            Get
                Return pViewJobState
            End Get
            Set(ByVal Value As Integer)
                pViewJobState = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' What GoldMine History to show in CCTool Tree View
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property TreeViewHistoryOption() As Integer
            Get
                Return pViewHistoryState
            End Get
            Set(ByVal Value As Integer)
                pViewHistoryState = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default SM Printer 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property DefaultPrinter() As String
            Get
                Return Me.pDefaultPrinter
            End Get
            Set(ByVal Value As String)
                Me.pDefaultPrinter = Value
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default method of sending stuff out
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	03/08/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property DefaultSendMethod() As MedscreenLib.Constants.SendMethod
            Get
                If pDefaultSend = 0 Then
                    pDefaultSend = MedscreenLib.Constants.SendMethod.Email
                End If
                Return pDefaultSend
            End Get
            Set(ByVal Value As MedscreenLib.Constants.SendMethod)
                pDefaultSend = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Show Tool Tips in the CCTool, tree View
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ShowTreeToolTips() As Boolean
            Get
                Return Me.pShowTreeToolTips
            End Get
            Set(ByVal Value As Boolean)
                pShowTreeToolTips = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Show SMIDComments
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	04/06/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Property ShowSMIDCOmments() As Boolean
            Get
                Return Me.pSMIDComments
            End Get
            Set(ByVal Value As Boolean)
                Me.pSMIDComments = Value
            End Set
        End Property

        Public Property CollDestination() As Integer
            Get
                Return pCollDestination
            End Get
            Set(ByVal Value As Integer)
                pCollDestination = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' What Colours to use in the Diary
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DiaryColours() As UserColourCollection
            Get
                Return Me.meDiaryColours
            End Get
            Set(ByVal Value As UserColourCollection)
                meDiaryColours = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' What Refresh interval to set 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property RefreshTime() As Integer
            Get
                Return Me.pRefreshTime * aMinute
            End Get
            Set(ByVal Value As Integer)
                Me.pRefreshTime = Value / aMinute
            End Set
        End Property

        Public Property PastDays() As Integer
            Get
                Return Me.pPastDays
            End Get
            Set(ByVal Value As Integer)
                Me.pPastDays = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load these options 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Load(Optional ByVal NormalLoad As Boolean = True) As Boolean
            ' To read the file, create a FileStream object.
            Dim strFileName As String
            'If NormalLoad Then
            strFileName = Filename()
            'Else
            'strFileName = DefaultFilename()
            'End If
            'MsgBox(Filename)
            If Not IO.File.Exists(strFileName) Then
                strFileName = DefaultFilename()
                'MsgBox(strFileName)
            End If
            Return DoLoad(strFileName)
        End Function

        Public Function LoadDefaults() As Boolean
            Return DoLoad(DefaultFilename())
        End Function

        Private Function DoLoad(ByVal strFilename As String) As Boolean
            Dim blnReturn As Boolean = True
            Try
                Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(UserOptions))
                If IO.File.Exists(strFilename) Then
                    Try
                        Dim myFileStream As IO.FileStream = _
                        New IO.FileStream(strFilename, IO.FileMode.Open)
                        Dim Myobject As UserOptions = CType( _
                        mySerializer.Deserialize(myFileStream), UserOptions)

                        With Myobject
                            Me.DiaryStatusFilter = .DiaryStatusFilter
                            Me.DiaryTypeFilter = .DiaryTypeFilter
                            Me.TreeViewOption = .TreeViewOption
                            Me.ShowMyCollections = .ShowMyCollections
                            Me.ShowMyCustomers = .ShowMyCustomers
                            Me.TreeViewJobOption = .TreeViewJobOption
                            Me.TreeViewHistoryOption = .TreeViewHistoryOption
                            Me.CollDestination = .CollDestination
                            Me.DiaryColumn = .DiaryColumn
                            Me.DiarySortDirection = .DiarySortDirection
                            Me.DiaryColours = .DiaryColours
                            Me.ShowTreeToolTips = .ShowTreeToolTips
                            Me.ShowSMIDCOmments = .ShowSMIDCOmments
                            Me.RefreshTime = .RefreshTime
                            Me.DefaultSendMethod = .DefaultSendMethod
                            Me.pDefaultPrinter = .pDefaultPrinter
                            Me.pViewIndustry = .pViewIndustry
                            If Me.DiaryColours.Count = 0 Then
                                Me.DiaryColours.AddDefaults()
                            End If
                            Me.DiaryColumns = .DiaryColumns
                            If Me.DiaryColumns.Count = 0 Then
                                Me.DiaryColumns = CommonForms.ListViewColumns.LVDiaryList
                            End If
                            Me.CustSelectColumns = .CustSelectColumns
                            If Me.CustSelectColumns.Count = 0 Then
                                Me.CustSelectColumns = CommonForms.ListViewColumns.LVCustSelectList
                            End If
                            If Me.CustSelectColumns.Count < CommonForms.ListViewColumns.LVCustSelectList.Count Then
                                Dim i As Integer
                                For i = Me.CustSelectColumns.Count To CommonForms.ListViewColumns.LVCustSelectList.Count - 1
                                    Me.CustSelectColumns.Add(CommonForms.ListViewColumns.LVCustSelectList.Item(i).Clone)
                                Next
                            End If
                            Me.pShowRSS = .pShowRSS

                        End With
                    Catch ex As Exception
                        blnReturn = False
                    End Try
                Else
                    MsgBox("No user options found")
                    blnReturn = False
                End If
            Catch ex As Exception
            End Try
            Return blnReturn
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' What is the File name for these options
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Filename(Optional ByVal Opt As Integer = 0) As String
            Dim strFilename As String = Application.UserAppDataPath

            'Dim i As Integer = strFilename.Length

            'While i > 0 And strFilename.Chars(i - 1) <> "\"
            '    i -= 1
            'End While

            'strFilename = Mid(strFilename, 1, i) & "UserInfo"


            ''IO.Directory.SetCurrentDirectory(Application.ExecutablePath)
            'If Not IO.Directory.Exists(strFilename) Then

            '    IO.Directory.CreateDirectory(strFilename)
            'End If
            strFilename += "\"
            Dim strName = Security.Principal.WindowsIdentity.GetCurrent.Name
            Dim iPos As Integer = InStr(strName, "\")
            If iPos > 0 Then
                strFilename += Mid(strName, iPos + 1) & ".xml"
            Else
                strFilename += "Default.xml"
            End If

            Return strFilename
        End Function

        Private Function DefaultFilename() As String
            Dim strFilename As String = MedscreenLib.Constants.GCST_X_DRIVE & "\XMLDEFAULTS\"

            Dim i As Integer = strFilename.Length

            While i > 0 And strFilename.Chars(i - 1) <> "\"
                i -= 1
            End While

            'strFilename = Mid(strFilename, 1, i) & "Defaults"
            'IO.Directory.SetCurrentDirectory(Application.ExecutablePath)
            If Not IO.Directory.Exists(strFilename) Then

                IO.Directory.CreateDirectory(strFilename)
            End If
            strFilename += "DefaultUser.XML"
            Return strFilename

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Save these data to a file (serialise to XML)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Save() As Boolean
            Dim myWriter As IO.StreamWriter = Nothing
            Dim blnReturn As Boolean = True
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(UserOptions))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = Filename()
                myWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
            Catch ex As Exception
                blnReturn = False
            Finally
                myWriter.Close()

            End Try
            Return blnReturn
        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.UserColours
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A class that ties together a colour and a User Option
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class UserColours
        Private myBackground As Integer
        Private myForeground As Integer
        Private myUserOption As String
        Private Alpha As Long = 4278190080

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Background Colour as an integer 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Background() As Integer
            Get
                Return myBackground
            End Get
            Set(ByVal Value As Integer)
                myBackground = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground colour as an integer 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Foreground() As Integer
            Get
                Return myForeground
            End Get
            Set(ByVal Value As Integer)
                myForeground = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Background blue channel
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property BackBlue() As Integer
            Get
                Return (myBackground And 255)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Background Green Channel (second byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property BackGreen() As Integer
            Get
                Return (myBackground And 65280) / 256
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Background Red Channel (third byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property BackRed() As Integer
            Get
                Return (myBackground And 16711680) / 65536
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Background colour as a Color 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>Uses FromARGB with an alpha of 255
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property BackColour() As System.Drawing.Color
            Get
                Return Color.FromArgb(255, BackRed, BackGreen, BackBlue)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground colour for the option
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property ForeColour() As Color
            Get
                Return Color.FromArgb(255, foreRed, foreGreen, foreBlue)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground Blue channel (ist byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property foreBlue() As Integer
            Get
                Return (myForeground And 255)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground Green channel (2nd byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property foreGreen() As Integer
            Get
                Return (myForeground And 65280) / 256
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground Red channel (3rd byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property foreRed() As Integer
            Get
                Return (myForeground And 16711680) / 65536
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Foreground Alpha channel (4th byte)
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property foreAlpha() As Integer
            Get
                Return (CLng(myForeground) And CLng(4278190080)) / 16711680
            End Get
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' User option associated with this object
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property UserOption() As String
            Get
                Return myUserOption
            End Get
            Set(ByVal Value As String)
                myUserOption = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new object 
        ''' </summary>
        ''' <param name="User">User Option</param>
        ''' <param name="inBackground">Background Colour</param>
        ''' <param name="inForeground">Foreground Colour</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal User As String, ByVal inBackground As Integer, ByVal inForeground As Integer)
            MyBase.New()
            UserOption = User
            Background = inBackground
            Foreground = inForeground
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new UserColour object 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.UserColourCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of User Colours
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class UserColourCollection

        Inherits CollectionBase

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Collection 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get an Item by position from the collection 
        ''' </summary>
        ''' <param name="index">Position to get</param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As Integer) As UserColours
            Get
                Return CType(MyBase.List.Item(index), UserColours)
            End Get
            Set(ByVal Value As UserColours)
                MyBase.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get Item by User Option 
        ''' </summary>
        ''' <param name="index">User Option</param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As String) As UserColours
            Get
                Dim uc As UserColours
                Dim i As Integer
                For i = 0 To Count - 1
                    uc = Item(i)
                    If uc.UserOption.ToUpper = index.ToUpper Then
                        Exit For
                    End If
                    uc = Nothing
                Next
                Return uc
            End Get
            Set(ByVal Value As UserColours)

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add the various defaults to a new entry 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function AddDefaults() As Boolean
            Dim DC As New UserColours("SELECTED", -2147483635, -2147483634)
            Me.Add(DC)
            DC = New UserColours("C", 12648384, 26127)
            Me.Add(DC)
            DC = New UserColours("P", 16761024, 8388608)
            Me.Add(DC)
            DC = New UserColours("V", 12632319, 128)
            Me.Add(DC)
            DC = New UserColours("W", 12648447, 30604)
            Me.Add(DC)
            DC = New UserColours("X", 7697781, 13224393)
            Me.Add(DC)
            DC = New UserColours("H", 12648384, 13224393)
            Me.Add(DC)
            DC = New UserColours("M", 7697781, 128)
            Me.Add(DC)
            DC = New UserColours("INVOICED", Color.LightBlue.ToArgb, Color.DarkBlue.ToArgb)
            Me.Add(DC)

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Set the default values
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function setDefaults() As Boolean

            Dim PaleLilac As Color = Color.FromArgb(255, 252, 250, 255)
            Dim DC As UserColours
            DC = Me.Item("SELECTED")
            If DC Is Nothing Then
                DC = New UserColours("SELECTED", -2147483635, -2147483634)
                Me.Add(DC)
            Else
                DC.Background = -2147483635
                DC.Foreground = -2147483634
            End If

            DC = Me.Item("C")
            If DC Is Nothing Then
                DC = New UserColours("C", PaleLilac.ToArgb, 26127)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = 26127
            End If

            DC = Me.Item("P")
            If DC Is Nothing Then
                DC = New UserColours("P", PaleLilac.ToArgb, 8388608)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = 8388608
            End If

            DC = Me.Item("V")
            If DC Is Nothing Then
                DC = New UserColours("V", PaleLilac.ToArgb, 128)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = 128
            End If

            DC = Me.Item("W")
            If DC Is Nothing Then
                DC = New UserColours("W", PaleLilac.ToArgb, 30604)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = 30604
            End If

            DC = Me.Item("X")
            If DC Is Nothing Then
                DC = New UserColours("X", PaleLilac.ToArgb, Color.Red.ToArgb)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = Color.Red.ToArgb
            End If

            DC = Me.Item("H")
            If DC Is Nothing Then
                DC = New UserColours("H", PaleLilac.ToArgb, Color.DarkKhaki.ToArgb)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = Color.DarkKhaki.ToArgb
            End If

            DC = Me.Item("M")
            If DC Is Nothing Then
                DC = New UserColours("M", PaleLilac.ToArgb, Color.DarkBlue.ToArgb)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = Color.DarkBlue.ToArgb
            End If
            DC = Me.Item("INVOICED")
            If DC Is Nothing Then
                DC = New UserColours("INVOICED", PaleLilac.ToArgb, Color.DarkBlue.ToArgb)
                Me.Add(DC)
            Else
                DC.Background = PaleLilac.ToArgb
                DC.Foreground = Color.DarkBlue.ToArgb
            End If

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add a  new USerColour Item to the collection 
        ''' </summary>
        ''' <param name="item">UserColour to Add</param>
        ''' <returns>Position of Added Item</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As UserColours) As Integer
            Return MyBase.List.Add(item)
        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.ColumnInfo
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A column in a list view 
    ''' </summary>
    ''' <remarks>
    ''' This object ties together a Column in a list view and field in the database
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ColumnInfo
        Implements IComparable
        Implements ICloneable
        Private intIndex As Integer = -1
        Private strFieldName As String = ""
        Private strHeaderText As String = ""
        Private intColumnWidth As Integer = -1
        Private blnShow As Boolean = True
        Private intImageIndex As Integer = -1
        Private myFixedWidth As Boolean = False

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Position of the column in the unsorted list
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Index() As Integer
            Get
                Return intIndex
            End Get
            Set(ByVal Value As Integer)
                intIndex = Value
            End Set
        End Property

        Public Property FixedWidth() As Boolean
            Get
                Return myFixedWidth
            End Get
            Set(ByVal value As Boolean)
                myFixedWidth = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Index of any bitmap associated with the column 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ImageIndex() As Integer
            Get
                Return intImageIndex
            End Get
            Set(ByVal Value As Integer)
                intImageIndex = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Display with of column 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ColumnWidth() As Integer
            Get
                If intColumnWidth = 0 Then
                    Return -1
                Else
                    Return intColumnWidth
                End If

            End Get
            Set(ByVal Value As Integer)
                intColumnWidth = Value
            End Set
        End Property



        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' NAme of field associated with column 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property FieldName() As String
            Get
                Return strFieldName
            End Get
            Set(ByVal Value As String)
                strFieldName = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Header text for column 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property HeaderText() As String
            Get
                Return strHeaderText
            End Get
            Set(ByVal Value As String)
                strHeaderText = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Make column visible or not 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Show() As Boolean
            Get
                Return blnShow
            End Get
            Set(ByVal Value As Boolean)
                blnShow = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new ColumnInfo Object 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Compare to method
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
            If Me.Index > CType(obj, ColumnInfo).Index Then
                Return 1
            ElseIf Me.Index < CType(obj, ColumnInfo).Index Then
                Return 1
            Else
                Return 0
            End If
        End Function

  
        Public Function Clone() As Object Implements System.ICloneable.Clone
            Dim objCI As New ColumnInfo()
            objCI = Me.MemberwiseClone
            Return objCI
        End Function
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.ColumnInfoCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of Column Info Items
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ColumnInfoCollection

        Inherits CollectionBase

#Region "Declarations"


        Private strListViewName As String


        Private Declare Function SendMessage Lib "user32" _
        Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal msg As Int32, ByVal wParam As Int32, ByRef lParam As LV_COLUMN) As Integer

        Private Declare Function SendMessageI Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal msg As Int32, ByVal wParam As Int32, ByRef lParam As Int32) As IntPtr

        Private Declare Function SendMessageP Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal msg As Int32, ByVal wParam As Int32, ByRef lParam As IntPtr) As IntPtr

        Structure LV_COLUMN
            Public mask As Int32
            Public fmt As Int32
            Public cx As Int32
            Public pszText As [String]
            Public cchTextMax As Int32
            Public iSubItem As Int32
            Public iImage As Int32
            Public iOrder As Int32
        End Structure 'LV_COLUMN

        Const LVM_FIRST As Int32 = &H1000
        Const LVM_GETCOLUMN As Int32 = LVM_FIRST + 95
        Const LVM_SETCOLUMN As Int32 = LVM_FIRST + 96
        Const LVM_GETBKCOLOR As Int32 = (LVM_FIRST + 0)
        Const LVM_SETBKCOLOR As Int32 = (LVM_FIRST + 1)
        Const LVM_GETIMAGELIST As Int32 = (LVM_FIRST + 2)
        Const LVM_SETIMAGELIST As Int32 = (LVM_FIRST + 3)
        Const LVM_GETITEMCOUNT As Int32 = (LVM_FIRST + 4)

        Const LVM_DELETEITEM As Int32 = (LVM_FIRST + 8)
        Const LVM_DELETEALLITEMS As Int32 = (LVM_FIRST + 9)
        Const LVM_GETCALLBACKMASK As Int32 = (LVM_FIRST + 10)
        Const LVM_SETCALLBACKMASK As Int32 = (LVM_FIRST + 11)
        Const LVM_GETNEXTITEM As Int32 = (LVM_FIRST + 12)

        Const LVM_SETITEMPOSITION As Int32 = (LVM_FIRST + 15)
        Const LVM_GETITEMPOSITION As Int32 = (LVM_FIRST + 16)

        Const LVM_HITTEST As Int32 = (LVM_FIRST + 18)
        Const LVM_ENSUREVISIBLE As Int32 = (LVM_FIRST + 19)
        Const LVM_SCROLL As Int32 = (LVM_FIRST + 20)
        Const LVM_REDRAWITEMS As Int32 = (LVM_FIRST + 21)
        Const LVM_ARRANGE As Int32 = (LVM_FIRST + 22)

        Const LVM_GETEDITCONTROL As Int32 = (LVM_FIRST + 24)

        Const LVM_DELETECOLUMN As Int32 = (LVM_FIRST + 28)
        Const LVM_GETCOLUMNWIDTH As Int32 = (LVM_FIRST + 29)
        Const LVM_SETCOLUMNWIDTH As Int32 = (LVM_FIRST + 30)

        Const LVM_GETHEADER As Int32 = (LVM_FIRST + 31)     'IE 3+ only

        Const LVM_CREATEDRAGIMAGE As Int32 = (LVM_FIRST + 33)
        Const LVM_GETVIEWRECT As Int32 = (LVM_FIRST + 34)
        Const LVM_GETTEXTCOLOR As Int32 = (LVM_FIRST + 35)
        Const LVM_SETTEXTCOLOR As Int32 = (LVM_FIRST + 36)
        Const LVM_GETTEXTBKCOLOR As Int32 = (LVM_FIRST + 37)
        Const LVM_SETTEXTBKCOLOR As Int32 = (LVM_FIRST + 38)
        Const LVM_GETTOPINDEX As Int32 = (LVM_FIRST + 39)
        Const LVM_GETCOUNTPERPAGE As Int32 = (LVM_FIRST + 40)
        Const LVM_GETORIGIN As Int32 = (LVM_FIRST + 41)
        Const LVM_UPDATE As Int32 = (LVM_FIRST + 42)
        Const LVM_SETITEMSTATE As Int32 = (LVM_FIRST + 43)
        Const LVM_GETITEMSTATE As Int32 = (LVM_FIRST + 44)
        Const LVM_SETITEMCOUNT As Int32 = (LVM_FIRST + 47)
        Const LVM_SORTITEMS As Int32 = (LVM_FIRST + 48)
        Const LVM_SETITEMPOSITION32 As Int32 = (LVM_FIRST + 49)
        Const LVM_GETSELECTEDCOUNT As Int32 = (LVM_FIRST + 50)
        Const LVM_GETITEMSPACING As Int32 = (LVM_FIRST + 51)

        Const LVM_SETICONSPACING As Int32 = (LVM_FIRST + 53)     'IE 3+ only

        Const LVM_GETSUBITEMRECT As Int32 = (LVM_FIRST + 56)
        Const LVM_SUBITEMHITTEST As Int32 = (LVM_FIRST + 57)
        Const LVM_SETCOLUMNORDERARRAY As Int32 = (LVM_FIRST + 58)
        Const LVM_GETCOLUMNORDERARRAY As Int32 = (LVM_FIRST + 59)
        Const LVM_SETHOTITEM As Int32 = (LVM_FIRST + 60)
        Const LVM_GETHOTITEM As Int32 = (LVM_FIRST + 61)
        Const LVM_SETHOTCURSOR As Int32 = (LVM_FIRST + 62)
        Const LVM_GETHOTCURSOR As Int32 = (LVM_FIRST + 63)
        Const LVM_APPROXIMATEVIEWRECT As Int32 = (LVM_FIRST + 64)
        Const LVM_SETWORKAREA As Int32 = (LVM_FIRST + 65)

        Const LVCF_FMT As Int32 = &H1
        Const LVCF_WIDTH As Int32 = &H2
        Const LVCF_TEXT As Int32 = &H4
        Const LVCF_SUBITEM As Int32 = &H8
        Const LVCF_IMAGE As Int32 = &H10
        Const LVCF_ORDER As Int32 = &H20
        Const LVCFMT_IMAGE As Int32 = 2048
#End Region


#Region "Header Control Messages"
        Const HDM_FIRST = &H1200
        Const HDM_GETITEMRECT = (HDM_FIRST + 7)
        Const HDM_HITTEST = (HDM_FIRST + 6)
        Const HDM_SETIMAGELIST = (HDM_FIRST + 8)
        Const HDM_GETITEMW = (HDM_FIRST + 11)
        Const HDM_ORDERTOINDEX = (HDM_FIRST + 15)
#End Region

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' What field to index the item property on 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Enum CICIndexOn
            ''' <summary>Field Name</summary>
            FieldName
            ''' <summary>Column Order</summary>
            Order
        End Enum

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection for a particular List View
        ''' </summary>
        ''' <param name="ListViewName"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal ListViewName As String)
            MyBase.New()
            strListViewName = ListViewName
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Expose the name of the list view 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property ListViewName() As String
            Get
                Return strListViewName
            End Get
            Set(ByVal Value As String)
                strListViewName = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a ColumnInfo Item from the collection by position 
        ''' </summary>
        ''' <param name="index">Index of the item to get </param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As Integer) As ColumnInfo
            Get
                Return CType(MyBase.List.Item(index), ColumnInfo)
            End Get
            Set(ByVal Value As ColumnInfo)
                MyBase.List.Item(index) = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get item by value the type of which is defined in the second parameter from an enumeration
        ''' </summary>
        ''' <param name="index">Item to get</param>
        ''' <param name="Indexon">What type is the the index</param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Default Public Property Item(ByVal index As String, Optional ByVal Indexon As CICIndexOn = CICIndexOn.FieldName) As ColumnInfo
            Get
                Dim uc As ColumnInfo
                Dim i As Integer
                For i = 0 To Count - 1
                    uc = Item(i)
                    If Indexon = CICIndexOn.FieldName Then
                        If uc.FieldName.ToUpper = index.ToUpper Then
                            Exit For
                        End If
                    ElseIf Indexon = CICIndexOn.Order Then
                        If uc.Index = CInt(index) Then
                            Exit For
                        End If
                    End If
                    uc = Nothing
                Next
                Return uc
            End Get
            Set(ByVal Value As ColumnInfo)
                Dim uc As ColumnInfo
                Dim i As Integer
                For i = 0 To Count - 1
                    uc = Item(i)
                    If Indexon = CICIndexOn.FieldName Then
                        If uc.FieldName.ToUpper = index.ToUpper Then
                            Item(i) = Value
                            Exit For
                        End If
                    ElseIf Indexon = CICIndexOn.Order Then
                        If uc.Index = CInt(index) Then
                            Item(i) = Value
                            Exit For
                        End If
                    End If

                Next

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Sort the collection, using internal comparer
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub Sort()
            Dim oCp As New ColInfoComparer(0, 1)
            MyBase.InnerList.Sort(oCp)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create the columns for the list view adding them into the list view
        ''' </summary>
        ''' <param name="MyListView">List view object</param>
        ''' <param name="myColImgList">List of images</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function CreateColumns(ByRef MyListView As ListView, ByVal myColImgList As ImageList, Optional ByVal RemoveFirst As Boolean = False)
            MyListView.Columns.Clear()                  'Get rid of any columns
            Dim oColInfo As ColumnInfo
            Me.Sort()                                   'Sort myself



            For Each oColInfo In Me.List                'For each column to add

                If oColInfo.Show Then                   'If visible
                    Dim ColHead As ColumnHeader = New ColumnHeader() 'create a new column header
                    ColHead.Text = oColInfo.HeaderText               'Add header text
                    If Not oColInfo.FixedWidth Then
                        ColHead.Width = oColInfo.ColumnWidth             'Set column width
                        If ColHead.Width < 50 Then ColHead.Width = 50
                    End If
                    MyListView.Columns.Add(ColHead)                  'Add Column to collection
                End If
                'Check to see if we need to almost hide the first column

            Next


        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get the columns out of a list view
        ''' </summary>
        ''' <param name="ListView">Listview to use </param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub GetColumns(ByVal ListView As ListView)
            Try
                Dim column As ColumnHeader                      'Create a column header object
                For Each column In ListView.Columns             'Go through each object in list
                    Dim pcol As New LV_COLUMN()                 'Create a new structure 

                    pcol.mask = LVCF_ORDER Or LVCF_WIDTH        'Set the mask to width and position 

                    'Ask for info on that column using a message to control 
                    Dim ret As Boolean = SendMessage(ListView.Handle, LVM_GETCOLUMN, column.Index, pcol)
                    If ret Then                                     'If succesful
                        Me.Item(column.Index).Index = pcol.iOrder   'Set up Item 
                        Me.Item(column.Index).ColumnWidth = pcol.cx
                    End If
                    'listViewCols.Add(New ListViewColumn(column.Text, column.Width, pcol.iOrder))
                Next column
            Catch ex As Exception

            End Try

        End Sub

        Public Function Save(ByVal Filename As String)
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(ColumnInfoCollection))
            Dim writer As New IO.FileStream(Filename, IO.FileMode.Create)
            ' Serialize the object, and close the TextWriter
            mySerializer.Serialize(writer, Me)
            writer.Close()

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add the item into collection
        ''' </summary>
        ''' <param name="item">ColumnInfo item to add</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As ColumnInfo) As Integer
            Return MyBase.List.Add(item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' Project	 : MedscreenCommonGui
        ''' Class	 : UserDefaults.ColumnInfoCollection.ColInfoComparer
        ''' 
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Comparer for this collection
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Class ColInfoComparer
            Implements IComparer

            Private myDirection As Integer

            Private col As Integer

            Public Sub New()
                col = 0
            End Sub

            Public Sub New(ByVal iOption As Integer, ByVal idIrection As Integer)
                myDirection = idIrection
                col = iOption
            End Sub

            Public Overloads Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements IComparer.Compare


                Try
                    Dim xObj As ColumnInfo = CType(x, ColumnInfo)
                    Dim yObj As ColumnInfo = CType(y, ColumnInfo)

                    Return [String].Compare(xObj.Index.ToString("000"), yObj.Index.ToString("000")) * myDirection

                Catch ex As Exception
                    Return -1
                End Try
            End Function
        End Class
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenCommonGui
    ''' Class	 : UserDefaults.ColumnDefaults
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Defaults for columns of a particular ListView
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class ColumnDefaults
        Private myLVDiaryList As ColumnInfoCollection
        Private myCustSelectList As ColumnInfoCollection

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new Column Default Object
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            Load("LvDiaryColumns", myLVDiaryList)
            Load("LvCustSelectColumns", Me.myCustSelectList)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' List of column defaults 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property LVDiaryList() As ColumnInfoCollection
            Get
                Return myLVDiaryList
            End Get
            Set(ByVal Value As ColumnInfoCollection)
                myLVDiaryList = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Customer Selection List defaults 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [09/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property LVCustSelectList() As ColumnInfoCollection
            Get
                Return Me.myCustSelectList
            End Get
            Set(ByVal Value As ColumnInfoCollection)
                myCustSelectList = Value
            End Set
        End Property
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load these defaults from the file 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Private Function Load(ByVal Path As String, ByRef InfoColl As ColumnInfoCollection)
            ' To read the file, create a FileStream object.
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(ColumnInfoCollection))

            Dim strFileName As String = Filename(Path)
            If IO.File.Exists(strFileName) Then
                Dim myFileStream As IO.FileStream
                Try
                    myFileStream = New IO.FileStream(strFileName, IO.FileMode.Open)
                    InfoColl = CType( _
                    mySerializer.Deserialize(myFileStream), ColumnInfoCollection)
                Catch ex As Exception
                Finally
                    myFileStream.Close()
                End Try
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create Filename for object 
        ''' </summary>
        ''' <param name="ListViewID">Id of List view</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Filename(ByVal ListViewID As String) As String
            Dim strFilename As String = "\\corp.concateno.com\medscreen\common\xmldefaults"

            'Dim i As Integer = strFilename.Length

            'While i > 0 And strFilename.Chars(i - 1) <> "\"
            '    i -= 1
            'End While
            'strFilename = Mid(strFilename, 1, i) & "XML"


            ''IO.Directory.SetCurrentDirectory(Application.ExecutablePath)
            'If Not IO.Directory.Exists(strFilename) Then

            '    IO.Directory.CreateDirectory(strFilename)
            'End If
            strFilename += "\" & ListViewID & ".xml"
            Return strFilename
        End Function


    End Class

End Namespace
