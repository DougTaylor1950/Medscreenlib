Imports System.Xml
Imports System.Xml.Serialization
Imports System.Windows.Forms
Imports MedscreenLib.Medscreen

Namespace Glossary


#Region "Options"

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.Options
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Individual option definitions that are available to customer options
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/02/2005]</date><Action>New fields added (reference and functional group)
    ''' </Action></revision>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class Options

#Region "Shared"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Converts the DUE_OFF option value in to human readable form
        ''' </summary>
        ''' <param name="optValue"></param>
        ''' <returns></returns>
        ''' <remarks>
        '''   If the customer does not have the DUE_OFF options set, the default value
        '''   from the options table should be passed in.
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[Boughton]</Author><date> [25/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Shared Function Option_Translate_DueOff(ByVal optValue As String) As String
            Dim sValue As String, sOp As String, iHour As Single, sField As String = "", sSetTime As String
            Dim bContinue As Boolean, iPos As Integer, sHour As String = "", sLastOp As String = ""
            Dim description As New System.Text.StringBuilder(optValue.Trim & Environment.NewLine)

            If optValue Is Nothing OrElse optValue.Trim = "" Then
                Return ("Option value is blank")
            End If

            '{ Process each section of option (comma delimited) }
            For Each optValue In optValue.Split(","c)
                optValue = optValue.Trim
                '{ Find the operator for the current segment }
                sOp = optValue.Substring(0, 1)

                If sOp = "#" Then
                    ' Line format: #Library.Routine# [&]
                    '	See comments below for additional processing.
                    ' Complex calculation, call VGL routine to calculate due off date 
                    ' Library is between first # and . routine is between . and final # 
                    sField = ExtractStringSection(optValue, 1, "#", ".")
                    sValue = ExtractStringSection(optValue, 1, ".", "#")
                    description.Append("Custom calculation using routine " & sValue & " in " & sField)
                    ' Reset field in case there are further instructions
                    sField = ""
                Else
                    ' Line format (items in [] optional): op H = [F] +"D" [@"T"] [!][&] [, ...]
                    '  	Where op = operator, H = hour received, F = field, D = days to add, T = time
                    '	  Spaces are optional, H is in numeric format, D and T interval format
                    '	  op: ! = Non workday, < = before H, > = after H, = = unconditional
                    '	  F should be "REC" (default) or "COL" for received or collected and may only be specified once
                    '	  The & operator indicates processing continues even after a match
                    '	  The end ! operator indicates result does not have to be working day 

                    ' == H section ==
                    '  Conditional received time, appears before the = sign 
                    If optValue.IndexOf("=") > 0 Then
                        sHour = optValue.Substring(1, optValue.IndexOf("=") - 1)
                        If IsNumeric(sHour) Then
                            '{ iHour is in real format, convert to hours and minutes }
                            iHour = CType(sHour, Single)
                            sHour = iHour.ToString("00")
                            iHour = (iHour - CSng(sHour)) * 60
                            sHour = sHour & ":" & iHour.ToString("00")
                        Else
                            sHour = ""
                        End If
                    End If

                    ' == F section ==
                    '  Only pick up date field if we haven't already 
                    If sField = "" Then
                        '{ The field to use for calculation appears between the = a + signs }
                        sField = ExtractStringSection(optValue, 1, "=", "+").Trim.ToUpper
                        If Left(sField, 3) = "COL" Then
                            sField = " collected "
                        Else
                            sField = " received "
                        End If
                    End If

                    ' Create a description based on op F and H 
                    If (sOp = "=") Then
                        ' An "=" instruction on a new line, must mean "otherwise".
                        ' (On the first line, it is the only instruction so "otherwise" doesn't make sense) 
                        If (Not sLastOp Is Nothing) And sLastOp <> "=" And Not bContinue Then
                            description.Append("Otherwise ")
                        End If
                    ElseIf sOp = "<" Then
                        description.Append("If" & sField & "before " & sHour & ", ")
                    ElseIf sOp = ">" Then
                        description.Append("If" & sField & "after " & sHour & ", ")
                    ElseIf sOp = "!" Then
                        description.Append("If" & sField & "on a non-working day, ")
                    End If

                    ' == D Section ==
                    '  Find the amount of time to add to the date for due off.
                    '  Working days to add appears in quotes after = sign.
                    '  The value is in interval format, so may contain days, hours, minutes etc 
                    sValue = ExtractStringSection(optValue, InStr(optValue, "="), "", "").Trim

                    ' Ensure sValue will convert to an interval 
                    Dim addTime As TimeSpan
                    Try
                        addTime = Interval.Parse(sValue).ToTimeSpan
                    Catch
                        addTime = TimeSpan.Zero
                    End Try

                    ' Append due out details to description 
                    If addTime.Days > 0 Then
                        description.Append("Due out in " & addTime.Days & " day(s)")
                        If addTime.Hours > 0 Then
                            description.Append(" " & addTime.Hours & " hour(s)")
                        End If
                    Else
                        description.Append("Due out the same day")
                        If addTime.Hours > 0 Then
                            description.Append(" in " & addTime.Hours & " hour(s)")
                        End If
                    End If

                    ' == @T Section ==
                    '  Check to see whether due off at a specific time 
                    iPos = optValue.IndexOf("@")
                    If iPos >= 0 Then
                        sSetTime = ExtractStringSection(optValue, iPos, "", "").Trim
                        description.Append(" by " & sSetTime.Substring(0, 5))
                    ElseIf Not bContinue Then
                        description.Append(" from" & sField & "time")
                    End If

                End If

                ' == & and ! end sections == 
                bContinue = optValue.EndsWith("&")
                If Not bContinue Then
                    ' Check whether due off date must be a working day 
                    If optValue.LastIndexOf("!") >= optValue.Length - 2 Then
                        description.Append(" even if it is a non-working day")
                    Else
                        description.Append(" or the next working day after that date")
                    End If
                Else
                    ' Next line is included in calculation, even if this line matches
                    description.Append(" and then")
                End If
                description.Append(Environment.NewLine)
                sLastOp = sOp
            Next

            Return (description.ToString)

        End Function
#End Region

#Region "Declarations"


        Private PhraseList As MedscreenLib.Glossary.PhraseCollection

        Private myFields As New TableFields("OPTIONS")

        Private objOptionId As New StringField("IDENTITY", "", 10, True)
        Private objOptionType As New StringField("OPTION_TYPE", "", 5)
        Private objOptionDesc As New StringField("OPTION_DESCRIPTION", "", 30)
        Private objOptionDefault As New StringField("OPTION_DEFAULT", "", 100)
        Private objOptionAuthority As New StringField("AUTHORITY", "", 10)
        Private objFuncGroup As New StringField("FUNC_GROUP", "", 20)
        Private objReference As New StringField("REFERENCE", "", 30)
        Private objLibrary As New StringField("LIBRARY", "", 20)
        Private objROUTINE As New StringField("ROUTINE", "", 40)
        Private objHelpComment As New StringField("HELP_COMMENT", "", 255)
        Private objRemoveFlag As New BooleanField("REMOVEFLAG", "F")
        Private myOptionPlus As OptionPlus = Nothing
        Private myOptionSettings As OptionSettings = Nothing
        Private myColourList As ColourList = Nothing
#End Region


#Region "Public Functions"
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert option to XML
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToXML() As String
            Dim strRet As New Text.StringBuilder("")
            strRet.Append(Me.Fields.ToXMLSchema(True, , "OPTIONS", True))
            strRet.Append("<Default>")
            If Me.OptionType = "BOOL" Then      'Deal with Booleans
                If Me.DefaultValue = "T" Then
                    strRet.Append("True")
                Else
                    strRet.Append("False")
                End If
            ElseIf Me.OptionType = "RANGE" Then
                Dim objPhrase As MedscreenLib.Glossary.Phrase = Me.Phrases.Item(Me.DefaultValue)
                If Not objPhrase Is Nothing Then
                    strRet.Append(objPhrase.PhraseText)
                End If
            Else
                strRet.Append(Me.DefaultValue)
            End If
            strRet.Append("</Default>")
            strRet.Append("</OPTIONS>")
            Return strRet.ToString
        End Function
#End Region

        Friend Property Fields() As TableFields
            Get
                Return myFields
            End Get
            Set(ByVal Value As TableFields)
                myFields = Value
            End Set
        End Property

        Public Property RemoveFlag() As Boolean
            Get
                Return Me.objRemoveFlag.Value
            End Get
            Set(ByVal value As Boolean)
                Me.objRemoveFlag.Value = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new option entity
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()

            myFields.Add(Me.objOptionId)
            myFields.Add(Me.objOptionType)
            myFields.Add(Me.objOptionDesc)
            myFields.Add(Me.objOptionDefault)
            myFields.Add(Me.objOptionAuthority)
            myFields.Add(Me.objFuncGroup)
            myFields.Add(Me.objReference)
            myFields.Add(Me.objLibrary)
            myFields.Add(Me.objROUTINE)
            myFields.Add(Me.objHelpComment)
            myFields.Add(objRemoveFlag)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get an option by id
        ''' </summary>
        ''' <param name="id"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	11/01/2008	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal id As String)
            MyClass.New()
            Me.OptionID = id
            Dim ocoll As New Collection()
            ocoll.Add(CConnection.StringParameter("ID", id, 10))
            Me.myFields.Load(CConnection.DbConnection, "identity = ?", ocoll, True)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Provide editing capability for form 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Edit() As Boolean
            Dim myReturn As Boolean = False
            Dim myForm As New frmOptions()
            Try
                myForm._Option = Me
                If myForm.ShowDialog = DialogResult.OK Then
                    With myForm._Option
                        Me.DefaultValue = .DefaultValue
                        Me.FunctionalGroup = .FunctionalGroup
                        Me.HelpComment = .HelpComment
                        Me.Library = .Library
                        Me.Reference = .Reference
                        Me.Routine = .Routine
                        Me.OptionAuthority = .OptionAuthority
                        Me.OptionDescription = .OptionDescription
                        Me.Fields.Update(MedConnection.Connection)
                    End With
                    myReturn = True
                End If
            Catch ex As Exception
            End Try
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Default value for option
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property DefaultValue() As String
            Get
                Return Me.objOptionDefault.Value
            End Get
            Set(ByVal Value As String)
                objOptionDefault.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Check to see if phrase type if so load phrase list
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action>
        ''' Phrases are now stored in the reference field</Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function PhraseCheck() As Boolean
            If Me.OptionType = OptionCollection.GCST_OPTION_RANGE Or _
            Me.OptionType = OptionCollection.GCST_OPTION_BOOLRANGE Then
                PhraseList = New MedscreenLib.Glossary.PhraseCollection(Me.Reference)
                'PhraseList.Load()
            ElseIf Me.OptionType = OptionCollection.GCST_OPTION_TABLE Then
                PhraseList = New MedscreenLib.Glossary.PhraseCollection(Me.Reference, PhraseCollection.BuildBy.Table, "Removeflag = 'F'")
            End If
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Update Option 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/06/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Update() As Boolean
            Return Me.Fields.Update(MedConnection.Connection)
        End Function

        Public Function Insert() As Boolean
            Return Me.Fields.Insert(MedConnection.Connection)
        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Phrases associated with phrase type options
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Phrases() As MedscreenLib.Glossary.PhraseCollection
            Get
                If Not PhraseList Is Nothing Then
                    If PhraseList.Count = 0 Then
                        PhraseList.Load()
                    End If
                End If
                Return PhraseList
            End Get
            Set(ByVal Value As MedscreenLib.Glossary.PhraseCollection)
                PhraseList = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Id of option
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionID() As String
            Get
                Return Me.objOptionId.Value
            End Get
            Set(ByVal Value As String)
                Me.objOptionId.Value = Value
                If Value.Trim.Length > 0 Then           'Load Option Plus
                    OptionPlus.OptionID = Value
                Else
                    OptionPlus = Nothing
                End If
            End Set
        End Property

        Public ReadOnly Property Editable() As Boolean
            Get
                Return OptionPlus.Editable
            End Get
        End Property

        

        Public ReadOnly Property Expireable() As Boolean
            Get
                Return Not OptionPlus.NoExpire
            End Get
        End Property

        Public Property ColorList() As ColourList
            Get
                If myColourList Is Nothing Then
                    myColourList = OptionSettings.ColourList
                End If
                Return myColourList
            End Get
            Set(ByVal value As ColourList)
                myColourList = value
            End Set
        End Property

        'Private Property OptSettings() As OptionSettings
        '    Get
        '        If myOptionSettings Is Nothing Then
        '            myOptionSettings = OptionSettings.GetSettings
        '        End If
        '        Return myOptionSettings
        '    End Get
        '    Set(ByVal value As OptionSettings)
        '        myOptionSettings = value
        '    End Set
        'End Property

        Private Property OptionPlus() As OptionPlus
            Get
                If myOptionPlus Is Nothing Then
                    'OptSettings.ToString()
                    Try
                        myOptionPlus = OptionSettings.OptionPlusList.FindOption(OptionID)
                        If myOptionPlus Is Nothing Then
                            myOptionPlus = New OptionPlus With {.OptionID = OptionID}
                            OptionSettings.OptionPlusList.Add(myOptionPlus)
                        End If
                    Catch
                    End Try
                End If
                Return myOptionPlus
            End Get
            Set(ByVal value As OptionPlus)
                myOptionPlus = value
            End Set
        End Property


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Routine associated with option 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Routine() As String
            Get
                Return Me.objROUTINE.Value
            End Get
            Set(ByVal Value As String)
                Me.objROUTINE.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Library associated with Option 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Library() As String
            Get
                Return Me.objLibrary.Value
            End Get
            Set(ByVal Value As String)
                Me.objLibrary.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Long and boring explanation for option 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [24/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property HelpComment() As String
            Get
                Return Me.objHelpComment.Value
            End Get
            Set(ByVal Value As String)
                Me.objHelpComment.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Reference to other item 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Reference() As String
            Get
                Return Me.objReference.Value
            End Get
            Set(ByVal Value As String)

            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Functional Group that option belongs in 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [03/02/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property FunctionalGroup() As String
            Get
                Return Me.objFuncGroup.Value
            End Get
            Set(ByVal Value As String)
                Me.objFuncGroup.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Type of option
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionType() As String
            Get
                Return Me.objOptionType.Value
            End Get
            Set(ByVal Value As String)
                Me.objOptionType.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Authority level associated  with option 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionAuthority() As String
            Get
                Return Me.objOptionAuthority.Value
            End Get
            Set(ByVal Value As String)
                Me.objOptionAuthority.Value = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The user will be able to edit this option if they are in the view USER_CUST_OPTIONS with this option
        ''' </summary>
        ''' <param name="UserID"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	02/10/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function UserCanEdit(ByVal UserID As String) As Boolean
            Dim oColl As New Collection()
            oColl.Add(UserID)
            oColl.Add(Me.OptionID)
            Dim strRet As String = CConnection.PackageStringList("LIB_UTILS.UserCanEditOption", oColl)
            Return (strRet.Trim = "T")
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Role required to edit the option
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' This is used for the List views so that the users can see which role is required if the user feels they should be 
        ''' able to edit the option but can't.  This is to enable user management of roles
        ''' </remarks>
        ''' <history>
        ''' 	[taylor]	03/10/2007	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Function RoleToEdit() As String
            Dim strRet As String = CConnection.PackageStringList("LIB_UTILS.RoleToEditOption", Me.OptionID)
            Return strRet

        End Function
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Description associated with option 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/05/2006]</date><Action>Made Read/Write</Action></revision>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property OptionDescription() As String
            Get
                Return Me.objOptionDesc.Value
            End Get
            Set(ByVal Value As String)
                Me.objOptionDesc.Value = Value
            End Set
        End Property
    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : Glossary.OptionCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of option entities 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class OptionCollection
        Inherits CollectionBase
        Private myFields As New TableFields("OPTIONS")

        Private objOptionId As New StringField("IDENTITY", "", 10, True)
        Private objOptionType As New StringField("OPTION_TYPE", "", 5)
        Private objOptionDesc As New StringField("OPTION_DESCRIPTION", "", 30)
        Private objOptionDefault As New StringField("OPTION_DEFAULT", "", 100)
        Private objOptionAuthority As New StringField("AUTHORITY", "", 10)
        Private objFuncGroup As New StringField("FUNC_GROUP", "", 20)
        Private objReference As New StringField("REFERENCE", "", 30)
        Private objLibrary As New StringField("LIBRARY", "", 20)
        Private objROUTINE As New StringField("ROUTINE", "", 40)
        Private objHelpComment As New StringField("HELP_COMMENT", "", 255)
        Private objRemoveFlag As New BooleanField("REMOVEFLAG", "F")


        ''' <summary>Option of type range tied to a phrase of the same name</summary>
        Public Const GCST_OPTION_RANGE As String = "RANGE"
        ''' <summary>Option of type Boolean</summary>
        Public Const GCST_OPTION_BOOL As String = "BOOL"
        ''' <summary>Option of type value</summary>
        Public Const GCST_OPTION_VALUE As String = "VALUE"
        ''' <summary>Option that is both boolean and value</summary>
        Public Const GCST_OPTION_BOOLVALUE As String = "BLNVL"
        ''' <summary>Option that is boolean and range, tied to a phrase of the same name</summary>
        Public Const GCST_OPTION_BOOLRANGE As String = "BLNRG"
        ''' <summary>Option that relates to an item found in a particular table, 
        ''' Reference field stores the table name, OPTION_VALUE_RANGE in the customer option table 
        ''' stores the field value to find (will always be the primary key field </summary>
        Public Const GCST_OPTION_MULTI As String = "MULTI"
        Public Const GCST_OPTION_TABLE As String = "TABLE"



        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a new collection 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
            myFields.Add(Me.objOptionId)
            myFields.Add(Me.objOptionType)
            myFields.Add(Me.objOptionDesc)
            myFields.Add(Me.objOptionDefault)
            myFields.Add(Me.objOptionAuthority)
            myFields.Add(Me.objFuncGroup)
            myFields.Add(Me.objReference)
            myFields.Add(Me.objLibrary)
            myFields.Add(Me.objROUTINE)
            myFields.Add(Me.objHelpComment)
            myFields.Add(objRemoveFlag)

        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Load a new collection 
        ''' </summary>
        ''' <returns>True if succesful</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Load() As Boolean
            Dim oCmd As New OleDb.OleDbCommand()
            Dim oRead As OleDb.OleDbDataReader = Nothing
            'Dim strID As String
            'Dim strType As String
            'Dim strDesc As String
            Dim objOpt As Options

            Try

                oCmd.Connection = MedConnection.Connection
                oCmd.CommandText = Me.myFields.FullRowSelect & " where removeflag = 'F'"
                CConnection.SetConnOpen()

                oRead = oCmd.ExecuteReader
                While oRead.Read
                    Try

                        objOpt = New Options()
                        objOpt.Fields.readfields(oRead)
                        ''<Removed on 16-Apr-2007 06:35 by taylor> 
                        'objOpt.PhraseCheck()
                        '
                        '</Removed on 16-Apr-2007 06:35 by taylor> 
                        Add(objOpt)
                    Catch ex As Exception
                        Medscreen.LogError(ex, , "Loading option collection")
                    End Try
                End While

            Catch ex As Exception
                Return False
            Finally
                If Not oRead Is Nothing Then
                    If Not oRead.IsClosed Then oRead.Close()
                End If
                CConnection.SetConnClosed()
            End Try
            'Having read them in we can do an option check 
            Try  ' Protecting
                For Each objOpt In Me.List
                    objOpt.PhraseCheck()
                Next
            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "OptionCollection-Load-10008")
            Finally
            End Try

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an option to the collection 
        ''' </summary>
        ''' <param name="objTest">Option to add </param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub Add(ByVal objTest As Options)
            Me.List.Add(objTest)
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Remove an option from the collection 
        ''' </summary>
        ''' <param name="index"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub Remove(ByVal index As Integer)
            ' Check to see if there is a widget at the supplied index.
            If index > Count - 1 Or index < 0 Then
                ' If no widget exists, a messagebox is shown and the operation is 
                ' cancelled.
                'System.Windows.Forms.MessageBox.Show("Index not valid!")
            Else
                ' Invokes the RemoveAt method of the List object.
                List.RemoveAt(index)
            End If
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a option by position 
        ''' </summary>
        ''' <param name="index">Position to get </param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads ReadOnly Property Item(ByVal index As Integer) As Options
            Get
                ' The appropriate item is retrieved from the List object and 
                ' explicitly cast to the Widget type, then returned to the 
                ' caller.
                Return CType(List.Item(index), Options)
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get an option by type 
        ''' </summary>
        ''' <param name="index"></param>
        ''' <param name="intOpt"></param>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [18/10/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads ReadOnly Property Item(ByVal index As String, Optional ByVal intOpt As Integer = 0) As Options
            Get
                ' The appropriate item is retrieved from the List object and 
                ' explicitly cast to the Widget type, then returned to the 
                ' caller.
                Dim objOpt As Options

                For Each objOpt In List
                    If intOpt = 0 Then
                        If objOpt.OptionID = index Then
                            Return objOpt
                            Exit For
                        End If
                    ElseIf intOpt = 1 Then
                        If objOpt.OptionDescription = index Then
                            Return objOpt
                            Exit For
                        End If

                    End If

                Next
                Return Nothing
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Convert to XML 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [30/05/2006]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ToXML() As String
            Dim strRet As New System.Text.StringBuilder("<Options>")
            Dim objOption As Options
            For Each objOption In Me.List
                strRet.Append(objOption.ToXML)
                strRet.Append(Environment.NewLine)
            Next

            strRet.Append("</Options>")
            Return strRet.ToString
        End Function
    End Class

    ''' <summary>
    ''' Add on class to support Options
    ''' </summary>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    ''' <author></author>
    Public Class OptionPlus



        Private myEditable As Boolean = True
        Public Property Editable() As Boolean
            Get
                Return myEditable
            End Get
            Set(ByVal value As Boolean)
                myEditable = value
            End Set
        End Property


        Private myOptionID As String
        Public Property OptionID() As String
            Get
                Return myOptionID
            End Get
            Set(ByVal value As String)
                myOptionID = value
            End Set
        End Property


        Private myNoExpire As Boolean = False
        Public Property NoExpire() As Boolean
            Get
                Return myNoExpire
            End Get
            Set(ByVal value As Boolean)
                myNoExpire = value
            End Set
        End Property


    End Class

    Public Class OptionColours


        Private myOptionProperty As String = ""
        Public Property OptionProperty() As String
            Get
                Return myOptionProperty
            End Get
            Set(ByVal value As String)
                myOptionProperty = value
            End Set
        End Property



        Private myForeColour As String = ""
        Public Property ForeColour() As String
            Get
                Return myForeColour
            End Get
            Set(ByVal value As String)
                myForeColour = value
            End Set
        End Property


        Private myBackColour As String = ""
        Public Property BackColour() As String
            Get
                Return myBackColour
            End Get
            Set(ByVal value As String)
                myBackColour = value
            End Set
        End Property


    End Class

    Public Class ColourList
        Inherits System.Collections.Generic.List(Of OptionColours)


        Public Overloads Sub Add(ByVal Item As OptionColours)
            Dim c As OptionColours = FindColour(Item.OptionProperty)
            If c Is Nothing Then

                MyBase.Add(Item)
            Else
                c = Item

            End If


        End Sub

        Public Function FindColour(ByVal optionId As String) As OptionColours
            Dim pred As New PredicateClass(optionId)
            Dim optColor As OptionColours = Me.Find(AddressOf pred.PredicateFunction)
            
            Return optColor
        End Function

        Private Class PredicateClass
            Private Id As String
            Public Sub New(ByVal code As String)
                Id = code
            End Sub

            Public Function PredicateFunction(ByVal Item As OptionColours) As Boolean
                Return Item.OptionProperty.ToUpper = Id.ToUpper
            End Function
        End Class

    End Class


    Public Class OptionPlusList
        Inherits System.Collections.Generic.List(Of optionplus)
        Public STR_OptionPath As String = MedConnection.Instance.ServerPath & "\configuration\options.xml"
        Public Shared STR_OptionSettingsPath As String = MedConnection.Instance.ServerPath & "\configuration\optionSettings.xml"
        Private Shared myList As OptionPlusList

        Public Shared Function GetList(Optional ByVal Read As Boolean = True) As OptionPlusList
            If myList Is Nothing Then
                myList = New OptionPlusList
                If Read Then myList.Read()
            End If
            Return myList
        End Function
        Private FindId As String = ""

        Public Function FindOption(ByVal optionId As String) As OptionPlus
            Dim pred As New PredicateClass(optionId)
            Return Me.Find(AddressOf pred.PredicateFunction)
        End Function

        Public Overloads Sub Add(ByVal Item As OptionPlus)
            Dim c As OptionPlus = FindOption(Item.OptionID)
            If c Is Nothing Then

                MyBase.Add(Item)
            Else
                c = Item

            End If
        End Sub

        Private Sub New()
            MyBase.New()
        End Sub

        Public Sub Read()
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(OptionPlusList))
            ' To read the file, create a FileStream object.
            Dim strFileName As String = STR_OptionPath
            If IO.File.Exists(strFileName) Then
                Try
                    Dim myFileStream As IO.FileStream = _
                    New IO.FileStream(strFileName, IO.FileMode.Open)
                    Dim Myobject As OptionPlusList = CType( _
                    mySerializer.Deserialize(myFileStream), OptionPlusList)

                    With Myobject
                        For Each op As OptionPlus In Myobject
                            Me.Add(op)
                        Next

                    End With
                    myFileStream.Close()
                    myFileStream = Nothing
                Catch ex As Exception
                End Try
            End If
        End Sub




        Public Sub Write()
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(OptionPlusList))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = STR_OptionPath
                Dim myWriter As IO.StreamWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
                myWriter.Flush()
                myWriter.Close()

            Catch ex As Exception
            End Try
        End Sub



        Private Class PredicateClass
            Private Id As String
            Public Sub New(ByVal code As String)
                Id = code
            End Sub

            Public Function PredicateFunction(ByVal Item As OptionPlus) As Boolean
                Return Item.OptionID.ToUpper = Id.ToUpper
            End Function
        End Class

    End Class

    Public Class OptionSettings
        Private Shared mySettings As OptionSettings = Nothing
        Private Shared myColourList As ColourList
        Public Shared Property ColourList() As ColourList
            Get
                Return myColourList
            End Get
            Set(ByVal value As ColourList)
                myColourList = value
            End Set
        End Property

        Public Property NSOptionPlusList() As OptionPlusList
            Get
                Return myOptionPlusList
            End Get
            Set(ByVal value As OptionPlusList)
                myOptionPlusList = value
            End Set
        End Property

        Public Property NSColourList() As ColourList
            Get
                Return myColourList
            End Get
            Set(ByVal value As ColourList)
                myColourList = value
            End Set
        End Property

        Private Shared myOptionPlusList As OptionPlusList
        Public Shared Property OptionPlusList() As OptionPlusList
            Get
                If myOptionPlusList Is Nothing Then
                    myOptionPlusList = OptionPlusList.GetList(True)
                End If
                Return myOptionPlusList
            End Get
            Set(ByVal value As OptionPlusList)
                myOptionPlusList = value
            End Set
        End Property

        Private Sub New()
            MyBase.New()
            Read()

        End Sub

        Public Shared Function GetSettings() As OptionSettings
            If mySettings Is Nothing Then
                mySettings = New OptionSettings
                'myOptionPlusList = OptionPlusList.GetList(False)
                'myColourList = New ColourList
                mySettings.Read()
            End If
            Return mySettings
        End Function

        Public Sub Read()
            Dim mySerializer As XmlSerializer = New XmlSerializer(GetType(OptionSettings))
            ' To read the file, create a FileStream object.
            Dim strFileName As String = OptionPlusList.STR_OptionSettingsPath
            If IO.File.Exists(strFileName) Then
                Try
                    Dim myFileStream As IO.FileStream = _
                    New IO.FileStream(strFileName, IO.FileMode.Open)
                    Dim Myobject As OptionSettings = CType(mySerializer.Deserialize(myFileStream), OptionSettings)

                    With Myobject
                       

                    End With
                    myFileStream.Close()
                    myFileStream = Nothing
                Catch ex As Exception
                End Try
            End If
        End Sub




        Public Sub Write()
            Try
                Dim mySerializer As System.Xml.Serialization.XmlSerializer = _
                New System.Xml.Serialization.XmlSerializer(GetType(OptionSettings))
                ' To write to a file, create a StreamWriter object.
                Dim strFileName As String = OptionPlusList.STR_OptionSettingsPath
                Dim myWriter As IO.StreamWriter = New IO.StreamWriter(strFileName)
                mySerializer.Serialize(myWriter, Me)
                myWriter.Flush()
                myWriter.Close()

            Catch ex As Exception
            End Try
        End Sub

    End Class



#End Region
End Namespace