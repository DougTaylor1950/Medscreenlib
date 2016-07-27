
Option Strict Off
'$Revision: 1.4 $
'$Author: taylor $
'$Date: 2006-11-01 08:48:16+00 $
'$Log: Class1.vb,v $
'Revision 1.4  2006-11-01 08:48:16+00  taylor
'Add machine name to log error
'
'Revision 1.0  2005-09-02 07:41:45+01  taylor
'Checked as glossary namespace moved to a separate file
'
'Revision 1.3  2005-02-24 07:43:05+00  taylor
'Check in to get back into repostory
'
'Revision 1.2  2004-05-10 07:36:23+01  taylor
'<>
'
'Revision 1.1  2004-05-06 14:14:27+01  taylor
'Editing of Quick Mail to exclude emailing to a blank reciepient
'
'Revision 1.1  2004-05-03 16:58:21+01  taylor
'<>




Imports System

Imports System.IO
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Security.Principal
Imports System.Security.Permissions

Imports System.Xml
Imports System.Xml.linq
Imports System.Xml.XPath
Imports System.Xml.Xsl
Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.VisualBasic
Imports System.Data.OleDb



'''<summary>
'''Generic class to hold code for various projects 
'''</summary>
''' <remarks>
''' This is a class containing a widerange of shared utility functions <para/>
''' in many respects it is like the Sample Manager Lib-utils library.
''' </remarks>
''' <seealso cref="Medscreenlib.Glossary.Glossary" />
Public Class Medscreen  'Routines that are specific to Medscreen

#Region "Constants"
    '''<summary>The double quote character "</summary>
    Public Const Dquote As String = Chr(34)

    '''<summary>Database version parts "</summary>
    Public Enum DatabaseVersionParts
        '''<summary>Whole of ID "</summary>
        Whole
        '''<summary>Major part "</summary>
        Major
        '''<summary>Minor part (second bit) "</summary>
        Minor
        '''<summary>Release (Third Part) "</summary>
        Release
    End Enum
#End Region

#Region "Shared Declarations"


    Private Shared pvtXMLDirectory As String = Constants.GCST_X_DRIVE & "\lab programs\xml\"
    Private Shared pvtXSLDirectory As String = Constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
    Private Shared pvtResultsDirectory As String = Constants.GCST_X_DRIVE & "\Lab Programs\ResultFiles\"
    Private Shared pvtInifiles As String = "inifiles\"
    Private Shared pvtErrorLoggingDirectory As String = MedscreenLib.Constants.GCST_X_DRIVE & "\Error\"
    Public Shared GCSTCollectionXML As String = MedscreenLib.Constants.GCST_X_DRIVE & "\CollectionXML\"


    Private Shared strLiveRoot As String = "\\john\live\"
    Private Shared strDevRoot As String = "\\john\dev\"
    Private Shared strtestRoot As String = Constants.GCST_X_DRIVE & "\dbtest\"

    Private Shared strStyleSheet As String = "\\EM01\intranet\MedScreen.css"

    Private Shared strTemplates As String = Constants.GCST_X_DRIVE & "\lab programs\dbreports\templates\"
    Private Shared blnNoLog As Boolean = True
#End Region

#Region "Enumerations"
    '''<summary>
    '''     Mytpes enumeration of types supported by the parameter retriever
    ''' </summary>
    Public Enum MyTypes
        '''<summary>String Return</summary>
        typString = 0
        '''<summary>Date Return</summary>
        typDate = 1
        '''<summary>Integer Return</summary>
        typeInteger = 2
        '''<summary>Boolean Return</summary>
        typBoolean = 3
        '''<summary>Boolean Return</summary>
        typItem = 4
        File = 5
    End Enum
#End Region

    Declare Sub OutputDebugString Lib "kernel32" Alias "OutputDebugStringA" (ByVal lpOutputString As String)


#Region "Shared"

#Region "Shared Functions"
#Region "Misc"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Spell out a string as text 
    ''' </summary>
    ''' <param name="inWord"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' If special characters are added to the default password list  then they need to be added to configuration file here SpellSpecials
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	02/01/2008	Created
    ''' <revision><modified>17-Apr-2012 06:48 Add special character handling </modified><Author>CONCATENO\taylor</Author></revision>
    ''' <revision><modified>01-Dec-2011 06:03 Space handling added</modified><Author>CONCATENO\Taylor</Author></revision>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function SpellItOut(ByVal inWord As String) As String
        Dim Numbers As String() = {"ZERO", "ONE", "TWO", "THREE", "FOUR", "FIVE", "SIX", "SEVEN", "EIGHT", "NINE"}
        Dim Letters As String() = {"ALFA", "BRAVO", "CHARLIE", "DELTA", "ECHO", "FOXTROT", "GOLF", "HOTEL", "INDIA", "JULIET", "KILO", "LIMA", "MIKE", "NOVEMBER", "OSCAR", "PAPA", "QUEBEC", "ROMEO", "SIERRA", "TANGO", "UNIFORM", "VICTOR", "WHISKY", "X-RAY", "YANKEE", "ZULU"}
        Dim Special As String = MedscreenlibConfig.Miscellaneous("SpellSpecials")
        If Special.Trim.Length = 0 Then Special = "@| [AT],#| [HASH]"
        Dim Specials As String() = Special.Split(New Char() {","})
        Dim strRet As String = ""
        Dim strTemp As String = inWord.ToUpper
        Dim i As Integer
        For i = 0 To strTemp.Length - 1
            Dim aChar As Char
            aChar = strTemp.Chars(i)
            If Char.IsDigit(aChar) Then
                Dim iCPos As Integer = Asc(aChar) - Asc("0")
                strRet += aChar & " " & Numbers(iCPos) & ", "
            ElseIf Char.IsLetter(aChar) Then
                Dim iCPos As Integer = Asc(aChar) - Asc("A")
                Dim bChar As Char = inWord.Chars(i)
                If Char.IsLower(bChar) Then
                    strRet += Char.ToLower(aChar) & " " & CStr(Letters.GetValue(iCPos)).ToLower & ", "
                Else
                    strRet += aChar & " " & Letters(iCPos) & ", "
                End If
            ElseIf aChar = " " Then
                strRet += aChar & " [SPACE], "
                'ElseIf aChar = "@" Then
                '    strRet += aChar & " [AT], "
                'ElseIf aChar = "#" Then
                '    strRet += aChar & " [HASH], "
            Else    'some other character
                Dim Found As Boolean = False
                For Each sChar As String In Specials

                    If sChar.Chars(0) = aChar Then  'Found
                        Found = True
                        Dim Elements As String() = sChar.Split(New Char() {"|"})
                        strRet += aChar & Elements(1) & ", "
                    End If
                Next
                If Not Found Then strRet += aChar & ", "

            End If
        Next
        If strRet.Length > 2 Then strRet = Mid(strRet, 1, strRet.Length - 2)
        Return strRet
    End Function


    '''<summary>
    ''' Add International Dialling code to phone numbers
    ''' </summary>
    ''' <param name='strPhone'>Phone number to add dialling code to</param>
    ''' <param name='intCountry'>International Dialling Code</param>
    ''' <returns>Formatted International number</returns>
    Public Shared Function AddCountry(ByVal strPhone As String, ByVal intCountry As Integer) As String
        If strPhone.Trim.Length = 0 Then
            Return ""
        Else
            If intCountry <= 1000 Then
                Return "+" & intCountry & " " & Trim(strPhone)
            Else
                Return "+1 " & Trim(strPhone)
            End If
        End If

    End Function

    '''<summary>
    ''' Adds a field to an update string
    ''' </summary>
    ''' <param name='strFieldName'>Name of teh field to add</param>
    ''' <param name='instance'>Instance to add to</param>
    ''' <param name='chTerminator'>Terminator to use</param>
    ''' <param name='intMaxLen'>Maximum length of variable</param>
    ''' <returns>Formatted string</returns>
    ''' <remarks>deprecated do not use</remarks>
    Public Shared Function AddUpdateField(ByVal strFieldName As String, _
           ByVal instance As Object, _
           ByVal chTerminator As Char, ByVal intMaxLen As Integer) As String
        Dim strRet As String = ""
        Dim strTemp As String
        Dim datTemp As Date

        Dim od As Object

        od = instance
        If TypeOf od Is String Then
            strTemp = CStr(od)
            If strTemp.Length > intMaxLen Then
                strTemp = Mid(strTemp, 1, intMaxLen)
            End If
            If strTemp.Trim.Length = 0 Then
                strRet = strFieldName & " = NULL" & chTerminator
            Else
                strTemp = FixQuotes(strTemp)
                strRet = strFieldName & " = '" & strTemp & "'" & chTerminator
            End If
        ElseIf TypeOf od Is Date Then
            datTemp = CDate(od)
            strRet = strFieldName & " = TO_DATE('" & datTemp.ToString("ddMMyyyy HHmm") & "','DDMMYYYY HH24mi')" & chTerminator
        ElseIf TypeOf od Is Boolean Then
            strTemp = "F"
            If CType(od, Boolean) Then strTemp = "T"
            strRet = strFieldName & " = '" & strTemp & "'" & chTerminator
        Else
            If od Is System.DBNull.Value Then
                strRet = strFieldName & " = NULL " & chTerminator
            Else
                strRet = strFieldName & " = " & CStr(od) & chTerminator
            End If
        End If
        Return strRet

    End Function

    '''<summary>
    ''' Capitalise the supplied string
    ''' </summary>
    ''' <param name='strIn'>String to capitialise</param>
    ''' <returns>Capitalised string</returns>
    Public Shared Function Capitalise(ByVal strIn As String) As String
        Dim strTemp As String = ""
        Dim strT As String
        Dim i As Integer
        Dim iStart As Integer
        Dim blnUpper As Boolean
        Dim AlwaysUpper As String() = {"LLP", "BsC"}

        strIn = strIn.ToLower


        blnUpper = True
        While strIn.Length > 0
            strT = NextWord(strIn, "()[]{} ,;")
            'Check for Mc
            iStart = 0
            If (InStr(strT, "mc") = 1) AndAlso (strT.Length > 2) Then

                strTemp = "Mc" & UCase(strT.Chars(2))
                strT = Mid(strT, 4)
                iStart = 0
                blnUpper = False
            End If
            For i = iStart To strT.Length - 1
                If blnUpper Then
                    strTemp += Char.ToUpper(strT.Chars(i), System.Globalization.CultureInfo.CurrentCulture)
                Else
                    strTemp += strT.Chars(i)
                End If
                blnUpper = (InStr(" .('", strT.Chars(i)) <> 0) Or strT.Length <= 3 Or Array.IndexOf(AlwaysUpper, strT.Trim.ToUpper) >= 0
            Next
        End While
        Return strTemp
    End Function

    '''<summary>
    ''' Purge files from named directory
    ''' </summary>
    ''' <param name='Path'>Path to directory</param>
    ''' <param name='Before'>Date before which to purge</param>
    ''' <param name='Mask'>Mask of files which to purge</param>
    ''' <returns>Void</returns>
    Public Shared Function PurgeFiles(ByVal Path As String, _
    ByVal Before As Date, Optional ByVal Mask As String = "*.*") As Boolean

        Dim files As New DirectoryInfo(Path)
        Dim file As FileInfo

        Try
            For Each file In files.GetFiles(Mask)
                If Date.Compare(file.CreationTime, Before) = -1 Then
                    Debug.WriteLine(file.FullName)
                    file.Delete()
                End If
            Next
        Catch ex As Exception
        End Try

        Return True
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 		Compares version numbers in form a[.b[.c[.d[...]]]]
    ''' </summary>
    ''' <param name="Version1"></param>
    ''' <param name="Version2"></param>
    ''' <returns></returns>
    ''' <remarks>		If the number of points in one schema exceeds that in the other, the former
    '''	is considered greater.<p/>
    '''	Comparison starts with the first point and only continues until the values
    '''	differ.<p/>
    ''' The return value is of the form [Difference].[Point] where Point is the
    ''' number of the point at which versions differ.  If versions have different
    ''' numbers of points, Point will be 0.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CompareVersions(ByVal Version1 As String, ByVal Version2 As String) As Integer
        Dim verArray1 As String() = Version1.Split(New Char() {"."c})
        Dim verArray2 As String() = Version2.Split(New Char() {"."c})
        Dim intReturn As Integer = verArray1.Length - verArray2.Length
        Dim intIndex As Integer = 0

        While (intReturn = 0) AndAlso (intIndex < verArray2.Length)
            intReturn = (CInt(verArray1.GetValue(intIndex)) - CInt(verArray2.GetValue(intIndex)))
            intIndex += 1
        End While

        Return intReturn
    End Function

    '''<summary>
    ''' Function that returns the current logged in user
    ''' </summary>
    ''' <returns>Current logged in user as a string</returns>
    Public Shared Function WindowsUser() As String
        Dim strUser As String = System.Security.Principal.WindowsIdentity.GetCurrent.Name
        Dim intPos As Integer = InStr(strUser, "\")
        If intPos > 0 Then
            strUser = Mid(strUser, intPos + 1)
        End If
        Return strUser
    End Function

    '''<summary>
    ''' Return the mapping name of an object
    ''' </summary>
    ''' <param name='src'>Object to map</param>
    ''' <returns>Mapping name as string</returns>
    Public Shared Function GetMappingName(ByVal src As Object) As String
        Dim list As IList = Nothing
        Dim t As Type = Nothing

        If TypeOf (src) Is Array Then
            t = src.GetType()
            list = CType(src, IList)
        Else
            If TypeOf src Is System.ComponentModel.IListSource Then
                src = CType(src, System.ComponentModel.IListSource).GetList()
            End If

            If TypeOf src Is IList Then
                t = src.GetType()
                list = CType(src, IList)
            Else
                Return ""
            End If
        End If



        If TypeOf list Is System.ComponentModel.ITypedList Then
            Return (CType(list, System.ComponentModel.ITypedList).GetListName(Nothing))
        Else
            Return (t.Name)
        End If

    End Function

    '''<summary>
    ''' create a temporary filename for interval information
    ''' </summary>
    ''' <param name='Type'>Filename prefix</param>
    ''' <param name='Date1'>First Date to be used</param>
    ''' <param name='date2'>Second Date to be used</param>
    ''' <param name='extn'>Extension to be used</param>
    ''' <returns>A temporary filename</returns>
    ''' <remarks>This form is best used for interval data</remarks>
    Public Overloads Shared Function GetFileName(ByVal Type As String, ByVal Date1 As Date, ByVal date2 As Date, _
        Optional ByVal extn As String = "xml") As String

        Return GetTempPath() & Type & Date1.ToString("yyyyMMdd") & "-" & _
            date2.ToString("yyyyMMdd") & "." & extn
    End Function

    '''<summary>
    ''' create a temporary filename for information about or upto a date
    ''' </summary>
    ''' <param name='Type'>Filename prefix</param>
    ''' <param name='Date1'>First Date to be used</param>
    ''' <param name='extn'>Extension to be used</param>
    ''' <returns>A temporary filename</returns>
    ''' <remarks>This form is best used for information upto or on a specific date</remarks>
    ''' <revision><modified>12-Jan-2012 14:25 changed date format if hour is not zero</modified><Author>CONCATENO\taylor</Author></revision>
    Public Overloads Shared Function GetFileName(ByVal Type As String, ByVal Date1 As Date, _
Optional ByVal extn As String = "xml") As String
        Dim strRet As String = GetTempPath() & Type
        If Date1.Hour = 0 Then
            strRet += Date1.ToString("yyyyMMdd")
        Else
            strRet += Date1.ToString("yyyyMMddHHmmss")
        End If
        Return strRet & "." & extn
    End Function

    '''<summary>
    ''' create a temporary filename for interval information
    ''' </summary>
    ''' <param name='Type'>Filename prefix</param>
    ''' <param name='Value'>First Date to be used</param>
    ''' <param name='extn'>Extension to be used</param>
    ''' <returns>A temporary filename</returns>
    ''' <remarks>This form is the most generic</remarks>
    Public Overloads Shared Function GetFileName(ByVal Type As String, ByVal Value As String, _
Optional ByVal extn As String = "xml") As String
        Return GetTempPath() & Type & Value & _
            "." & extn
    End Function


    '''<summary>
    ''' create a filename building it from various parameters
    ''' </summary>
    '''     ''' <param name='newname'>the prefix min body of name</param>
    ''' <param name='Suffix'>the suffix to be used as an extension</param>
    ''' <param name='StartAt'>the number to start at </param>
    ''' <param name='AltSuffix'>alternate suffix</param>
    ''' <returns>A new filename</returns>
    ''' <remarks>Reporter will change a filename after it completes to sent or all_done<para/>
    ''' in order to cope with this and stop overwrites the alternate suffix should be set to the completed name</remarks>
    Public Overloads Shared Function GetNextFileName(ByVal newname As String, ByVal Suffix As String, _
    Optional ByVal StartAt As Integer = 0, Optional ByVal AltSuffix As String = "") As String

        Dim Ext As Integer
        Dim File_To_Find As String

        If AltSuffix.Trim.Length > 0 Then
            Ext = StartAt
            File_To_Find = newname & "-" & CStr(Ext).Trim & "." & AltSuffix.Trim()
            While IO.File.Exists(File_To_Find)
                Ext += 1
                File_To_Find = newname & "-" & CStr(Ext).Trim & "." & AltSuffix.Trim()
            End While
            File_To_Find = newname & "-" & CStr(Ext).Trim & "." & Suffix.Trim()
            Return (File_To_Find)
            Exit Function
        End If
        Ext = StartAt
        File_To_Find = newname & "-" & CStr(Ext).Trim & "." & Suffix.Trim()
        While IO.File.Exists(File_To_Find)
            Ext += 1
            File_To_Find = newname & "-" & CStr(Ext).Trim & "." & Suffix.Trim()
        End While
        Return (File_To_Find)


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Overload that should provide incremental filenumbers
    ''' </summary>
    ''' <param name="InDirectory"></param>
    ''' <param name="newname"></param>
    ''' <param name="Suffix"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function GetNextFileName(ByVal InDirectory As String, ByVal newname As String, ByVal Suffix As String) As String

        Dim Ext As Integer
        Dim File_To_Find As String

        'If AltSuffix.Trim.Length > 0 Then
        File_To_Find = newname & "-*" '& CStr(Ext).Trim & "." & AltSuffix.Trim()
        Dim existingFiles As String() = Directory.GetFiles(InDirectory, File_To_Find)
        'While IO.File.Exists(File_To_Find)
        Ext = existingFiles.Length
        File_To_Find = newname & "-" & CStr(Ext).Trim & "." & Suffix.Trim()
        'End While
        Return (InDirectory & "\" & File_To_Find)


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Try and convert a phone number to international standards.
    ''' </summary>
    ''' <param name="PhoneNumber"></param>
    ''' <param name="Country"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FixUpPhone(ByVal PhoneNumber As String, ByVal Country As Integer) As String

        Dim strRet As String = PhoneNumber
        If Country <= 0 Then
            Return strRet
            Exit Function
        End If
        strRet = StripCountry(PhoneNumber, CStr(Country))
        strRet = AddCountry(strRet, Country)
        Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Given a string and positioned beyond a possible string strip the text out of it
    ''' </summary>
    ''' <param name="Ipos">Start Position</param>
    ''' <param name="SF">Input String</param>
    ''' <param name="Direction">Direction of movement, defaults to downwards (-1)</param>
    ''' <returns></returns>
    ''' <remarks>This routine is primarily used to strip names embedded in filenames used for temporary XML files. <para/>
    ''' This routine enables the name to be stripped out and used.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/12/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function StripAlpha(ByVal Ipos As Integer, ByVal SF As String, Optional ByVal Direction As Integer = -1, Optional ByVal skiptochar As Boolean = False, Optional ByVal AlsoThese As String = "") As String
        Dim ij As Integer
        If SF.Trim.Length = 0 Then Return "" : Exit Function
        If Ipos > SF.Length + 1 Then Ipos = SF.Length + 1
        If Direction = -1 AndAlso Ipos <= 1 Then Direction = 1
        ij = Ipos + Direction
        'Deal with skipping 

        Dim aChar As Char = CChar(Mid(SF, ij, 1))
        If skiptochar Then
            If Direction = 1 Then
                While Not Char.IsLetter(aChar) AndAlso ij <= SF.Length
                    ij += 1
                    aChar = CChar(Mid(SF, ij, 1))
                End While
            Else
                While Not Char.IsLetter(aChar) AndAlso ij > 1
                    ij -= 1
                    aChar = CChar(Mid(SF, ij, 1))
                End While

            End If
        End If
        Dim strProgMan As String = ""
        If Direction = -1 Then  'Going backwards need to get to start and then go forwards
            While (Char.IsLetter(aChar) OrElse InStr(AlsoThese, aChar) > 0) AndAlso ij > 1
                ij -= 1
                aChar = CChar(Mid(SF, ij, 1))
            End While
            While ij < Ipos - 1
                ij += 1
                strProgMan += Mid(SF, ij, 1)
            End While
        Else
            While (Char.IsLetter(aChar) OrElse InStr(AlsoThese, aChar) > 0) AndAlso ij <= SF.Length
                strProgMan += Mid(SF, ij, 1)
                ij += 1
                aChar = CChar(Mid(SF, ij, 1))
            End While
        End If

        Return strProgMan       'Return Text 
    End Function

    '''<summary>
    ''' Remove country info from phone number
    ''' </summary>
    ''' <param name='strPhone'>Supplied phone number</param>
    ''' <returns>Corrected Phone number</returns>
    Public Shared Function StripCountry(ByVal strPhone As String, Optional ByVal CountryID As String = "") As String
        Dim intpos As Integer
        Dim strPreserve As String = Trim(strPhone)
        strPhone = Trim(strPhone)               'Convert Phone number

        If Len(strPhone) <= 3 Then              'If shorter than 3 characters can't be international coded
            strPhone = ""
            Return Trim(strPhone)
        Else                                    'May have international code
            intpos = InStr(strPhone, " ")
            While (Mid(strPhone, 1, 1) = "+") And (strPhone.Length > 3) And intpos > 0
                intpos = InStr(strPhone, " ")
                If intpos <> 0 Then
                    strPhone = Mid(strPhone, intpos + 1)
                End If
            End While
            'Check new phone length, if no International code exit
            If strPhone.Length <= 3 Then
                strPhone = strPreserve
            End If

            If CountryID.Trim.Length > 0 Then       'we need to check for 00XX format
                If CountryID = "1002" Then CountryID = "1"
                If InStr(strPhone, "00" & CountryID.Trim) = 1 Then
                    strPhone = Mid(strPhone, CountryID.Trim.Length + 3)
                End If
                If InStr(strPhone, "00 " & CountryID.Trim) = 1 Then
                    strPhone = Mid(strPhone, CountryID.Trim.Length + 4)
                End If
            End If
            Return strPhone
        End If

    End Function

    Public Shared Function GetExtnApp(ByVal Extn As String) As String
        Dim areg As Microsoft.Win32.RegistryKey
        Dim apath As String = ""
        Try
            areg = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(Extn) '\Microsoft\windows\current version\app paths\" & app)
            areg = areg.OpenSubKey("OpenWithList")
            Dim names As String() = areg.GetSubKeyNames()
            If names.Length = 1 Then apath = names(0)
        Catch ex As Exception

        End Try
        Return apath
    End Function

    Public Shared Function GetInstallationDirectory(ByVal app As String) As String
        Dim areg As Microsoft.Win32.RegistryKey
        Dim apath As String = ""
        Try
            areg = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("software") '\Microsoft\windows\current version\app paths\" & app)
            areg = areg.OpenSubKey("Microsoft")
            areg = areg.OpenSubKey("Windows")
            areg = areg.OpenSubKey("currentversion")
            areg = areg.OpenSubKey("app paths")
            areg = areg.OpenSubKey(app)
            apath = areg.GetValue("path", "")
        Catch ex As Exception
        End Try
        Return apath
    End Function

    '''<summary>
    ''' Convert a string to a byte array
    ''' </summary>
    ''' <param name='st'>String to convert</param>
    ''' <returns>Array of byte</returns>
    ''' <remarks>Used in the main with in memory streams</remarks>
    Public Shared Function StringToByteArray(ByVal st As String) As System.Array
        Dim buffer As Array = Array.CreateInstance(GetType(Byte), st.Length)

        Dim i As Integer

        For i = 0 To st.Length - 1
            buffer.SetValue(CType(Asc(st.Chars(i)), Byte), i)
        Next

        Return buffer
    End Function


    Public Shared Function ByteArraytoString(ByVal dbytes As Array) As String
        ' VB.NET to convert a byte array to a string:

        Dim str As String
        Dim enc As New System.Text.UTF8Encoding()
        str = enc.GetString(dbytes)
        Return str
    End Function

#End Region

#Region "Password Security"

    '''<summary>
    ''' Change the users password
    ''' </summary>
    ''' <param name='NewPWD'>New Password</param>
    ''' <returns>Success or failure</returns>
    Public Shared Function ChangePassword(ByRef NewPWD As String) As Boolean
        Dim frm As New frmPassword()
        frm.LockUser = True
        frm.ChangePassword = True
        If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
            NewPWD = frm.txtNewPassword.Text
            Return True
        Else
            NewPWD = ""
            Return False
        End If

    End Function

    Public Shared Function FindUsers(ByVal stub As String) As String
        Dim UserList As String = ActiveDirectoryWrapper.PC.ADWrapper.FindUsers(stub)
        Return UserList
    End Function

    Public Shared Function GetLoginId(ByVal username As String) As String
        Dim strRet As String = ""
        Try
            Dim Wrapper As ActiveDirectoryWrapper.Concateno.DirectoryEntry = ActiveDirectoryWrapper.PC.ADWrapper.FindConcatenoUser(username)
            strRet = Wrapper.UserID
        Catch ex As Exception
        End Try
        Return strRet
    End Function

    '''<summary>
    ''' Check a password to see if it is valid
    ''' </summary>
    ''' <param name='UserName'>Windows Identity</param>
    ''' <param name='Password'>Password to check</param>
    ''' <returns>True If password is correct for user name</returns>
    Public Overloads Shared Function CheckPassword(ByVal UserName As String, ByVal Password As String, ByRef ErrorText As String) As Boolean
        'Dim ns As New ts01.Service1()
        Dim blnRet As Boolean
        Try
            blnRet = ActiveDirectoryWrapper.PC.ADWrapper.IsUserValidThrow(UserName, Password)
            'Dim strRet As String = ns.CheckPassword(UserName, Password)
            'ErrorText = strRet
        Catch ex As Exception
            blnRet = False
            Medscreen.LogAction("Password check failed error :- " & ex.ToString)

        End Try
        Return blnRet

        'Return False


    End Function
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="UserName"></param>
    ''' <param name="Password"></param>
    ''' <param name="ErrorText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><modified>13-Mar-2015 09:49</modified><Author>CONCATENO\taylor</Author><REV>Added support for URL to be changed</REV></revision></revisionHistory>
    Public Overloads Shared Function CheckPasswordBySoap(ByVal UserName As String, ByVal Password As String, ByRef ErrorText As String) As Boolean
        Dim ns As New ts01.Service1(Support.Support.PasswordSOAPServiceURL)
        Dim blnRet As Boolean = True
        Try
            'blnRet = ActiveDirectoryWrapper.PC.ADWrapper.IsUserValidThrow(UserName, Password)
            ns.Timeout = 60000                                'Set Timeout to 60,000 milliseconds or 1 minute

            Dim strRet As String = ns.CheckPassword(UserName, Password)
            blnRet = (strRet.ToUpper = "TRUE")
            ErrorText = strRet
        Catch ex As System.Net.WebException
            blnRet = False
            If ex.Status = Net.WebExceptionStatus.Timeout Then
                Medscreen.LogAction("Password check failed error :- Timeout exception")
                Throw New MedscreenExceptions.MedscreenLoginCheckTimeoutException
            End If
        Catch ex As Exception
            blnRet = False
            Medscreen.LogAction("Password check failed error :- " & ex.ToString)

        End Try
        Return blnRet

        'Return False


    End Function

    '''<summary>
    ''' Check a password to see if it is valid
    ''' </summary>
    ''' <param name='changeUser'>Allow user to be changed</param>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/10/2005]</date><Action>Changed to not use encrypted password service</Action></revision>
    ''' </revisionHistory>
    ''' <returns>True If password is correct for user name</returns>
    Public Overloads Shared Function CheckPassword(Optional ByVal changeUser As Boolean = False) As Boolean
        Dim ns As New ts01.Service1()

        Try
            Dim frm As New frmPassword()
            frm.LockUser = Not changeUser
            frm.ChangePassword = False
            Dim strRet As String = ""
            If frm.ShowDialog = Windows.Forms.DialogResult.OK Then
                Medscreen.LogAction("Attempting to check password for " & frm.UserName)
                If frm.Password.Length > 0 And frm.UserName.Length > 0 Then
                    strRet = ns.CheckPassword(frm.UserName, frm.Password)
                    'LogError(strRet & " " & frm.UserName)
                    Return (strRet = "TRUE")
                Else
                    Return False
                End If
            Else
                Medscreen.LogAction("cancel check password for " & frm.UserName)
                Return False
            End If
            Return True
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
        End Try

    End Function


    '''<summary>
    ''' Check whether a username/password pair is correct
    '''</summary>
    ''' <remarks>
    ''' if DOMAIN is omitted, it uses the local account database 
    ''' and then asks trusted domains to search their account databases
    ''' until it finds the account or the search is exhausted
    ''' use DOMAIN="." to search only the local account database
    ''' IMPORTANT: works only under Windows 2003
    ''' </remarks>
    ''' <param name='userName'>Name of user</param>
    ''' <param name='password'>Password of user</param>
    ''' <param name='domain'>Login domain</param>
    ''' <returns>Error String or "TRUE"</returns>
    <PermissionSetAttribute(SecurityAction.Demand, Name:="FullTrust")> _
 Public Shared Function CheckWindowsUser(ByVal userName As String, ByVal password As String, _
     Optional ByVal domain As String = "") As String
        Const LOGON32_PROVIDER_DEFAULT As Integer = 0&
        'Const LOGON32_LOGON_NETWORK As Integer = 3&

        Const LOGON32_LOGON_INTERACTIVE As Integer = 2

        'Const LOGON32_PROVIDER_WINNT35 As Integer = 1


        Dim hToken As New IntPtr(0)


        ' provide a default for the Domain name
        'If domain.Length = 0 Then domain = Nothing
        ' check the username/password pair using LOGON32_LOGON_NETWORK delivers the 
        ' best performance

        hToken = IntPtr.Zero

        Dim ilong As Long = RevertToSelf()

        Dim returnValue As Boolean = LogonUser(userName, domain, password, _
         LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken)

        'ret = LogonUser(userName, domain, password, LOGON32_LOGON_NETWORK, _
        '      LOGON32_PROVIDER_WINNT35, hToken)
        'a non-zero value means success
        Dim strRet As String = ""
        If Not returnValue Then
            Dim ret As Integer = Marshal.GetLastWin32Error()
            strRet = "LogonUser failed with error code : " & ret
            strRet += " " & GetErrorMessage(ret)
        Else
            strRet = "TRUE"
        End If
        '
        Return strRet
    End Function

    '''<summary>
    ''' Encrypt/Decrypt string simplistic approach
    ''' </summary>
    ''' <param name='Pwd'>String to encrypt</param>
    ''' <returns>Encypted decrypted string</returns>
    Public Shared Function DoCrypt(ByVal Pwd As String) As String
        Dim Pos As Integer, iChar As Integer
        For Pos = 1 To Len(Pwd)
            Mid(Pwd, Pos, 1) = Chr(Asc(Mid(Pwd, Pos, 1)) Xor (iChar + 65))
            iChar = (iChar + 1) Mod 25
        Next
        DoCrypt = Pwd
    End Function



#End Region

#Region "Crystal Reports"
    'Public Shared Sub ExportToDisk(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal opt As CrystalDecisions.Shared.ExportFormatType, ByVal destination As String)
    '    Try
    '        Dim expOpt As CrystalDecisions.Shared.ExportOptions = cr.ExportOptions
    '        expOpt.ExportFormatType = opt
    '        If opt = CrystalDecisions.Shared.ExportFormatType.ExcelRecord Or opt = CrystalDecisions.Shared.ExportFormatType.Excel Then
    '            'Dim ExclOptions As CrystalDecisions.Shared.ExcelDataOnlyFormatOptions = CrystalDecisions.Shared.ExportOptions.CreateDataOnlyExcelFormatOptions
    '            'ExclOptions.ExportPageHeaderAndPageFooter = True
    '            'ExclOptions.SimplifyPageHeaders = True
    '            'expOpt.ExportFormatType = CrystalDecisions.Shared.ExportFormatType.ExcelRecord
    '            'expOpt.ExportFormatOptions = ExclOptions

    '        End If
    '        expOpt.ExportDestinationType = CrystalDecisions.[Shared].ExportDestinationType.DiskFile
    '        Dim diskexport As New CrystalDecisions.Shared.DiskFileDestinationOptions()
    '        diskexport.DiskFileName = destination
    '        expOpt.DestinationOptions = diskexport
    '        cr.ExportToDisk(opt, destination)
    '    Catch ex As Exception
    '        LogError(ex, , "Exporting -" & destination & "-" & cr.FilePath)
    '    End Try
    'End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Export as crystal report to appropriate file type
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <param name="SendMethod"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Shared Function ExportReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal SendMethod As Constants.SendMethod) As String
        'Dim tmpFileName As String
        'If SendMethod = Constants.SendMethod.Email Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "DOC")
        '    ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.WordForWindows, tmpFileName)
        'ElseIf SendMethod = Constants.SendMethod.PDF Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "PDF")
        '    ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, tmpFileName)
        'ElseIf SendMethod = Constants.SendMethod.Excel Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "XLS")
        '    ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.ExcelRecord, tmpFileName)
        'ElseIf SendMethod = Constants.SendMethod.RTF Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "RTF")
        '    ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.RichText, tmpFileName)
        'ElseIf SendMethod = Constants.SendMethod.HTML Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "HTM")
        '    ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpFileName)
        'End If
        Return CrystalSupport.ExportReport(cr, SendMethod)


    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set up a crystal report to print
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <param name="PrinterName"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Overloads Shared Function PrintReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal PrinterName As String) As Boolean

        cr.PrintOptions.PrinterName = PrinterName
        cr.PrintToPrinter(1, True, 0, 0)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Print Crystal report to Windows Default Printer
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	06/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Overloads Shared Function PrintReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument) As Boolean

        Dim PrinterName As String

        Dim prtdoc As New System.Drawing.Printing.PrintDocument()
        'Set to windows default printer
        PrinterName = prtdoc.PrinterSettings.PrinterName
        'Print to default printer
        cr.PrintOptions.PrinterName = PrinterName
        cr.PrintToPrinter(1, True, 0, 0)
    End Function


#End Region

#Region "Date and Time"
    '''<summary>
    ''' Get the time part of a date object
    ''' </summary>
    ''' <param name='datin'>Date to get time part from</param>
    ''' <returns>Date containg only the time part</returns>
    Public Shared Function ExtractTime(ByVal datin As Date) As Date
        Return TimeSerial(datin.Hour, datin.Minute, datin.Second)
    End Function

    '''<summary>
    ''' Convert a date and time to a date
    ''' </summary>
    ''' <param name='datIn'>Date to truncate</param>
    ''' <returns>Truncated date</returns>
    Public Shared Function TruncateDate(ByVal datIn As Date) As Date
        Return DateSerial(datIn.Year, datIn.Month, datIn.Day)
    End Function


#End Region

#Region "Control Functions"
    '''<summary>
    ''' Fill a combo box with the contents of a section in an inifile
    ''' </summary>
    ''' <param name='combo'>Combo box to fill</param>
    ''' <param name='IniFile'>Ini File to use</param>
    ''' <param name='Section'>Section in inifile to use</param>
    ''' <returns>True if succesful</returns>
    Public Overloads Shared Function FillList(ByVal combo As Windows.Forms.ComboBox, _
     ByVal IniFile As String, ByVal Section As String) As Boolean
        Dim inFile As IniFile.IniFiles
        Dim mCol As Collection
        Dim i As Integer

        inFile = New IniFile.IniFiles()
        inFile.FileName = IniFile

        mCol = inFile.ReadSection(Section)

        combo.Items.Clear()

        For i = 1 To mCol.Count
            combo.Items.Add(mCol.Item(i))
        Next

    End Function

    Public Overloads Shared Function FillList(ByVal combo As Windows.Forms.ComboBox, _
    ByVal List As Specialized.StringCollection) As Boolean
        
        combo.Items.Clear()

        For Each item As String In List
            combo.Items.Add(item)
        Next

    End Function

    '''<summary>
    ''' Set the index in a combo box to the supplied string
    ''' </summary>
    ''' <param name='Combo'>Combo box to set position</param>
    ''' <param name='Item'>Item (string) to find</param>
    ''' <returns>void</returns>
    Public Shared Function SetListIndex(ByVal Combo As Windows.Forms.ComboBox, _
    ByVal Item As String) As Boolean
        Dim i As Integer
        Dim myString As String



        Combo.SelectedIndex = -1
        For i = 0 To Combo.Items.Count - 1
            myString = CStr(Combo.Items(i))
            If myString.Trim.ToUpper = Item.Trim.ToUpper Then
                Combo.SelectedIndex = i
            End If
        Next

    End Function


#End Region

#Region "Translators"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Convert query by replacing parameters
    ''' </summary>
    ''' <param name="InputString"></param>
    ''' <param name="p1"></param>
    ''' <param name="p2"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/04/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function ExpandQuery(ByVal InputString As String, _
        Optional ByVal p1 As String = "", _
        Optional ByVal p2 As String = "", _
    Optional ByVal p3 As String = "") As String
        Dim tmpString As String = InputString
        If p1.Trim.Length > 0 Then
            tmpString = ReplaceString(tmpString, "{0}", p1)
        End If
        If p2.Trim.Length > 0 Then
            tmpString = ReplaceString(tmpString, "{1}", p2)
        End If
        If p3.Trim.Length > 0 Then
            tmpString = ReplaceString(tmpString, "{2}", p3)
        End If

        Return tmpString
    End Function


    '''<summary>
    ''' Translate a currency code into XML
    '''</summary>
    ''' <param name='Currency'>Currency to translate</param>
    ''' <returns>Translated currency value</returns>
    Public Shared Function CurrencyCodeXML(ByVal Currency As String) As String

        If Currency = "" Then
            Return "&amp;#163;"
        ElseIf Currency = "USD" Then
            Return "&amp;#036;"
        ElseIf Currency = "EUR" Then
            Return "&amp;#128;"
        Else
            Return "&amp;#163;"
        End If

    End Function

    '''<summary>
    ''' Translate a currency code in to HTML
    ''' </summary>
    ''' <param name='Currency'>Currency to translate</param>
    ''' <returns>Translated currency value</returns>
    Public Shared Function CurrencyCodeHTML(ByVal Currency As String) As String

        If Currency = "" Then
            Return "&#163;"
        ElseIf Currency = "USD" Then
            Return "&#036;"
        ElseIf Currency = "EUR" Then
            Return "&#128;"
        Else
            Return "&#163;"
        End If

    End Function

    '''<summary>
    ''' Convert a decimal integer to a hex string
    ''' </summary>
    ''' <param name='Value'>Input decimal integer</param>
    ''' <returns>Hexadecimal number</returns>
    Public Shared Function DecimalToHex(ByVal Value As Long) As String
        Dim Digits As String = "0123456789ABCDEF"

        Dim Remainder As Long
        Dim Number As Long
        Dim outString As String = ""

        If Value = 0 Then
            outString = "0"
        End If
        While Value > 0
            Number = CLng(Decimal.Truncate(CDec(Value / CLng(16))))
            Remainder = Value - (Number * 16)
            Value = Number
            'Console.WriteLine(Value / 16 & "," & Number & "," & Remainder)
            outString = Mid(Digits, CInt(Remainder) + 1, 1) & outString


        End While

        Return outString
    End Function

    '''<summary>
    ''' In HTML code fix up ampersands by duplicating them
    ''' </summary>
    ''' <param name='inpStr'>String rrquiring correction</param>
    ''' <returns>Corrected string</returns>
    ''' <remarks>In HTML code the &amp; character is treated as an escape character<para/>
    ''' so every instance needs to be replaced with a metacharacter set</remarks>
    Public Shared Function FixAmpersands(ByVal inpStr As String) As String
        Dim tmpStr As String
        Dim intPos As Integer
        'Dim intPos1 As Integer

        tmpStr = ""
        intPos = InStr(inpStr, "&")
        If intPos > 0 Then
            Do While intPos > 0
                tmpStr += Mid(inpStr, 1, intPos) & "amp;"
                inpStr = Mid(inpStr, intPos + 1)
                intPos = InStr(inpStr, "&")
            Loop
            tmpStr += inpStr
        Else
            Return inpStr
        End If
        Return tmpStr

    End Function

    '''<summary>
    ''' In XML code Fix &gt; characters  by replacing them HTML is a type of XML
    ''' </summary>
    ''' <param name='inpStr'>String rrquiring correction</param>
    ''' <returns>Corrected string</returns>
    ''' <remarks>In XML code the &gt; character has special meaning as the end of entity character<para/>
    ''' </remarks>
    Public Shared Function FixGreaterThen(ByVal inpStr As String) As String
        Dim tmpStr As String
        Dim intPos As Integer
        Dim intPos1 As Integer

        tmpStr = inpStr
        intPos = InStr(inpStr, ">")
        intPos1 = 1
        Do While intPos > 0
            tmpStr = Mid(tmpStr, intPos1, intPos - 1) & "&gt;" & Mid(tmpStr, intPos + 1)
            intPos1 = 1
            intPos = InStr(intPos1, tmpStr, ">")
        Loop
        Return tmpStr

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Fixup Xml 
    ''' </summary>
    ''' <param name="inpStr"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FixXML(ByVal inpStr As String) As String
        Return FixLessThan(FixGreaterThen(FixAmpersands(inpStr)))
    End Function

    '''<summary>
    ''' In XML code Fix &lt; characters  by replacing them HTML is a type of XML
    ''' </summary>
    ''' <param name='inpStr'>String rrquiring correction</param>
    ''' <returns>Corrected string</returns>
    ''' <remarks>In XML code the &lt; character has special meaning as the start of entity character<para/>
    ''' </remarks>
    Public Shared Function FixLessThan(ByVal inpStr As String) As String
        Dim tmpStr As String
        Dim intPos As Integer
        Dim intPos1 As Integer

        tmpStr = inpStr
        intPos = InStr(inpStr, "<")
        intPos1 = 1
        Do While intPos > 0
            tmpStr = Mid(tmpStr, intPos1, intPos - 1) & "&lt;" & Mid(tmpStr, intPos + 1)
            intPos1 = 1
            intPos = InStr(intPos1, tmpStr, "<")
        Loop
        Return tmpStr

    End Function


    '''<summary>
    ''' double ' characters
    ''' </summary>
    ''' <param name='inpStr'>String requiring correction</param>
    ''' <returns>Corrected string</returns>
    ''' <remarks>When writing to databases the ' needs to be doubled up, <para/>
    ''' this may occur in names like O'Reilly
    ''' </remarks>
    Public Shared Function FixQuotes(ByVal inpStr As String) As String
        Dim tmpStr As String
        Dim intPos As Integer
        Dim intPos1 As Integer

        tmpStr = inpStr
        intPos = InStr(inpStr, "'")
        intPos1 = 1
        Do While intPos > 0
            tmpStr = Mid(tmpStr, 1, intPos) & "'" & Mid(tmpStr, intPos + 1)
            intPos1 = intPos + 2
            intPos = InStr(intPos1, tmpStr, "'")
        Loop
        Return tmpStr

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     Generates a random alphanumeric password of specifie length and format
    ''' </summary>
    ''' <param name="intMinLength">(Numeric) Minimum length of password</param>
    ''' <param name="intMaxLength">(Numeric) Max length of password (must be >= intMinLength)</param>
    ''' <param name="blnMixedCase">(Boolean) If FALSE passwords will be all lower case</param>
    ''' <param name="blnAlphaNumeric">(Boolean) TRUE to allow numbers in password.</param>
    ''' <param name="intMinNumbers">(Numeric) Min number of numeric characters to include.
    '''                      Ignored (see NOTES) if blnAlphaNumeric is FALSE.
    '''                      May not be > MaxLength / 2 (for performance reasons).</param>
    ''' <returns>ERROR if any parameters of wrong type, or out of range
    '''  STRING of Length between intMinLength and intMaxLength (inclusive).</returns>
    ''' <remarks>      If blnAlphNumeric is FALSE and intMinNumbers is 0, passwords will always be of
    '''  length intMinLength.  To create random length passwords between min and max,
    '''  set intMinNumbers to a value of around 1/4 max length.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GeneratePassword(ByVal intMinLength As Integer, _
    ByVal intMaxLength As Integer, ByVal blnMixedCase As Boolean, ByVal blnAlphaNumeric As Boolean, ByVal intMinNumbers As Integer) As String

        ' Make sure all values are in range }
        If (intMaxLength < intMinLength) Or (intMinNumbers > (intMaxLength / 2)) Then
            Return Nothing
        End If
        '{ Generate a seed for the random number generator }

        Dim strSeed As String = Now.ToString("yyMMddHHmmss")
        Dim sngSeed As Double = CDbl(strSeed)
        Randomize(sngSeed)

        Dim strPassword As String = ""
        Dim intNumberCount As Integer = 0
        '{ Create a password }
        While (strPassword.Length < intMinLength) And (intNumberCount < intMinNumbers)
            '{ sngAscii is a number in range 48 to 84 (48 + 36) }
            Dim intBase As Integer = Asc("0")
            Dim intZ As Integer = Asc("z")
            Dim intSpan As Integer = intZ - intBase
            Dim sngAscii As Char = Chr(CInt((Rnd() * intSpan) + intBase))
            '{ Ensure only characters 0-9 A-Z and a-z are used }
            If Char.IsLetterOrDigit(sngAscii) Then
                If Char.IsDigit(sngAscii) Then
                    If sngAscii <> "0" Then 'Exclude Zero confusion with O
                        intNumberCount = intNumberCount + 1
                        If blnAlphaNumeric Then
                            strPassword += sngAscii
                        End If
                    End If
                Else
                    If sngAscii = "O" Then sngAscii = "o"c 'convert O's to lower case
                    If sngAscii = "i" Then sngAscii = "I"c 'convert i's to upper case
                    If sngAscii = "l" Then sngAscii = "L"c 'convert l's to upper case
                    strPassword += sngAscii
                End If
            End If
            '    { If the password has gotten too long (because we're waiting for enough numbers)
            'Truncate it and re-calculate the number of number characters in the string }
            If strPassword.Length > intMaxLength Then
                strPassword = Right(strPassword, intMaxLength)
                '{ Do not recalculate number of number characters for alpha only passwords,
                ' otherwise generation of variable length passwords may take some time! }
                If blnAlphaNumeric Then
                    intNumberCount = CountNumCharsInString(strPassword)
                End If
            End If
        End While
        If Not blnMixedCase Then
            strPassword = strPassword.ToLower
        End If
        Return strPassword
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' CountNumCharsInString - returns the number of characters in strExpression that are numeric 
    ''' </summary>
    ''' <param name="strIn"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/03/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function CountNumCharsInString(ByVal strIn As String) As Integer
        Dim i As Integer
        Dim intCharCount As Integer = 0

        For i = 0 To strIn.Length - 1
            If Char.IsDigit(strIn.Chars(i)) Then intCharCount += 1
        Next
        Return intCharCount
    End Function

    'GetErrorMessage formats and returns an error message
    'corresponding to the input errorCode.
    '''<summary>
    ''' Translate a Hexadecimal error Code into human readable form
    ''' </summary>
    ''' <param name='errorCode'>Hex error to translate</param>
    ''' <returns>Translated error message</returns>
    Public Shared Function GetErrorMessage(ByVal errorCode As Integer) As String
        Dim FORMAT_MESSAGE_ALLOCATE_BUFFER As Integer = &H100
        Dim FORMAT_MESSAGE_IGNORE_INSERTS As Integer = &H200
        Dim FORMAT_MESSAGE_FROM_SYSTEM As Integer = &H1000

        Dim messageSize As Integer = 255
        Dim lpMsgBuf As String = ""
        Dim dwFlags As Integer = FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS

        Dim ptrlpSource As IntPtr = IntPtr.Zero
        Dim prtArguments As IntPtr = IntPtr.Zero

        Dim retVal As Integer = FormatMessage(dwFlags, ptrlpSource, errorCode, 0, lpMsgBuf, _
            messageSize, prtArguments)
        If 0 = retVal Then
            Throw New Exception("Failed to format message for error code " + errorCode.ToString() + ". ")
        End If

        Return lpMsgBuf
    End Function 'GetErrorMessage


    '''<summary>
    ''' Convert XML to HTML as a string using a style sheet using an in memeory method
    ''' </summary>
    ''' <param name='source'>Source XML data </param>
    ''' <param name='StyleSheetId'>XSL transform sheet to be used</param>
    ''' <param name='opt'>place holder variable use 0, this is so the signature is different</param>
    ''' <param name='Args'>Arguments to be supplied to the stylesheet</param>
    ''' <param name='Path'>Path to style sheet</param>
    ''' <returns>HTML data as a string</returns>
    ''' <remarks>This method is used to convert XML data to formatted data for use in a report
    ''' the intention is to have a method that will rapidly produce reports and can the apperance can be easily changed </remarks>
    Public Overloads Shared Function ResolveStyleSheet(ByVal source As String, _
    ByVal StyleSheetId As String, ByVal opt As Integer, _
    Optional ByVal Args As Xml.Xsl.XsltArgumentList = Nothing, _
    Optional ByVal Path As String = "\\corp.concateno.com\medscreen\common\Lab Programs\Transforms\XSL\") As String


        'Dim StyleSheet As String = medscreenlib.constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
        Dim StyleSheet As String = Path & StyleSheetId
        Dim strTemp As String
        Try
            Dim st As New IO.MemoryStream(Medscreen.StringToByteArray("<?xml version=""1.0"" encoding=""Windows-1252"" ?>" & source))
            'Create a memeory stream using the supplied XML
            Dim doc As XPathDocument = New XPathDocument(st)    'Fill an Xpath document with the XML
            st.Close()                              'Close stream
            st = Nothing                            'Kill it

            Dim xslt As Xml.Xsl.XslTransform = New Xml.Xsl.XslTransform()   'Create an xlst transformer
            Dim resolver As XmlUrlResolver = New XmlUrlResolver()

            xslt.Load(StyleSheet, resolver)         'Load the style sheet using the resolver to find the necessary namespaces

            st = New IO.MemoryStream(0)
            Dim writer As XmlTextWriter = New XmlTextWriter(st, System.Text.Encoding.UTF8)
            xslt.Transform(doc, Args, writer) 'Transform the document to HTML output to writer in memory 

            doc = Nothing

            st.Position = 0                             'Move stream to start 
            Dim Readr As New IO.StreamReader(st)

            strTemp = Readr.ReadToEnd()                 'Read memory based streem 


            Readr.Close()                               'Tidy up variables used
            Readr = Nothing
            st.Close()
            st = Nothing
        Catch ex As Exception
            strTemp = ex.Message
            'Throw ex
        End Try
        Return strTemp                              'Return transformed XML as HTML


        'Return "Function Not Available yet"

    End Function



    
#End Region


#Region "Information Gathering"


    '''<summary>
    '''  Display html using a specific form for the purpose saves shelling to internet explorer
    ''' </summary>
    ''' <param name='html'> HTML to display</param>
    ''' <returns>Result from the HTML viewwer Form </returns>
    ''' <remark>This function uses the HTML Viewer Form</remark>
    Public Overloads Shared Function ShowHtml(ByVal html As String) As System.Windows.Forms.DialogResult
        Dim frmHTML As New frmHTMLView()
        frmHTML.Document() = html
        Return frmHTML.ShowDialog()
    End Function


    Public Overloads Shared Function ShowPDF(ByVal PDFFile As String, ByVal FormTitle As String) As System.Windows.Forms.DialogResult
        Dim myProcess As Process = New Process()
        'set program to use and arguments and window style then start process

        myProcess.StartInfo.FileName = PDFFile
        myProcess.StartInfo.Arguments = ""
        myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
        myProcess.StartInfo.CreateNoWindow = True
        ' myProcess.StartInfo.UseShellExecute = False                 'Required for redirecting
        'MedscreenLib.Medscreen.LogAction("starting process - " & strBlat)

        myProcess.Start()
    End Function

    Public Overloads Shared Function ShowHtml(ByVal html As String, _
                  ByVal DialogMessage As String, ByVal buttons As MsgBoxStyle, Optional ByVal Title As String = "") As System.Windows.Forms.DialogResult
        Dim frmHTML As New frmHTMLView()
        frmHTML.Dialogmessage = DialogMessage
        frmHTML.DialogStyle = buttons
        frmHTML.DialogTitle = Title
        frmHTML.Document() = html
        Return frmHTML.ShowDialog()

    End Function

    Public Shared Function AppPath() As String
        Dim fi As IO.FileInfo = New System.IO.FileInfo(System.Windows.Forms.Application.ExecutablePath)
        Dim strPath As String = fi.DirectoryName
        Return strPath

    End Function

    Public Shared Function HelpPath() As String
        Dim aHelpPath As String = MedscreenlibConfig.Paths("HELP")
        Return aHelpPath
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the version of the schema in use
    ''' </summary>
    ''' <param name="part"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function DataBaseVersion(Optional ByVal part As DatabaseVersionParts = DatabaseVersionParts.Whole) As String

        Dim strReturn As String = "0.0.0"

        Select Case part
            Case DatabaseVersionParts.Whole
                strReturn = CConnection.SchemaMajor.ToString.Trim & "." & _
                CConnection.SchemaMinor.ToString.Trim & "." & _
                CConnection.SchemaRelease.ToString.Trim
            Case DatabaseVersionParts.Major
                strReturn = CConnection.SchemaMajor.ToString
            Case DatabaseVersionParts.Minor
                'Dim intPos As Integer
                strReturn = CConnection.SchemaMinor.ToString
            Case DatabaseVersionParts.Release
                strReturn = CConnection.SchemaRelease.ToString
        End Select
        Return strReturn
    End Function

    '''<summary>
    ''' Find were the user wants output sent to
    ''' </summary>
    ''' <param name='outputMethod'>How the output is to be produced</param>
    ''' <param name='OutputDest'>Default of where the output should go to </param>
    ''' <returns>Output destination</returns>
    Public Shared Function GetOutputDestination(ByRef outputMethod As String, _
  Optional ByVal OutputDest As String = "") As String
        Try

            Dim frm As New frmOutputSelect()
            frm.OutputMethod = outputMethod
            frm.OutputDestination = OutputDest
            frm.ShowDialog()
            outputMethod = frm.OutputMethod
            Return frm.OutputDestination
        Catch ex As System.Exception
            ' Handle exception here
            MedscreenLib.Medscreen.LogError(ex, , "Exception log message")
            Return ""
        End Try
    End Function

    '''<summary>
    ''' Get a parameter from the user
    ''' </summary>
    ''' <param name='paramType'>Type of parameter</param>
    ''' <param name='paramname'>Description presented to the user of parameter</param>
    ''' <param name='Title'>Title bar displaye don form</param>
    ''' <param name='DefaultValue'>Default value presented to the user</param>
    ''' <param name='SingleLine'>Indicates whether the text box is a single or multi line display</param>
    '''     '''     ''' <returns>void</returns>
    Public Overloads Shared Function GetParameter(ByVal paramType As MyTypes, _
    ByVal paramname As String, Optional ByVal Title As String = "", _
        Optional ByVal DefaultValue As Object = Nothing, _
        Optional ByVal SingleLine As Boolean = True, _
        Optional ByVal ItemCollection As Collection = Nothing) As Object
        Dim myForm As New frmParameter()

        myForm.ParameterType = paramType
        myForm.ParameterName = paramname

        If paramType = MyTypes.typItem Then
            If ItemCollection Is Nothing Then Return Nothing : Exit Function
            myForm.Itemlist = ItemCollection
        End If
        myForm.TextBox1.Multiline = Not SingleLine

        If Not DefaultValue Is Nothing Then myForm.Value = DefaultValue
        If Title.Trim.Length > 0 Then myForm.Text = Title
        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myForm.Value
        Else
            Return Nothing
        End If
    End Function

    Public Overloads Shared Function GetParameter(ByVal paramType As MyTypes, _
  ByVal paramname As String, ByVal ItemCollection As Glossary.PhraseCollection) As Object
        Dim myForm As New frmParameter()

        myForm.ParameterType = paramType
        myForm.ParameterName = paramname

        If paramType = MyTypes.typItem Then
            If ItemCollection Is Nothing Then Return Nothing : Exit Function
            myForm.PhraseList = ItemCollection
        End If
        myForm.TextBox1.Multiline = False

        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myForm.Value
        Else
            Return Nothing
        End If
    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a parameter from a phrase list
    ''' </summary>
    ''' <param name="PhraseList"></param>
    ''' <param name="ParamNAme"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function GetParameter(ByVal PhraseList As Glossary.PhraseCollection, _
    ByVal ParamNAme As String, Optional ByVal Value As String = "") As Object
        Dim myForm As New frmParameter()

        myForm.ParameterType = MyTypes.typItem
        myForm.ParameterName = ParamNAme

        myForm.cmbOption.DataSource = PhraseList
        myForm.cmbOption.ValueMember = "PhraseID"
        myForm.cmbOption.DisplayMember = "PhraseText"



        ' set position 
        If Value.Trim.Length > 0 Then
            myForm.cmbOption.SelectedValue = Value
        End If

        myForm.TextBox1.Multiline = False

        'If Not DefaultValue Is Nothing Then myForm.Value = DefaultValue
        'If Title.Trim.Length > 0 Then myForm.Text = Title
        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myForm.cmbOption.SelectedValue
        Else
            Return Nothing
        End If

    End Function

    Public Overloads Shared Function GetParameter(ByVal DT As DataTable, ByVal IdColumn As Integer, ByVal textColumn As Integer) As Object
        Dim myForm As New frmParameter()

        myForm.ParameterType = MyTypes.typItem
        myForm.ParameterName = DT.Columns(textColumn).ColumnName

        myForm.cmbOption.DataSource = DT
        myForm.cmbOption.ValueMember = DT.Columns(IdColumn).ColumnName
        myForm.cmbOption.DisplayMember = DT.Columns(textColumn).ColumnName
        'myForm.cmbOption.Sorted = True

        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myForm.cmbOption.SelectedValue
        Else
            Return Nothing
        End If

    End Function


    Public Overloads Shared Function GetParameter(ByVal e As Type, _
     ByVal ParamNAme As String) As Object
        Dim myForm As New frmParameter()

        myForm.ParameterType = MyTypes.typItem
        myForm.ParameterName = ParamNAme

        myForm.cmbOption.DataSource = [Enum].GetNames(e)
        'myForm.cmbOption.ValueMember = "PhraseID"
        'myForm.cmbOption.DisplayMember = "PhraseText"

        myForm.TextBox1.Multiline = False

        'If Not DefaultValue Is Nothing Then myForm.Value = DefaultValue
        'If Title.Trim.Length > 0 Then myForm.Text = Title
        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myForm.cmbOption.SelectedValue
        Else
            Return Nothing
        End If

    End Function


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Specialist form of get contract amendment that will remember entry 
    ''' </summary>
    ''' <param name="Remember"></param>
    ''' <param name="DefaultValue"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [06/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function GetContractAmendment(ByRef Remember As Boolean, ByVal DefaultValue As String) As String
        Dim myForm As New frmParameter()

        myForm.ParameterType = MyTypes.typString
        myForm.ParameterName = "Contract Amendment"

        myForm.TextBox1.Multiline = False

        myForm.Value = DefaultValue
        myForm.ShowRemember = True
        myForm.Text = "Contract Amendment"
        If myForm.ShowDialog = Windows.Forms.DialogResult.OK Then
            Remember = myForm.Remebered
            Return CStr(myForm.Value)
        Else
            Return Nothing
        End If

    End Function

    '''<summary>
    ''' Instantiate a text editor
    ''' </summary>
    ''' <param name='inText'>Input text for editor</param>
    ''' <param name='isReadOnly'>Is text read write or reead only</param>
    ''' <param name='width'>Width of text box default = 416</param>
    ''' <param name='Title'>Title on form</param>
    ''' <returns>Edited string</returns>
    Public Shared Function TextEditor(ByVal inText As String, _
 Optional ByVal isReadOnly As Boolean = False, _
 Optional ByVal width As Integer = 416, _
 Optional ByVal Title As String = "Editor") As String
        Dim myform As New frmEditText()

        myform.EditText = inText
        myform.Width = width
        myform.Text = Title
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            Return myform.EditText
        Else
            Return inText
        End If
    End Function


#End Region

#Region "Paths"

    '''<summary>
    ''' Location of Collman Ini file
    ''' </summary>
    ''' <returns>IniFile.IniFiles</returns>
    ''' <remarks>Location of file typically medscreenlib.constants.GCST_X_DRIVE &amp; '\lab programs\vbcollman\collman.ini</remarks>
    Public Shared Function CollmanIni() As IniFile.IniFiles
        Static myCollmanini As IniFile.IniFiles
        If myCollmanini Is Nothing Then
            myCollmanini = New IniFile.IniFiles()
            myCollmanini.FileName = MedscreenLib.Constants.GCST_X_DRIVE & "\lab programs\vbcollman\collman.ini"
        End If
        Return myCollmanini
    End Function

    '''<summary>
    ''' returns the location of the current windows temporary directory
    ''' </summary>
    ''' <returns>path to temporary directory</returns>
    Public Overloads Shared Function GetTempPath() As String
        Dim pth As String = "\\corp.concateno.com\medscreen\common\xmltemp\" & Now.ToString("yyyyMMddHHmmss")
        Return Path.GetDirectoryName(pth) & "\"
    End Function

    '''<summary>
    ''' Return a path to the root of the live file directory for emaple N:\
    ''' </summary>
    ''' <result>The path to the Live directory as a string</result>
    Public Shared Function LiveRoot() As String
        If MedscreenLib.CConnection.DatabaseInUse = CConnection.useDatabase.LIVE Then
            Return strLiveRoot
        ElseIf MedscreenLib.CConnection.DatabaseInUse = CConnection.useDatabase.DEV Then
            Return strDevRoot
        Else
            Return strtestRoot
        End If
    End Function

    '''<summary>
    ''' return the path of the word templates directory typically
    ''' "\\corp.concateno.com\medscreen\common\lab programs\dbreports\templates\"
    ''' </summary>
    ''' <result>The path to the Word Templates Directory as a string</result>
    Public Shared Function Templates() As String
        Return strTemplates
    End Function
#End Region

#Region "HTML"

    '''<summary>
    ''' Load a file of HTML and display it
    ''' </summary>
    ''' <param name='path'>Path and filename to the HTML document</param>
    ''' <result>Return from HTML Viewer</result>
    Public Shared Function LoadHtml(ByVal path As System.Uri) As System.Windows.Forms.DialogResult
        Dim frmhtml As New frmHTMLView()
        frmhtml.URL() = path
        Return frmhtml.ShowDialog()
    End Function

    '''<summary>
    ''' Return a standard HTML header saves duplicating code
    ''' </summary>
    ''' <param name='CloseHead'>Close header or leave it open</param>
    ''' <result>HTML document header</result>
    Public Shared Function HTMLHeader(Optional ByVal CloseHead As Boolean = True, Optional ByVal addstyle As Boolean = True) As String
        Dim strReturn As String = "<HTML><head><LINK REL=STYLESHEET TYPE=" & _
                Dquote & "text/css" & Dquote & " HREF = " & Dquote & _
                StyleSheet & Dquote & " Title=" & Dquote & _
                "Style" & Dquote & ">"
        If addstyle Then
            strReturn += HTMLStyle()
        End If
        If CloseHead Then
            strReturn += "</head>"
        End If
        Return strReturn
    End Function

    '''<summary>
    ''' A style section
    ''' </summary>
    ''' <result>A formatted style section</result>
    Public Shared Function HTMLStyle() As String
        Dim strReturn As String = "				<STYLE>"
        strReturn += "BODY.Plain {background:white; link:#99ccff; vlink:#99ccff; topmargin:0; " & _
            "margin-left: 10pt; "
        strReturn += "@page Section1" & _
            "{size:595.3pt 841.9pt;} }"
        strReturn += "TABLE           { display: table; margin-left: auto; margin-right: auto; border-color:black; border-style: solid; " & _
            "border: outset 2pt; empty-cells:show; " & _
            "border-collapse:collapse; }"

        strReturn += "TABLE.FixedWid  {table-layout : fixed; margin-left: auto;" & _
            "margin-right:auto; border-bottom-color:Gray; " & _
            "border-left-color:gray; border-right-color:White; border-top-color:white }"

        strReturn += "TABLE.FixedWidNB  {table-layout : fixed; margin-left: auto;" & _
            "margin-right: auto; border-style: none;  }"

        strReturn += "TH.NoBorder 	{border-style: none;    background:aliceblue;}"
        strReturn += "TH.NoBorderLA 	{border-style: none;  text-align: left;     background:aliceblue;}"
        strReturn += "TD.NbAL8	{ border-style: none; font-size: 8pt; text-align: Left; font-weight:normal; }"
        strReturn += "TD.NbAL8rd	{ border-style: none; font-size: 8pt; text-align: Left; font-weight:normal; background:salmon}"
        strReturn += "TD.NbAL80	{ border-style: none; font-size: 8pt; text-align: Left; font-weight:normal; background:white;}"
        strReturn += "TD.NbAL81	{ border-style: none; font-size: 8pt; text-align: Left; font-weight:normal; background:beige;}"
        strReturn += "TD.NbAL8gr	{ border-style: none; font-size: 8pt; text-align: Left; font-weight:normal;  background:gainsboro; }"
        strReturn += "DIV.Centred	{ text-align: center;  position:relative; font-size: 14pt; font-weight:bold;}"
        strReturn += "DIV.CentredSml	{ text-align: center;  position:relative; font-size: 8pt; font-weight:bold;}"

        strReturn += "</STYLE>"

        Return strReturn
    End Function

    '''<summary>
    ''' Header for Invoice table
    ''' </summary>
    ''' <returns>Invoice Table Header</returns>
    Public Shared Function InvoiceTableHeader() As String
        Return Constants.cstTableHead & _
            Constants.cstColTableHead & "200> Line Item </th>" & _
            Constants.cstColTableHead & "40> Qty</TH> " & _
            Constants.cstColTableHead & "80> Unit Price </TH>" & _
            Constants.cstColTableHead & "80> Total </TH></Thead>"
    End Function


    '''<summary>
    ''' Convert XML to HTML as a string using a style sheet using a file
    ''' </summary>
    ''' <param name='Filename'>Filename and path of file</param>
    ''' <param name='StyleSheetId'>XSL transform sheet to be used</param>
    ''' <param name='blnDelSource'>Delete source file after use</param>
    ''' <param name='Args'>Arguments to be supplied to the stylesheet</param>
    ''' <param name='Path'>Path to style sheet</param>
    ''' <returns>HTML data as a string</returns>
    ''' <remarks>This method is used to convert XML data to formatted data for use in a report
    ''' the intention is to have a method that will rapidly produce reports and can the apperance can be easily changed </remarks>
    Public Overloads Shared Function ResolveStyleSheet(ByVal Filename As String, _
        ByVal StyleSheetId As String, _
        Optional ByVal blnDelSource As Boolean = True, _
    Optional ByVal Args As Xml.Xsl.XsltArgumentList = Nothing, _
    Optional ByVal Path As String = "") As String



        Dim StyleSheet As String = MedscreenLib.Constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
        If Path.Trim.Length = 0 Then
            StyleSheet += StyleSheetId
        Else
            StyleSheet = Path & StyleSheetId
        End If
        Dim fileNameX As String = "C:\temp\" & Now.ToString("yyyyMMddhhmmss").Trim & ".html"

        Dim resolver As XmlUrlResolver = New XmlUrlResolver()           'Create a resolver
        Dim xslt As Xml.Xsl.XslTransform = New Xml.Xsl.XslTransform()   'Create a style sheet transformer
        xslt.Load(StyleSheet, resolver)

        Dim doc As Xml.XPath.XPathDocument = New Xml.XPath.XPathDocument(Filename)              'Read in XML into an Xpath document 
        Dim writer As XmlTextWriter = New XmlTextWriter(fileNameX, System.Text.Encoding.UTF8)   'Create a writer to output HTML

        xslt.Transform(doc, Args, writer)                               'Do the XLST transform in the style sheet 
        doc = Nothing                                                   'We don't need the xpath any longer

        writer.Flush()                                                  'Flush the resolved document
        writer.Close()                                                  'Close it 

        writer = Nothing
        xslt = Nothing

        If blnDelSource Then                                            'Keep or don't keep XML
            System.Threading.Thread.Sleep(500)
            System.IO.File.Delete(Filename)
        End If

        Dim Readr As New IO.StreamReader(fileNameX)                     'Create a reader

        Dim strTemp As String = Readr.ReadToEnd()                       'Fill a string with the HTML

        Readr.Close()                                                   'Close the reader
        Readr = Nothing
        System.Threading.Thread.Sleep(500)
        System.IO.File.Delete(fileNameX)                                'And kill the temporary file 

        Return strTemp

    End Function

#End Region

#Region "string support"

    Private Shared Function MatchWholeWordIsTerminator(ByVal term As Char) As Integer
        If Char.IsPunctuation(term) Then
            Return 1
        ElseIf Char.IsWhiteSpace(term) Then
            Return 1
        Else
            Return 0
        End If

    End Function
    Public Shared Function MatchWholeWord(ByVal Source As String, ByVal target As String) As Integer
        If Source.ToUpper.Contains(target.ToUpper) Then
            Dim intpos As Integer = InStr(Source, target, CompareMethod.Text)
            If intpos = 1 Then ' at start of line
                If Source.Length = target.Length Then
                    Return 1
                Else
                    Dim term As Char = Source.Chars(target.Length)
                    Return MatchWholeWordIsTerminator(term)
                End If
            Else
                Dim prev As Char = Source.Chars(intpos - 2)
                Dim termchar As Integer = MatchWholeWordIsTerminator(prev)
                If termchar = 0 Then
                    Return 0
                Else
                    If Source.Length = target.Length + intpos - 1 Then ' finshes sentenace
                        Return 1
                    Else
                        Dim term As Char = Source.Chars(target.Length + intpos - 1)
                        Return MatchWholeWordIsTerminator(term)
                    End If

                End If
            End If
        End If
    End Function

    '''<summary>
    ''' Return the next word from string
    ''' </summary>
    ''' <param name='InString'>Input String </param>
    ''' <param name='terms'>Terminators</param>
    ''' <returns>Next word in string</returns>
    Public Shared Function NextWord(ByRef InString As String, ByVal terms As String) As String
        Dim i As Integer
        Dim ch As Char
        Dim strRet As String = ""

        i = 0
        ch = InString.Chars(i)
        While (i < InString.Length - 1) And (InStr(terms, ch) = 0)
            strRet += ch
            i += 1
            ch = InString.Chars(i)
        End While
        strRet += ch
        i = i + 2
        InString = Mid(InString, i)
        Return strRet


    End Function

    '''<summary>
    ''' Remove characters from a string 
    ''' </summary>
    ''' <param name='inString'>Input String</param>
    ''' <param name='ch'>Character to remove</param>
    ''' <returns>Fixed string</returns>
    Public Overloads Shared Function RemoveChars(ByVal inString As String, ByVal ch As Char) As String
        Dim retStr As String = ""
        Dim i As Integer

        For i = 0 To inString.Length - 1
            If inString.Chars(i) <> ch Then
                retStr += inString.Chars(i)
            End If
        Next

        Return retStr


    End Function



    Public Shared Function StripCharsFromString(ByVal String1 As String, ByVal NotAllowed As String) As String
        'mezString1 string to have characters striped from
        'mezString2 string of characters that will be striped out of mezString1
        ' e.g. ? mezStripChrsFromString("Mr O'Grady 12345","1 3'2")
        ' returns MrOGrady45
        Dim strRet As String = ""

        For Each ch As Char In String1
            If InStr(NotAllowed, ch) < 1 Then
                strRet += ch
            End If
        Next
        Return strRet
    End Function


    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="inString"></param>
    ''' <param name="inChar"></param>
    ''' <param name="repStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Overloads Shared Function ReplaceLastCharWithString(ByVal inString As String, _
        ByVal inChar As Char, ByVal repStr As String) As String
        Dim retStr As String = ""
        If inString.Trim.Length = 0 OrElse InStr(inString, inChar) = 0 Then
            Return inString
        Else
            Dim i As Integer = inString.Length - 1
            Dim ResStr As String = ""

            While i > 0 And inString.Chars(i) <> inChar
                i -= 1
            End While

            If i > 0 Then
                ResStr = Mid(inString, 1, i) & repStr & Mid(inString, i + 2)
            End If
            Return ResStr
        End If
    End Function


    '''<summary>
    ''' Replaces a string elements in a string with another string
    ''' </summary>
    ''' <param name='inString'>Input String</param>
    ''' <param name='inStr1'>Input String</param>
    ''' <param name='repStr'>Input String</param>
    ''' <returns>Fixed String</returns>
    Public Overloads Shared Function ReplaceString(ByVal inString As String, _
        ByVal inStr1 As String, ByVal repStr As String) As String
        Dim retStr As String = ""
        Dim i As Integer
        Dim ResStr As String = ""

        i = InStr(inString, inStr1, CompareMethod.Text)
        If i = 0 Then
            ResStr = inString
        Else
            While i > 0
                ResStr += Mid(inString, 1, i - 1) & repStr
                inString = Mid(inString, i + inStr1.Length)
                i = InStr(inString, inStr1, CompareMethod.Text)
            End While
            ResStr += inString
        End If

        Return ResStr


    End Function
#End Region

#Region "Reporter"

    '''<summary>
    '''Send a fax via Email 
    ''' </summary>
    ''' <param name='Doc'>File containing fax</param>
    ''' <param name='Number'>Fax number</param>
    ''' <param name='Subject'>Subject of the fax</param>
    ''' <param name='Body'>Body of the fax</param>
    ''' <param name='From'>Sender</param>
    ''' <returns>Void</returns>
    Public Shared Function QuickFax(ByVal Doc As String, ByVal Number As String, _
        ByVal Subject As String, _
        ByVal Body As String, Optional ByVal From As String = "info.medscreen.com") As Boolean
        Dim objMail As Quiksoft.FreeSMTP.EmailMessage

        Try
            objMail = New Quiksoft.FreeSMTP.EmailMessage()          'Create a new email
            objMail.Recipients.Add(Number & Constants.GCST_EMAIL_To_Fax_Send)       'Set the recipient
            objMail.From.Email = From                               'Set the originator    
            objMail.Subject = Subject                               'Set the subject of the fax

            Dim W As IO.StreamReader
            W = New IO.StreamReader(Doc)

            Body = W.ReadToEnd()                                    'Get the body of the Fax
            W.Close()

            objMail.BodyParts.Add(Body, Quiksoft.FreeSMTP.BodyPartFormat.HTML)      'Fill Fax
            'objMail.Attachments.Add(Doc)
            Dim objSMTP As Quiksoft.FreeSMTP.SMTP
            'Create SMTP object
            Dim strEmailIp As String = MedscreenlibConfig.Servers("EMAILServer")
            objSMTP = New Quiksoft.FreeSMTP.SMTP(strEmailIp)
            objSMTP.Send(objMail)                                                   'And send it 
        Catch ex As Exception
            LogError(ex, True, " Sending Message")
        End Try


    End Function

    '''<summary>
    '''Send a message using the reporter program, default the output method to FAX.
    '''Takes an array of parameters of arbitary length to pass the information
    '''</summary>
    ''' <param name='OutAddress'>Output Address</param>
    ''' <param name='Template'>Template to use</param>
    ''' <param name='Subject'>Subject of the message</param>
    ''' <param name='Recipient'>Recipient(s) of message</param>
    ''' <param name='Parameters'>Parameters Array of </param>
    ''' <param name='OutputMethod'>Method for reporting default is FAX</param>
    ''' <param name='OutputDirectory'>Directory to use default is WordOutput</param>
    ''' <param name='OutputDesc'>Prefix of file name</param>
    ''' <param name='NoDups'>Prevent duplicate messages default is true</param>
    ''' <param name='ParamCount'>No of parameters</param>
    ''' <returns>Void</returns>
    Public Shared Function QuickReporter(ByVal OutAddress As String, _
    ByVal Template As String, _
    ByVal Subject As String, ByVal Recipient As String, _
    ByVal Parameters As Array, Optional ByVal OutputMethod As String = "FAX", _
    Optional ByVal OutputDirectory As String = "", _
    Optional ByVal OutputDesc As String = "QRP", _
    Optional ByVal NoDups As Boolean = False, _
    Optional ByVal ParamCount As Integer = 0) As Boolean

        'If no destination then die
        If OutAddress.Trim.Length = 0 Then Exit Function

        'Create output file name 
        If OutputDirectory.Trim.Length = 0 Then _
                OutputDirectory = LiveRoot() & "wordoutput\"
        If OutputDirectory.Chars(OutputDirectory.Length - 1) <> "\" Then
            OutputDirectory += "\"
        End If
        Dim f As String
        Dim blnRet As Boolean = True

        If Template Is Nothing OrElse Template.Trim.Length = 0 Then
            blnRet = False
            Exit Function
        End If
        If OutputDesc = "QRP" Then
            f = OutputDirectory & OutputDesc & Now.ToString("yyyyMMddHHmmss") & ".out"
        Else
            Dim int1 As Integer = 1
            Dim fStub As String = OutputDirectory & OutputDesc & "-" & CStr(int1).Trim
            f = fStub & ".out"
            If NoDups And (IO.File.Exists(fStub & ".out") Or IO.File.Exists(fStub & ".sent") Or IO.File.Exists(fStub & ".All_Done")) Then  'Not allowing dups so exit    
                Return False
                Exit Function
            End If
            While IO.File.Exists(f)     ' Loop until unique number found 
                int1 += 1
                f = OutputDirectory & OutputDesc & "-" & CStr(int1).Trim & ".out"
            End While
        End If

        Dim w As New IO.StreamWriter(f)
        w.WriteLine("'created by Customer Centre tool user ")
        w.WriteLine("'version 1.01 ")
        w.WriteLine("OUTPUT_TYPE=" & OutputMethod)              'Select Output method defaulting to FAX
        w.WriteLine("OUTPUT_ADDRESS=" & OutAddress)             'Recipients address i.e email or fax number
        'Set the email from address in BLAT
        w.WriteLine(MedscreenLib.Constants.ReporterEmailFrom)


        If Recipient.Trim.Length > 0 Then
            w.WriteLine("OUTPUT_TO=" & Recipient)                   'Recipient
        End If
        w.WriteLine("REPORT_TEMPLATE=" & Template)              'Template that reporter will use.
        'w.WriteLine("SUBJECT=" & Subject)                       'Subject of message

        Dim i As Integer
        Dim strBookmark As String
        Dim strValue As String
        'Add all the parameters 'Bookmark is in array[0,n] value in array[1,n]
        Dim UpprC As Integer
        If ParamCount = 0 Then
            UpprC = Parameters.GetUpperBound(1)
        Else
            UpprC = ParamCount
        End If
        For i = 0 To UpprC
            strBookmark = CStr(Parameters.GetValue(0, i))
            If Not strBookmark Is Nothing Then
                If strBookmark.Trim.Length > 0 Then
                    strValue = CStr(Parameters.GetValue(1, i))
                    w.WriteLine(strBookmark & "=" & strValue)
                End If
            End If
        Next
        w.Flush()               'Clear output stream
        w.Close()               'Close it 
        w = Nothing             'Destroy it 
        Return blnRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Private variable storing Blat path
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	20/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Shared strBlatPath As String = Constants.GCST_BLAT_DIRECTORY
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Path to BLAT email program
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	20/02/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property BlatPath() As String
        Get
            Return strBlatPath
        End Get
        Set(ByVal value As String)
            strBlatPath = value
        End Set
    End Property

    ''' <developer>CONCATENO\taylor</developer>
    ''' <summary>
    ''' Rename file
    ''' </summary>
    ''' <param name="Filename"></param>
    ''' <remarks>Some file already exists errors on rename, so adding seconds and millisecs to make it as unique as possible</remarks>
    ''' <revisionHistory>
    ''' <revision><created>28-Aug-2015 11:45</created><Author>CONCATENO\taylor</Author></revision>
    ''' </revisionHistory>
    Private Shared Sub MoveEmailProtected(ByVal Filename As String)
        Try
            Dim postfix As String = Now.ToString("ssfff")             'Millisecs to stop dupes
            Dim newName As String = ReplaceString(Filename, ".txt", postfix & ".done")
            IO.File.Move(Filename, newName)
        Catch ex As Exception
        End Try
    End Sub
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Send an email by BLAT
    ''' </summary>
    ''' <param name="subject"></param>
    ''' <param name="body"></param>
    ''' <param name="recipient"></param>
    ''' <param name="Originator"></param>
    ''' <param name="HTML"></param>
    ''' <param name="BCC"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	12/06/2007	Created
    ''' <revision><modified>28-Aug-2015 11:46</modified><Author>CONCATENO\taylor</Author><REV>Renaming moved out and made protected</REV></revision>
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function BlatEmail(ByVal subject As String, ByVal body As String, _
            Optional ByVal recipient As String = "doug.taylor@concateno.com", _
        Optional ByVal Originator As String = "toxeu.LabReports@alere.com", _
        Optional ByVal HTML As Boolean = True, Optional ByVal BCC As String = "", Optional ByVal attachment As String = "", _
        Optional ByVal SendSecure As Boolean = False) As Boolean

        Try
            Dim Filename As String = GetFileName("Blat", Now, "txt")
            Dim iof As IO.StreamWriter
            Try  ' Protecting
                iof = New IO.StreamWriter(Filename)
                iof.WriteLine(body)
                iof.Flush()
                iof.Close()
                'If Not CConnection.DBInstance Is Nothing AndAlso InStr(CConnection.DBInstance, "LIVE") = 0 Then
                '    recipient = "Doug.Taylor@concateno.com"
                'End If
                'See If we have any groups
                Dim iPos As Integer = InStr(recipient, "[")
                While iPos > 0
                    Dim groupName As String
                    Dim iPosEnd As Integer = InStr(recipient, "]")
                    If iPosEnd > 0 Then
                        groupName = Mid(recipient, iPos + 1, iPosEnd - iPos - 1)
                        Dim oColl As Collection = ActiveDirectoryWrapper.PC.ADWrapper.GetGroupMembersEmail(groupName)
                        'remove group from name
                        recipient = Mid(recipient, 1, iPos - 1) & Mid(recipient, iPosEnd + 1).Trim
                        recipient = Medscreen.ReplaceString(recipient, ",,", ",")
                        'add group members
                        For Each member As String In oColl
                            If recipient.Trim.Length > 1 Then recipient += ","
                            recipient += member
                        Next

                    End If
                    iPos = InStr(recipient, "[")
                End While

                ' If we are sending secure add indicator to subject line
                If SendSecure Then
                    subject = "[SecureMail] " & subject
                End If

                'ensure we use commas 
                recipient = recipient.Replace(";", ",")
                Dim strServer As String = MedscreenlibConfig.BLAT("SERVER")
                Dim strBlat As String = Filename & " -to " & recipient & " -f " & Originator & " -subject " & Chr(34) & subject & Chr(34)
                If strServer.Trim.Length > 0 Then
                    strBlat += " -server " & strServer
                End If
                If HTML Then
                    strBlat += " -HTML"
                End If
                'do bcc
                If BCC.Trim.Length > 0 Then
                    strBlat += " -bcc " & BCC
                End If
                If attachment.Trim.Length > 0 Then
                    strBlat += " -attach " & attachment
                End If
                strBlat += " " & MedscreenlibConfig.BLAT("BLATCredentials")

                'Create process rather than shelling to gain more control
                Dim myProcess As Process = New Process()
                'set program to use and arguments and window style then start process

                myProcess.StartInfo.FileName = """\\corp.concateno.com\medscreen\common\lab programs\blat\Blat.exe"""
                myProcess.StartInfo.Arguments = strBlat
                myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                myProcess.StartInfo.CreateNoWindow = True
                myProcess.StartInfo.RedirectStandardOutput = True
                myProcess.StartInfo.UseShellExecute = False                 'Required for redirecting
                MedscreenLib.Medscreen.LogAction("starting process - " & strBlat)

                myProcess.Start()
                Dim iso As IO.StreamReader = myProcess.StandardOutput

                'Wait 5 secs for the process to exit naturally if not kill the process
                myProcess.WaitForExit(20000)
                If Not myProcess.HasExited Then
                    myProcess.Kill()
                End If
                Dim StrLog As String = iso.ReadToEnd
                MedscreenLib.Medscreen.LogAction(myProcess.ExitTime & "." & "  Exit Code: " & myProcess.ExitCode & " - " & StrLog)
                iso.Close()


                'record the result of the call so errors can be found
                Debug.WriteLine(myProcess.ExitTime & "." & "  Exit Code: " & myProcess.ExitCode)
                'Dim intProc As Integer = Shell(strBlat, AppWinStyle.Hide, True, 2000)
                'Debug.WriteLine(intProc)
                myProcess.Close()
                LogAction("Email sent to " & recipient & " : " & subject, False)
            Catch ex As Exception
                LogAction(ex.ToString)

                MedscreenLib.Medscreen.LogError(ex, , "Medscreen-BlatEmail-1775")
            Finally
                While FileInUse(Filename)
                    Threading.Thread.Sleep(1000)
                End While

                MoveEmailProtected(Filename)                            'Added 27/08/2015

            End Try

        Catch ex As Exception
        End Try
    End Function

    Public Shared Function TOPDF(ByVal strfileName As String, ByVal strNewName As String) As Boolean
        Dim blnReturn As Boolean = True
        Try
            Dim strPath As String = " " & strfileName & " " & strNewName
            'Shell(strpath)

            Dim myProcess As Process = New Process()
            'set program to use and arguments and window style then start process

            Dim ScriptPath As String = MedscreenLib.MedscreenlibConfig.Servers("GhostScript")

            myProcess.StartInfo.FileName = Chr(34) & ScriptPath & Chr(34)
            myProcess.StartInfo.Arguments = strPath
            myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.UseShellExecute = False                 'Required for redirecting
            MedscreenLib.Medscreen.LogAction("starting process - " & strPath)

            myProcess.Start()
            Dim iso As IO.StreamReader = myProcess.StandardOutput

            'Wait 5 secs for the process to exit naturally if not kill the process
            myProcess.WaitForExit(5000)
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            Dim StrLog As String = iso.ReadToEnd
            MedscreenLib.Medscreen.LogAction(myProcess.ExitTime & "." & "  Exit Code: " & myProcess.ExitCode & " - " & StrLog)
            iso.Close()
        Catch ex As Exception
            blnReturn = False
        End Try
        Return blnReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check to see whether a file is in use or not
    ''' </summary>
    ''' <param name="sFile"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	05/03/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function FileInUse(ByVal sFile As String) As Boolean
        If System.IO.File.Exists(sFile) Then
            Try
                Dim F As Short = FreeFile()
                FileOpen(F, sFile, OpenMode.Binary, OpenAccess.ReadWrite, OpenShare.LockReadWrite)
                FileClose(F)
            Catch
                Return True
            End Try
        End If
    End Function
    '''<summary>
    '''Send an email out 
    '''</summary>
    ''' <param name='subject'>Subject of the message</param>
    ''' <param name='body'>body of message</param>
    ''' <param name='recipient'>Recipient(s) of message</param>
    ''' <param name='Originator'>Sender) of message</param>
    ''' <param name='HTML'>Send message as HTML</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function QuckEmail(ByVal subject As String, ByVal body As String, _
        Optional ByVal recipient As String = "doug.taylor@concateno.com", _
    Optional ByVal Originator As String = "toxeu.LabReports@alere.com", _
    Optional ByVal HTML As Boolean = True, Optional ByVal BCC As String = "") As Boolean
        Dim blnRet As Boolean = BlatEmail(subject, body, recipient, Originator, HTML, BCC)
        'Dim objMail As Quiksoft.FreeSMTP.EmailMessage
        'Dim objSMTP As Quiksoft.FreeSMTP.SMTP
        'Dim Recipients As String()
        'Dim i As Integer

        'Dim blnRet As Boolean = True

        'If recipient.Trim.Length = 0 Then Exit Function 'No valid recipient so don't send it

        'Try
        '    'Create Email
        '    objMail = New Quiksoft.FreeSMTP.EmailMessage()
        '    Recipients = recipient.Split(New Char() {","c, ";"c}, 100)
        '    'Get recipients from list 
        '    For i = 0 To Recipients.Length - 1
        '        If Recipients(i).Trim.Length > 0 Then
        '            objMail.Recipients.Add(Recipients(i))
        '        End If
        '    Next

        '    If BCC.Trim.Length > 0 Then
        '        Dim Customheader As Quiksoft.FreeSMTP.CustomHeader = New Quiksoft.FreeSMTP.CustomHeader()
        '        Customheader.Name = "Bcc"
        '        Customheader.Value = BCC
        '        objMail.CustomHeaders.Add(Customheader)
        '    End If

        '    'objMail.CustomHeaders.Add(New Quiksoft.FreeSMTP.CustomHeader("BCC", strBCC))
        '    'Set the originator to Info
        '    objMail.From.Email = Originator
        '    'Add the supplied subject 
        '    objMail.Subject = subject

        '    'Add the body of the message is in HTML format
        '    If HTML Then
        '        objMail.BodyParts.Add(body, Quiksoft.FreeSMTP.BodyPartFormat.HTML)
        '    Else
        '        objMail.BodyParts.Add(body, Quiksoft.FreeSMTP.BodyPartFormat.Plain)
        '    End If
        '    'Create SMTP envelope and send it to Server
        '    Dim strEmailIp As String = MedscreenlibConfig.Servers("EMAILServer")
        '    objSMTP = New Quiksoft.FreeSMTP.SMTP(strEmailIp)
        '    objSMTP.Send(objMail)
        'Catch ex As Exception
        '    LogError(ex, True)
        '    blnRet = False
        'End Try
        Return blnRet

    End Function

    '''<summary>
    '''Send an email out 
    '''</summary>
    ''' <param name='subject'>Subject of the message</param>
    ''' <param name='body'>body of message</param>
    ''' <param name='Attachments'>Array of file attachments</param>
    ''' <param name='recipient'>Recipient(s) of message</param>
    ''' <param name='Originator'>Sender) of message</param>
    ''' <param name='HTML'>Send message as HTML</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function QuckEmail(ByVal subject As String, ByVal body As String, _
    ByVal Attachments As String(), Optional ByVal recipient As String = "doug.taylor@concateno.com", _
Optional ByVal Originator As String = "toxeu.LabReports@alere.com", _
Optional ByVal HTML As Boolean = True, Optional ByVal BCC As String = "") As Boolean
        Dim i As Integer
        Dim blnRet As Boolean = True
        For i = 0 To Attachments.Length - 1
            Dim strAttach As String = Attachments(i)
            If strAttach.Trim.Length > 0 And IO.File.Exists(strAttach) Then
                Medscreen.BlatEmail(subject, body, recipient, Originator, HTML, BCC, strAttach)
            End If
        Next

        'Dim objMail As Quiksoft.FreeSMTP.EmailMessage
        'Dim objSMTP As Quiksoft.FreeSMTP.SMTP
        'Dim Recipients As String()
        'Dim i As Integer

        'Dim blnRet As Boolean = True

        'If recipient.Trim.Length = 0 Then Exit Function 'No valid recipient so don't send it

        'Try
        '    'Create Email
        '    objMail = New Quiksoft.FreeSMTP.EmailMessage()
        '    Recipients = recipient.Split(New Char() {","c, ";"c}, 100)
        '    'Get recipients from list 
        '    For i = 0 To Recipients.Length - 1
        '        If Recipients(i).Trim.Length > 0 Then
        '            objMail.Recipients.Add(Recipients(i))
        '        End If
        '    Next
        '    If BCC.Trim.Length > 0 Then
        '        Dim Customheader As Quiksoft.FreeSMTP.CustomHeader = New Quiksoft.FreeSMTP.CustomHeader()
        '        Customheader.Name = "BCC"
        '        Customheader.Value = BCC
        '        objMail.CustomHeaders.Add(Customheader)
        '    End If
        '    For i = 0 To Attachments.Length - 1
        '        Dim strAttach As String = Attachments(i)
        '        If strAttach.Trim.Length > 0 Then objMail.Attachments.Add(New Quiksoft.FreeSMTP.Attachment(strAttach))
        '    Next
        '    'Set the originator to Info
        '    objMail.From.Email = Originator
        '    'Add the supplied subject 
        '    objMail.Subject = subject

        '    'Add the body of the message is in HTML format
        '    If HTML Then
        '        objMail.BodyParts.Add(body, Quiksoft.FreeSMTP.BodyPartFormat.HTML)
        '    Else
        '        objMail.BodyParts.Add(body, Quiksoft.FreeSMTP.BodyPartFormat.Plain)
        '    End If
        '    'Create SMTP envelope and send it to Server
        '    Dim strEmailIp As String = MedscreenlibConfig.Servers("EMAILServer")

        '    objSMTP = New Quiksoft.FreeSMTP.SMTP(strEmailIp)
        '    objSMTP.Send(objMail)
        'Catch ex As Exception
        '    LogError(ex, True)
        '    blnRet = False
        'End Try
        Return blnRet

    End Function



#End Region

#Region "Database"

    '''<summary>
    ''' Get a recordset based on supplied query, connection left open as is reader
    ''' </summary>
    ''' <param name='Query'>Query used to retrieve dataset</param>
    ''' <param name='Conn'>Connection used</param>
    ''' <returns>populated data reader</returns>
    Public Shared Function GetRecordSet(ByVal Query As String, ByVal Conn As OleDb.OleDbConnection, _
    Optional ByVal Parameter As OleDb.OleDbParameter = Nothing) As OleDb.OleDbDataReader
        Dim oCmd As New OleDb.OleDbCommand()
        oCmd.Connection = Conn
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        oCmd.CommandText = Query
        If Not Parameter Is Nothing Then
            oCmd.Parameters.Add(Parameter)
        End If

        Dim oRead As OleDb.OleDbDataReader

        Try
            oRead = oCmd.ExecuteReader
            Return oRead
        Catch ex As Exception
            Return Nothing
            LogError(ex, , Query)
        End Try
    End Function

    '''<summary>
    ''' Convert an XML into a tablset
    ''' </summary>
    ''' <param name='XML'>Source of XML to convert</param>
    ''' <returns>A generic dataset</returns>
    Public Shared Function ToDataset(ByVal XML As String) As DataSet
        Try
            Dim datadoc As New System.Xml.XmlDataDocument()

            Dim st As New IO.MemoryStream(Medscreen.StringToByteArray(XML))

            Dim xr As XmlTextReader = New XmlTextReader(st)

            datadoc.DataSet.ReadXml(xr, XmlReadMode.InferSchema)
            st = Nothing
            xr = Nothing
            Return datadoc.DataSet
        Catch ex As Exception
            Return Nothing
        Finally

        End Try

    End Function


    '''<summary>
    ''' Return the name of the table set given a date
    ''' </summary>
    ''' <param name='inDate'>Date to produce table set for </param>
    ''' <returns>Formatted table set name</returns>
    Public Shared Function TableSet(ByVal inDate As Date) As String
        Dim strret As String
        If inDate.Year > 2006 Then
            Return inDate.ToString("yyyy")
        Else
            If inDate.Month >= 9 Then
                strret = "3"
            ElseIf inDate.Month >= 5 Then
                strret = "2"
            Else
                strret = "1"
            End If
            Return strret & Mid(inDate.ToString("yyyy"), 2, 3)
        End If
    End Function

#End Region

#Region "Calculators"
    '''<summary>
    ''' Calculate the length of time that is out of hours
    ''' </summary>
    ''' <param name='Dat1'>Starting date time variable</param>
    ''' <param name='dat2'>Ending date time variable</param>
    ''' <returns>Time out of hours as Timespan</returns>
    Public Shared Function HoursOutOfHours(ByVal Dat1 As Date, _
      ByVal dat2 As Date, Optional ByVal Country As String = "UK", Optional ByVal StartTime As Integer = 8, Optional ByVal endtime As Integer = 18) As TimeSpan

        Dim Elapse As TimeSpan
        Dim ts1 As TimeSpan
        Dim ts2 As TimeSpan
        'Dim datTemp As Date
        Dim bln1 As Boolean
        Dim bln2 As Boolean
        If Country = "GB" Then Country = "UK"
        Elapse = TimeSpan.Zero
        If Date.Compare(Dat1, DateSerial(1970, 1, 2)) < 0 Then Return TimeSpan.Zero 'Check to see if date1 is old
        If Date.Compare(dat2, DateSerial(1970, 1, 2)) < 0 Then Return TimeSpan.Zero 'Check to see if date2 is old
        If Date.Compare(dat2, Dat1) = 0 Then Return TimeSpan.Zero 'Check to see if dates are the same

        bln1 = IsOutOfHours(Dat1, Country)
        bln2 = IsOutOfHours(dat2, Country)
        Console.WriteLine(bln1 & "-" & bln2)

        If bln1 And bln2 Then
            If (Dat1.DayOfWeek = DayOfWeek.Saturday) OrElse (Dat1.DayOfWeek = DayOfWeek.Sunday) OrElse IsCountryHoliday(Dat1, Country) Then
                Return dat2.Subtract(Dat1)
            ElseIf Dat1.Day = dat2.Day AndAlso ((Dat1.Hour < StartTime And dat2.Hour < StartTime) OrElse (Dat1.Hour >= 18 And dat2.Hour >= 18)) Then
                Return dat2.Subtract(Dat1)
            ElseIf dat2.Day = Dat1.AddDays(1).Day AndAlso (Dat1.Hour >= endtime And dat2.Hour < StartTime) Then
                Return dat2.Subtract(Dat1)
            ElseIf Dat1.Hour = 0 AndAlso Dat1.Minute = 0 AndAlso Dat1.Second = 0 AndAlso dat2.Hour > endtime Then
                Return Dat1.Subtract(dat2)
            Else
                ts1 = FirstTime(Dat1)
                ts2 = SecondTime(Dat1, dat2)
                Return ts1.Add(ts2)
            End If
            Exit Function
        End If

        If (Not bln1) And (Not bln2) Then
            If (Dat1.DayOfWeek = DayOfWeek.Saturday) Or (Dat1.DayOfWeek = DayOfWeek.Sunday) Then
                Return dat2.Subtract(Dat1)
            Else
                Return TimeSpan.Zero
            End If
            Exit Function
        End If

        'only one is true 
        If bln1 Then 'only start time is out of hours
            If (Dat1.DayOfWeek = DayOfWeek.Saturday) Or (Dat1.DayOfWeek = DayOfWeek.Sunday) Then
                Return dat2.Subtract(Dat1)
            Else
                Return FirstTime(Dat1)
            End If
        End If

        If bln2 Then 'Normal start but work late 
            If (Dat1.DayOfWeek = DayOfWeek.Saturday) Or (Dat1.DayOfWeek = DayOfWeek.Sunday) Then
                Console.WriteLine(Dat1.DayOfWeek)
                Return dat2.Subtract(Dat1)
            Else
                Return SecondTime(Dat1, dat2)
            End If
        End If


        '
        'Elapse = Elapse.Add(datTemp.Subtract(Dat1))


    End Function

#End Region

#Region "Hyperlinks"

    '''<summary>
    ''' Provide a hyperlink to a client
    ''' </summary>
    ''' <param name='Dr'>Customer ID</param>
    ''' <returns>Hyperlink to intranet customer report</returns>
    Public Shared Function HyperClient(ByVal Dr As String, ByVal SMIDPRofile As String) As String
        Dim strStyleSheet As String = MedscreenlibConfig.HyperLinks.Item("CustStyleSheet")
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Client")
        HyperClient = "<a href=""" & strHyper & _
            Dr & "&stylesheet=" & strStyleSheet & """>" & SMIDPRofile & "</a>"
    End Function


    Public Shared Function URLClient(ByVal Dr As String, ByVal SMIDPRofile As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Client")
        Dim strStyleSheet As String = MedscreenlibConfig.HyperLinks.Item("CustStyleSheet")
        Return strHyper & Dr & "&stylesheet=" & strStyleSheet & "&SMID=" & SMIDPRofile
    End Function

    Public Shared Function URLVessel(ByVal VesselId As String, ByVal NAME As String) As String
        Return "http://TS01/WebGCMSLate/WebFormPositives.aspx?Action=VESSEL&ID=" & VesselId & "&name=" & NAME
    End Function

    '''<summary>
    ''' Provide a hyperlink to a sample manager job
    ''' </summary>
    ''' <param name='strJob'>Job name</param>
    ''' <returns>Hyperlink to intranet job report</returns>
    Public Shared Function HyperJob(ByVal strJob As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("job")
        Dim strStyleSheet As String = MedscreenlibConfig.HyperLinks.Item("JobStyleSheet")
        Return "<a href=""" & strHyper & _
        strJob & "&stylesheet=" & strStyleSheet & """>" & strJob & "</a>"
    End Function


    '''<summary>
    ''' Provide a URL to Job webpage
    ''' </summary>
    ''' <param name='strJob'>Job Name</param>
    ''' <returns>URL to intranet job web page</returns>
    Public Shared Function URLJob(ByVal strJob As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("job")
        Return strHyper & _
        strJob
    End Function

    Public Shared Function URLCollections(ByVal CustomerID As String, ByVal StartDate As Date, Optional ByVal EndDate As String = "") As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("jobs")
        strHyper = strHyper & CustomerID & ";&StartDate=" & StartDate.ToString("dd-MMM-yyyy")
        If EndDate.Trim.Length = 0 Then
            strHyper += ";&EndDate=" & Today.ToString("dd-MMM-yyyy")
        Else
            strHyper += ";&EndDate=" & EndDate
        End If
        Return strHyper

    End Function

    '''<summary>
    ''' Provide a hyperlink to a collection
    ''' </summary>
    ''' <param name='strJob'>Collection ID</param>
    ''' <returns>Hyperlink to intranet collection page</returns>
    Public Shared Function HyperCollection(ByVal strJob As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Collection")
        Dim strStyleSheet As String = MedscreenlibConfig.HyperLinks.Item("CollStyleSheet")
        Return "<a href=""" & strHyper & _
        strJob & "&Stylesheet=" & strStyleSheet & """>" & strJob & "</a>"
    End Function

    '''<summary>
    ''' Provide a hyperlink to a sample
    ''' </summary>
    ''' <param name='Barcode'>Sample Barcode</param>
    ''' <returns>Hyperlink to intranet sample web page</returns>
    Public Shared Function HyperBarcode(ByVal Barcode As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Sample")
        Return "<a href=""" & strHyper & Barcode & "&barcode=TRUE"">" & Barcode & "</a>"
    End Function

    '''<summary>
    ''' Provide a URL to a Sample
    ''' </summary>
    ''' <param name='Barcode'>Sample Barcode</param>
    ''' <returns>URL to intranet sample webpage</returns>
    Public Shared Function URLBarcode(ByVal Barcode As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Sample")
        Return strHyper & Barcode & "&barcode=TRUE"
    End Function


    '''<summary>
    ''' Provide a hyperlink to an invoice
    ''' </summary>
    ''' <param name='strInvoice'>Invoice Number</param>
    ''' <returns>Hyperlink to intranet invoice web page</returns>
    Public Shared Function HyperInvoice(ByVal strInvoice As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Invoice")
        If strInvoice = "FOC" Then
            Return "Not Invoiced"
        Else
            Return "<a href=""" & strHyper & _
            strInvoice & """>" & strInvoice & "</a>"
        End If
    End Function

    '''<summary>
    ''' Provide a URL to a Invoice
    ''' </summary>
    ''' <param name='strInvoice'>Invoice Number</param>
    ''' <returns>URL to intranet Invoice web page</returns>
    Public Shared Function URLInvoice(ByVal strInvoice As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Invoice")

        Return strHyper & _
            strInvoice

    End Function

    '''<summary>
    ''' Provide a URL to a barcode
    ''' </summary>
    ''' <param name='strSampleId'>Barcode</param>
    ''' <returns>URL to intranet sample report</returns>
    Public Shared Function URLSample(ByVal strSampleId As String) As String
        Dim strHyper As String = MedscreenlibConfig.HyperLinks.Item("Sample")
        Return strHyper & strSampleId & "&barcode=TRUE"
    End Function

#End Region

#Region "User / Role Functions"
    Public Shared Function UserInRole(ByVal PersonnelID As String, ByVal RoleID As String) As Boolean
        Dim oColl As New Collection()
        oColl.Add(RoleID)
        oColl.Add(PersonnelID)
        Dim strPos As String = CConnection.PackageStringList("Lib_utils.PositionInRole", oColl)
        Return (InStr(strPos, "-1") = 0)
    End Function

    ''' <developer></developer>
    ''' <summary>
    ''' Get a list of email address for role members
    ''' </summary>
    ''' <param name="Role"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Shared Function RoleEmailList(ByVal Role As String) As String
        Dim recipString As String = CConnection.PackageStringList("lib_utils.emailrole", Role)
        If Not recipString Is Nothing Then
            Return recipString
        Else
            Return ""
        End If
    End Function
#End Region

#Region "Check Functions"

    Public Enum TimeUnit
        hr
        min
    End Enum
    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="StartDate"></param>
    ''' <param name="Days"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Shared Function AddWorkingDays(ByVal StartDate As Date, ByVal Days As Integer) As Date
        Dim RetDate As Date = StartDate
        While Days > 0 OrElse Not IsWorkingDay(RetDate)
            If Not IsWorkingDay(RetDate) Then
                RetDate = RetDate.AddDays(1)
            Else

                Days -= 1
                RetDate = RetDate.AddDays(1)
            End If

        End While
        Return RetDate
    End Function

    Public Shared Function IsWorkingDay(ByVal StartDate As Date) As Boolean
        Dim RetDate As Date = StartDate
        Return Not (RetDate.DayOfWeek = DayOfWeek.Saturday OrElse RetDate.DayOfWeek = DayOfWeek.Sunday OrElse IsAHoliday(RetDate))
    End Function

    '''<summary>
    ''' See if date is a holiday
    ''' </summary>
    ''' <param name='inDate'> Date to check </param>
    ''' <returns> ~True if a Holiday </returns>
    Public Shared Function IsAHoliday(ByVal inDate As Date) As Boolean
        'Dim Holidays() As Date = {#12/25/2003#, #12/26/2003#, #1/1/2004#, #4/9/2004#, #4/12/2004#, _
        '    #5/3/2004#, #5/31/2004#, #8/30/2004#, #12/27/2004#, #12/28/2004#, #1/2/2004#}
        'Dim i As Integer
        'Dim blnRet As Boolean = False
        'inDate = DateSerial(inDate.Year, inDate.Month, inDate.Day)
        'For i = 0 To Holidays.Length - 1
        '    If inDate = CType(Holidays(i), Date) Then
        '        blnRet = True
        '        Exit For
        '    End If
        'Next
        Return Calendar.IsAHoliday(inDate)

    End Function

    '''<summary>
    '''Check to see if this client is a London Underground one 
    ''' </summary>
    ''' <param name='Customer'>Customer Id to check</param>
    ''' <returns>True if Customer is London Underground</returns>
    Public Shared Function IsLondonUnderground(ByVal Customer As String) As Boolean
        Dim STLUL As String = ",LULUN,LULFC,LULPE,LULNR,LULMO,LULPI,LULDAMSP,LULBARCODE,LULBR,"

        If Customer.Trim.Length = 0 Then
            Return False
            Exit Function
        End If
        Return (InStr(STLUL, "," & Customer.Trim.ToUpper & ",") <> 0)

    End Function

    Public Shared Function IsCountryHoliday(ByVal datin As Date, Optional ByVal Country As String = "UK") As Boolean
        Dim strTrue As String = "F"
        If Country = "GB" Then Country = "UK"
        Dim oColl As New Collection()
        oColl.Add(datin.ToString("dd-MMM-yy"))
        oColl.Add(Country)
        strTrue = CConnection.PackageStringList("IsCountryHoliday", oColl) 'See if it is a public holiday in the country GB is the default
        If strTrue = "T" Then
            Return True
        Else
            Return False
        End If
    End Function

    '''<summary>
    '''Normal outof hours check 
    ''' </summary>
    ''' <param name='datIn'> Date to Check</param>
    ''' <returns>True if is out of hours</returns>
    Public Shared Function IsOutOfHours(ByVal datIn As Date, Optional ByVal Country As String = "UK") As Boolean
        Dim strTrue As String = "F"
        Dim oColl As New Collection()
        If Country = "GB" Then Country = "UK"
        oColl.Add(datIn.ToString("dd-MMM-yy"))
        oColl.Add(Country)
        strTrue = CConnection.PackageStringList("IsCountryHoliday", oColl) 'See if it is a public holiday in the country GB is the default

        If strTrue = "T" Then
            Return True
        ElseIf datIn.DayOfWeek > DayOfWeek.Friday Then
            Return True

        ElseIf datIn.Hour <= 7 Or (datIn.Hour * 100 + datIn.Minute) > 1800 Then
            Return True
        Else
            Return False
        End If
    End Function

    '''<summary>
    '''Is the string a valid number
    ''' </summary>
    ''' <param name='Source'>String to check</param>
    ''' <returns>True if a valid number</returns>
    Public Shared Function IsNumber(ByVal Source As String) As Boolean
        Dim blnRet As Boolean = True
        Dim i As Integer = 0
        'Dim ch As Char

        Source = Source.Trim
        While i < Source.Length AndAlso blnRet
            blnRet = (Char.IsDigit(Source, i) Or Char.IsWhiteSpace(Source, i))
            i += 1
        End While
        Return blnRet
    End Function

    '''<summary>
    '''Check to see if the time and date specified is out of hours for an LUL collection.
    '''In hours is Monday - Friday 0700 - 1800 hours
    ''' </summary>
    ''' <param name='inDate'> Date to Check</param>
    ''' <returns>True if is out of hours</returns>
    Public Shared Function LULIsOutOfHours(ByVal inDate As Date) As Boolean
        Dim blnReturn As Boolean = True

        'Deal with holidays
        If IsAHoliday(inDate) Then
            Return True
            Exit Function
        End If

        'Deal with weekends 
        If inDate.DayOfWeek = DayOfWeek.Saturday Or inDate.DayOfWeek = DayOfWeek.Sunday Then
            Return blnReturn
            Exit Function
        End If

        'Deal with day time 
        If inDate.Hour < 7 Or inDate.Hour > 18 Then
            Return blnReturn
            Exit Function
        End If
        Return False
    End Function

    ''' <summary>
    ''' Validate a set of emails
    ''' </summary>
    ''' <param name="EmailAddressSet"></param>
    ''' <returns>-1 success, n where n is the first email address in error. </returns>
    ''' <remarks></remarks>
    Public Shared Function isValidEmailSet(ByVal EmailAddressSet As String) As Integer
        Dim emailAddresses As String() = EmailAddressSet.Split(New Char() {",", ";"})
        Dim intReturn As Integer = -1
        Dim i As Integer
        For i = 0 To emailAddresses.Length - 1
            Dim CurrentAddress As String = emailAddresses.GetValue(i)
            If Not IsValidEmail(CurrentAddress) OrElse CurrentAddress.Trim.Length = 0 Then
                intReturn = i
                Exit For
            End If
        Next
        Return intReturn

    End Function

    ''' <summary>
    ''' Function to validate an email address
    ''' </summary>
    ''' <param name="EmailAddress"></param>
    ''' <returns></returns>
    ''' <remarks>Uses a regular expression</remarks>
    Public Shared Function IsValidEmail(ByVal EmailAddress As String) As Boolean
        Dim blnRet As Boolean = False
        Dim strRegEx As String = MedscreenlibConfig.Validation.Item("EmailMask") ' ^([a-zA-Z0-9_&\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$ 'Email address validator
        Dim Regex As New System.Text.RegularExpressions.Regex(strRegEx)
        Dim myMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(EmailAddress.Trim, strRegEx)
        blnRet = myMatch.Success
        Return blnRet
    End Function
#End Region

#Region "Logging"

    Public Shared Event MessageLogged(ByVal e As MessageEventArgs)

    Public Enum ErrVisibility
        [Default]
        ShowAll
        HideAll
    End Enum

    Private Shared _errorVisibility As ErrVisibility
    Private Shared _errorLogFile As String = pvtErrorLoggingDirectory & "Error.txt"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Sets visibility to use when logging errors
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Boughton]</Author><date> [14/08/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Property ErrorVisibility() As ErrVisibility
        Get
            Return _errorVisibility
        End Get
        Set(ByVal Value As ErrVisibility)
            _errorVisibility = Value
        End Set
    End Property

    Public Shared Property ErrorLogFile() As String
        Get
            Return _errorLogFile
        End Get
        Set(ByVal Value As String)
            _errorLogFile = Value
        End Set
    End Property

    '''<summary>
    ''' Log a request
    ''' </summary>
    ''' <param name='ActionString'>Action to Log</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function LogRequest(ByVal ActionString As String, _
        Optional ByVal show As Boolean = False) As Boolean
        LogAction(ActionString, "EmailRequest.txt", show)
    End Function

    '''<summary>
    ''' Log a request
    ''' </summary>
    ''' <param name='ActionString'>Action to Log</param>
    ''' <param name='ActionFile'>The file to log to</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function LogRequest(ByVal ActionString As String, _
    ByVal ActionFile As String, Optional ByVal show As Boolean = False) As Boolean
        LogAction(ActionString, ActionFile, show)
    End Function

    Private Shared myActionPrefix As String = ""
    ''' <summary>
    ''' Set a variant for the action file, defaulting to nothing 
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>26-Feb-2015 08:12</created><Author>CONCATENO\taylor</Author></revision></revisionHistory>
    Public Shared Property ActionPreFix() As String
        Get
            Return myActionPrefix
        End Get
        Set(ByVal value As String)
            myActionPrefix = value
        End Set
    End Property

    Public Overloads Shared Function LogXML(ByVal Element As XElement, ByVal FileName As String) As Boolean
        Try
            Dim Log As New MedscreenLib.FileLogger(FileName)
            Log.WriteXML(Element)

        Catch ex As Exception

        End Try
    End Function

    '''<summary>
    ''' Log an action 
    ''' </summary>
    ''' <param name='ActionString'>Action to Log</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function LogAction(ByVal ActionString As String, _
      Optional ByVal show As Boolean = False) As Boolean
        LogAction(ActionString, ActionPreFix & Today.ToString("yyyy-MMM") & "-Action.txt", show)

    End Function

    '''<summary>
    ''' Log a request
    ''' </summary>
    ''' <param name='ActionString'>Action to Log</param>
    ''' <param name='ActionFile'>The file to log to</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function LogAction(ByVal ActionString As String, _
    ByVal ActionFile As String, Optional ByVal show As Boolean = False) As Boolean
        Dim e As New MessageEventArgs(ActionString, show)
        RaiseEvent MessageLogged(e)

        If Not e.Cancel Then
            OutputDebugString(ActionString) 'Send to remote debugger
            Dim Log As New MedscreenLib.FileLogger(ActionFile)
            Log.WriteLog(ActionString, Diagnostics.EventLogEntryType.Information)
            If e.Show Then MsgBox(ActionString, , "Information")
        End If
    End Function

    '''<summary>
    ''' Log an error
    ''' </summary>
    ''' <param name='ErrorString'>Action to Log</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    ''' <remarks>The show parameter can be overridden by the shared ErrorVisibility property</remarks>
    Public Overloads Shared Function LogError(ByVal ErrorString As String, ByVal show As Boolean) As Boolean
        show = ErrorVisibility <> ErrVisibility.HideAll AndAlso _
               (show Or (Medscreen.ErrorVisibility = ErrVisibility.ShowAll))
        Dim e As New MessageEventArgs(ErrorString, show, Diagnostics.EventLogEntryType.Error)
        RaiseEvent MessageLogged(e)
        If Not e.Cancel Then
            OutputDebugString(ErrorString) 'Send to remote debugger
            Dim Log As New MedscreenLib.FileLogger(ErrorLogFile)
            Log.WriteLog(ErrorString, Diagnostics.EventLogEntryType.Error)
            ' This line should become redundant once the ErrorVisibility property is set as necessary
            If Security.Principal.WindowsIdentity.GetCurrent.Name = "CONCATENO\FullAccess" Then Exit Function
            ' Check error visibility and report error to user if required
            If e.Show Then
                Windows.Forms.MessageBox.Show(ErrorString, "Error", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End If
        End If
    End Function

    '''<summary>
    ''' Log an error
    ''' </summary>
    ''' <param name='ErrorString'>Action to Log</param>
    ''' <returns>Void</returns>
    Public Overloads Shared Function LogError(ByVal ErrorString As String) As Boolean
        Return LogError(ErrorString, False)
    End Function
    ''' -----------------------------------------------------------------------------
    '''<summary>
    ''' Log an error
    ''' </summary>
    ''' <param name='Ex'>Exception to Log</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <param name='ExtraInfo'>Additional information to log</param>
    ''' <returns>Void</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/11/2006]</date><Action>Added support for machine name </Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function LogError(ByVal ex As Exception, _
    Optional ByVal show As Boolean = False, Optional ByVal ExtraInfo As String = " ") As Boolean
        'Dim Log As New MedscreenLib.FileLogger(pvtErrorLoggingDirectory & "Error.txt")
        Dim ErrorString As String = ex.ToString

        If Not ex.TargetSite Is Nothing Then
            ErrorString &= "-" & ex.TargetSite.Name
        End If

        ErrorString &= " " & ExtraInfo & " User : " & Security.Principal.WindowsIdentity.GetCurrent.Name & _
            " Machine : " & System.Environment.MachineName & _
            " OS : " & System.Environment.OSVersion.ToString

        'Log.WriteLog(ErrorString, Diagnostics.EventLogEntryType.Error)
        LogError(ErrorString, show)
    End Function


    Private Shared myErrorRecipient As String = "Simon.Walker@alere.com;Doug.Taylor@alere.com"
    ''' <summary>
    ''' Who will recieve Error Messages
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>11-Jun-2015 20:34</created><Author>CONCATENO\taylor</Author></revision></revisionHistory>
    Public Shared Property ErrorRecipient() As String
        Get
            Return myErrorRecipient
        End Get
        Set(ByVal value As String)
            myErrorRecipient = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' Email an error
    ''' </summary>
    ''' <param name="ex"></param>
    ''' <param name="show"></param>
    ''' <param name="ExtraInfo"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [18/04/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function EmailError(ByVal ex As Exception, _
  Optional ByVal show As Boolean = False, Optional ByVal ExtraInfo As String = " ") As Boolean
        Dim Log As New MedscreenLib.FileLogger(pvtErrorLoggingDirectory & "Error.txt")
        Dim ErrorString As String = ex.ToString

        If Not ex.TargetSite Is Nothing Then
            ErrorString += "-" & ex.TargetSite.Name
        End If

        ErrorString += " " & ExtraInfo & " User : " & Security.Principal.WindowsIdentity.GetCurrent.Name

        Medscreen.BlatEmail("Error in " & ExtraInfo, ErrorString, ErrorRecipient)
        If Security.Principal.WindowsIdentity.GetCurrent.Name.ToUpper = "CONCATENO\DEVMAN" Then Exit Function
        If show Then MsgBox(ErrorString, , "Error")
    End Function

    Public Overloads Shared Function EmailError(ByVal ErrorMessage As String) As Boolean
        Return Medscreen.QuckEmail("Error  Message: ", ErrorMessage)

    End Function

    '''<summary>
    ''' Log Timing
    ''' </summary>
    ''' <param name='ErrorString'>Message</param>
    ''' <param name='show'>Whether the request is shon</param>
    ''' <returns>Void</returns>
    Public Shared Function LogTiming(ByVal ErrorString As String, _
   Optional ByVal show As Boolean = False) As Boolean
        If blnNoLog Then Exit Function
        Dim Log As New MedscreenLib.FileLogger("Timing.txt")

        Log.WriteLog(ErrorString, Diagnostics.EventLogEntryType.Information)
        'If show Then MsgBox(ErrorString, , "Error")
    End Function

    Public Shared Property Logging() As Boolean
        Get
            Return Not blnNoLog
        End Get
        Set(ByVal Value As Boolean)
            blnNoLog = Not Value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set output Directory for error logging
    ''' </summary>
    ''' <param name="path">Path to use</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub SetLoggingDirectory(ByVal path As String)
        pvtErrorLoggingDirectory = path
        _errorLogFile = path & "Error" & Today.ToString("MMM-yyyy") & ".txt"
    End Sub
#End Region

#Region "MessageEventArgs"
    Public Class MessageEventArgs
        Private _message As String, _show As Boolean, _cancel As Boolean, _type As Diagnostics.EventLogEntryType

        Public Sub New(ByVal message As String, ByVal show As Boolean)
            Me.New(message, show, EventLogEntryType.Information)
        End Sub

        Public Sub New(ByVal message As String, ByVal show As Boolean, ByVal type As Diagnostics.EventLogEntryType)
            _message = message
            _show = show
            _type = type
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Gets or Sets a value indicating whether message will be cancelled
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        '''   Cancelled messages are not shown or written to the log
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Property Cancel() As Boolean
            Get
                Return _cancel
            End Get
            Set(ByVal Value As Boolean)
                _cancel = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Gets the type of message being logged
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Type() As Diagnostics.EventLogEntryType
            Get
                Return _type
            End Get
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Gets or Sets a valie indicating whether message will be displayed
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public Property Show() As Boolean
            Get
                Return _show
            End Get
            Set(ByVal value As Boolean)
                _show = value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        '''   Gets the message text
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' -----------------------------------------------------------------------------
        Public ReadOnly Property Message() As String
            Get
                Return _message
            End Get
        End Property

    End Class
#End Region
#Region "Reader"


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Read values from an XML Element 
    ''' </summary>
    ''' <param name="oread">XML Element to read from </param>
    ''' <param name="FieldName">Field name to find</param>
    ''' <param name="nvl"></param>
    ''' <param name="isNull"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function ReadValue(ByVal oread As Xml.XmlElement, ByVal FieldName As String, _
  ByVal nvl As Object, ByRef isNull As Boolean) As Object

        Dim aNode As Xml.XmlNode
        Dim objRet As Object = Nothing
        For Each aNode In oread
            If aNode.Name.ToLower = FieldName.ToLower Then
                Dim strValue As String = aNode.InnerText
                If strValue.Trim.Length = 0 Then
                    objRet = nvl
                Else
                    objRet = strValue
                End If
                Exit For
            End If
        Next
        Return objRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Read values from a node list 
    ''' </summary>
    ''' <param name="oread">Nodelist to read from </param>
    ''' <param name="FieldName">Field to find</param>
    ''' <param name="nvl">Value to use for nulls</param>
    ''' <param name="isNull"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Shared Function ReadValue(ByVal oread As Xml.XmlNodeList, ByVal FieldName As String, _
  ByVal nvl As Object, ByRef isNull As Boolean) As Object

        Dim aNode As Xml.XmlNode
        Dim objret As Object = Nothing
        For Each aNode In oread
            If aNode.Name.ToLower = FieldName.ToLower Then
                Dim strValue As String = aNode.InnerText
                If strValue.Trim.Length = 0 Then
                    objret = nvl
                Else
                    objret = strValue
                End If
                Exit For
            End If
        Next
        Return objret
    End Function

    '''<summary>
    ''' Read a value out of a OLEDB reader object
    ''' </summary>
    ''' <param name='Oread'>Reader</param>
    ''' <param name='FieldName'>Name of field to read</param>
    ''' <param name='nvl'>Null value, convert field to value if null</param>
    ''' <returns>Read field</returns>
    Public Overloads Shared Function ReadValue(ByVal Oread As OleDb.OleDbDataReader, _
  ByVal FieldName As String, _
  ByVal nvl As Object, ByRef isNull As Boolean) As Object
        Dim objret As Object = Nothing

        Dim intIndex As Integer
        Try
            intIndex = Oread.GetOrdinal(FieldName)
            If intIndex >= 0 Then 'Check that we have a valid field name, if not skip over it 
                If Oread.IsDBNull(intIndex) Then
                    isNull = True
                    objret = nvl
                Else
                    isNull = False
                    objret = Oread.GetValue(intIndex)
                End If
            End If
        Catch ex As Exception
            isNull = True
            objret = Nothing
        End Try
        Return objret
    End Function



    '''<summary>
    ''' Read a value out of a MSSQL reader object
    ''' </summary>
    ''' <param name='Oread'>Reader</param>
    ''' <param name='FieldName'>Name of field to read</param>
    ''' <param name='nvl'>Null value, convert field to value if null</param>
    ''' <returns>Read field</returns>
    Public Overloads Shared Function ReadValue(ByVal Oread As SqlClient.SqlDataReader, _
  ByVal FieldName As String, _
  ByVal nvl As Object, ByRef isNull As Boolean) As Object
        Dim objret As Object = Nothing

        Dim intIndex As Integer
        Try
            intIndex = Oread.GetOrdinal(FieldName)
            If intIndex >= 0 Then
                If Oread.IsDBNull(intIndex) Then
                    isNull = True
                    objret = nvl
                Else
                    isNull = False
                    objret = Oread.GetValue(intIndex)
                End If
            End If
        Catch ex As Exception
            isNull = True
            objret = Nothing
        End Try
        Return objret
    End Function


    '''<summary>
    ''' Read a value out of a reader object by position
    ''' </summary>
    ''' <param name='oread'>Reader</param>
    ''' <param name='intPosition'>Index of field</param>
    ''' <param name='myVar'>Value (returned)</param>
    ''' <param name='myDefault'>Null value, convert field to value if null</param>
    ''' <returns>True if succesful</returns>
    Public Shared Function ReadFromReader(ByVal oread As OleDb.OleDbDataReader, _
    ByVal intPosition As Integer, ByRef myVar As Object, _
    ByVal myDefault As Object) As Boolean

        If oread.IsDBNull(intPosition) Then
            myVar = myDefault
        Else
            If TypeOf myVar Is Boolean Then
                myVar = (CStr(oread.GetValue(intPosition)) = "T")
            Else
                myVar = oread.GetValue(intPosition)
            End If
        End If

    End Function


#End Region

#End Region

#Region "Shared procedures"
    '''<summary>
    ''' Set Login info in a crystal report
    ''' </summary>
    ''' <param name='cr'>Crystal Reprt Document</param>
    <CLSCompliant(False)> _
    Public Shared Sub CrLogin(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument)

        Dim objTab As CrystalDecisions.CrystalReports.Engine.Table
        Dim objLogInf As CrystalDecisions.Shared.TableLogOnInfo
        'Dim Query As CrystalDecisions.CrystalReports.Engine.DataDefinition
        For Each objTab In cr.Database.Tables
            objLogInf = objTab.LogOnInfo
            objLogInf.ConnectionInfo.ServerName = "john-live"
            objLogInf.ConnectionInfo.Password = "live"
            objLogInf.ConnectionInfo.UserID = "live"
            objTab.ApplyLogOnInfo(objLogInf)
        Next

    End Sub


    '''<summary>
    ''' Display the mapping name for an object in a message box</summary>
    Public Shared Sub ShowMappingName(ByVal src As Object)
        Dim list As IList = Nothing
        Dim t As Type = Nothing



        If TypeOf (src) Is Array Then
            t = src.GetType()
            list = CType(src, IList)
        Else
            If TypeOf src Is System.ComponentModel.IListSource Then
                src = CType(src, System.ComponentModel.IListSource).GetList()
            End If

            If TypeOf src Is IList Then
                t = src.GetType()
                list = CType(src, IList)
            Else
                MsgBox("Error")
                Return
            End If
        End If


        If TypeOf list Is System.ComponentModel.ITypedList Then
            MsgBox(CType(list, System.ComponentModel.ITypedList).GetListName(Nothing))
        Else
            MsgBox(t.Name)
        End If

    End Sub

    '''<summary>
    ''' Set the source column for a parameter
    ''' </summary>
    ''' <param name='Parameters'>Parameters to set</param>
    Public Shared Sub SetParameterSourceColumns(ByVal Parameters As OleDb.OleDbParameterCollection)
        Dim myparameter As OleDb.OleDbParameter
        For Each myparameter In Parameters
            myparameter.SourceColumn = myparameter.ParameterName
        Next
    End Sub


#End Region

#Region "Shared Properties"

    '''<summary>
    ''' The directory used to store XML
    ''' </summary>
    ''' <result>XML Directory as a string</result>
    Public Shared Property XMLDirectory() As String
        Get
            Return pvtXMLDirectory
        End Get
        Set(ByVal Value As String)
            pvtXMLDirectory = Value
        End Set
    End Property

    Public Shared Property ResultsDirectory() As String
        Get
            Return pvtResultsDirectory
        End Get
        Set(ByVal value As String)
            pvtResultsDirectory = value
        End Set
    End Property

    Private Shared pvtBackupFolder As String = ""
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Folder for back up drive
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property BackupFolder() As String
        Get
            Return pvtBackupFolder
        End Get
        Set(ByVal Value As String)
            pvtBackupFolder = Value
        End Set
    End Property

    Private Shared pvtBackupDrive As String = ""
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Location of the drive backups will be mapped to
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	03/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Property BackupDrive() As String
        Get
            Return pvtBackupDrive
        End Get
        Set(ByVal Value As String)
            pvtBackupDrive = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' mapping for gcmsdata drive
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/09/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared ReadOnly Property GCMSData() As String
        Get
            Return Constants.FileServer & "GCMSData"
        End Get
    End Property

    '''<summary>
    ''' IniFiles location of inifiles associated with application
    ''' </summary>
    Public Shared Property IniFiles() As String
        Get
            Return pvtInifiles
        End Get
        Set(ByVal Value As String)
            pvtInifiles = Value
        End Set
    End Property

    '''<summary>
    ''' Location of .XSL files XSL transforms
    ''' </summary>
    Public Shared Property XSLTransformsDirectory() As String
        Get
            If pvtXSLDirectory Is Nothing Then
                pvtXSLDirectory = Constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
            End If
            If pvtXSLDirectory.Trim.Length = 0 Then
                pvtXSLDirectory = Constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
            End If
            Return pvtXSLDirectory
        End Get
        Set(ByVal Value As String)
            pvtXSLDirectory = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Validate sage id 
    ''' </summary>
    ''' <param name="sageId"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	08/01/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Shared Function IsValidSageID(ByVal sageId As String) As Boolean
        Dim blnRet As Boolean = False
        Dim strRegEx As String = MedscreenlibConfig.Validation.Item("SageMask") ' ^([a-zA-Z0-9_&\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$ 'Email address validator
        'the next line must be commented out before going live
        'strRegEx = "^[A-Z][0-9][0-9][0-9][0-9]|^[A-Z][A-Z][0-9][0-9][0-9][0-9]|^[A-Z][A-Z][A-Z][0-9][0-9][0-9]|^[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]"
        Dim Regex As New System.Text.RegularExpressions.Regex(strRegEx)
        Dim myMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(sageId, strRegEx)
        blnRet = myMatch.Success
        Return blnRet

    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Check to see that the email address is valid
    ''' </summary>
    ''' <param name="Email"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [24/11/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Function ValidateEmail(ByVal Email As String) As Constants.EmailErrors
        Dim myReturn As Constants.EmailErrors = Constants.EmailErrors.None
        Dim intPos As Integer
        If Email.Trim.Length > 0 Then 'We have something to work on 
            intPos = InStr(Email, "@")       'See if we have an 'at
            If intPos = 0 Then                  'No hat
                myReturn = Constants.EmailErrors.NoAt
            Else                                'got an hat 
                Dim intpos2 As Integer = InStr(intPos, Email, ".")
                If intpos2 = 0 Then             'Is there a '.' after the @
                    myReturn = Constants.EmailErrors.NoDomain
                End If
            End If
            If InStr(Email.ToUpper, "INVALID") > 0 Then
                myReturn = Constants.EmailErrors.Invalid
            End If

            'check for a valid email set 
            Dim intPosSet As Integer = Medscreen.isValidEmailSet(Email)
            If intPosSet <> -1 Then
                myReturn = Constants.EmailErrors.InvalidSet
            End If
        Else
            myReturn = Constants.EmailErrors.NoAddress
        End If
        Return myReturn
    End Function

    '''<summary>
    ''' Location of the root on Live typically "\\john\live\"
    ''' </summary>
    Public Shared Property RootLive() As String
        Get
            Return strLiveRoot
        End Get
        Set(ByVal Value As String)
            strLiveRoot = Value
        End Set
    End Property

    '''<summary>
    ''' Location of the standard CSS Style Sheet. Typically "\\EM01\intranet\MedScreen.css"
    ''' </summary>
    Public Shared Property StyleSheet() As String
        Get
            Return strStyleSheet
        End Get
        Set(ByVal Value As String)
            strStyleSheet = Value
        End Set
    End Property

    Public Shared ReadOnly Property HTMLStyleHeader() As String
        Get
            Return "<HTML><HEAD><LINK REL=""STYLESHEET"" TYPE=""text/css""" & _
   "HREF = ""\\EM01\intranet\MedScreen.css"" Title=""Style""/><TITLE></TITLE></HEAD>"

        End Get
    End Property


#End Region
#End Region


    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], _
         ByVal lpszDomain As [String], ByVal lpszPassword As [String], _
         ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, _
         ByRef phToken As IntPtr) As Boolean

    '''<summary>
    ''' Kernel32 Function
    ''' </summary>
    ''' <param name='handle'>Handle to close</param>
    ''' <returns>True handle closed</returns>
    Public Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal handle As IntPtr) As Boolean

    Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

    '''<summary>
    ''' Format a windows error message
    ''' </summary>
    ''' <param name='dwFlags'></param>
    <DllImport("kernel32.dll")> _
     Public Shared Function FormatMessage(ByVal dwFlags As Integer, ByRef lpSource As IntPtr, _
         ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByRef lpBuffer As [String], _
         ByVal nSize As Integer, ByRef Arguments As IntPtr) As Integer

    End Function

    Private Shared Function FirstTime(ByVal dat1 As Date) As TimeSpan
        Dim datTemp As Date

        If dat1.Hour <= 8 Then ' Early morning start 
            datTemp = DateSerial(dat1.Year, dat1.Month, dat1.Day)
            datTemp = datTemp.AddHours(8)
            ' 8 in the morning
            Return datTemp.Subtract(dat1)
        Else ' late night start going through 
            datTemp = DateSerial(dat1.Year, dat1.Month, dat1.Day + 1)
            'datTemp = datTemp.AddHours(8)
            Return datTemp.Subtract(dat1)
        End If

    End Function

    Private Shared Function SecondTime(ByVal Dat1 As Date, ByVal dat2 As Date) As TimeSpan
        Dim datTemp As Date

        If dat2.Hour >= 18 Then 'late finish 
            datTemp = DateSerial(Dat1.Year, Dat1.Month, Dat1.Day)
            datTemp = datTemp.AddHours(18)
            TimeSerial(18, 0, 0) ' 6 in the evening same day
            Return dat2.Subtract(datTemp)
        ElseIf Dat1.Day <> dat2.Day Then 'early hours from previous day
            datTemp = DateSerial(dat2.Year, dat2.Month, dat2.Day)
            'datTemp = datTemp.AddHours(18)
            Dim ts As TimeSpan = dat2.Subtract(datTemp)
            If Dat1.Hour >= 18 Then
                Dim minutes As Integer = (24 - Dat1.Hour) * 60 - Dat1.Minute
                Dim hr As Integer = CInt(minutes / 60)
                minutes = minutes - hr * 60
                ts = ts.Add(New TimeSpan(hr, minutes, 0))
            Else
                ts = ts.Add(New TimeSpan(6, 0, 0))
            End If
            Return ts

        Else 'early hours
            datTemp = DateSerial(dat2.Year, dat2.Month, dat2.Day)
            'datTemp = datTemp.AddHours(18)
            Return dat2.Subtract(datTemp)
        End If

    End Function

End Class


#Region "Constants"


'''<summary>
''' Public Constants
''' </summary>
''' <remarks>
''' Constants these values are in the main replicated in LIB_CONSTANTS in the Sample Manager source files
''' </remarks>
Public Class Constants

#Region "Public Enums"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Address Types
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum AddressType
        '''<summary>Null type</summary>
        AddressTypeNull = 0
        '''<summary>Customer company address</summary>
        AddressMain = 1    ' Customer company address }
        '''<summary>Customer second address</summary>
        AddressSec = 2     ' Customer second address }
        '''<summary>Customer Invoice Mailing address</summary>
        AddressInv = 3     ' Customer Invoice Mailing address }
        '''<summary>Vessel Invoice address</summary>
        VesselAddressInv = 4   ' Vessel Invoice address }
        '''<summary>Address stored in address table</summary>
        TableAddress = 5        ' Address stored in address table }
        '''<summary>Adhoc Sample Company address</summary>
        AdhocAddressMain = 6   ' Adhoc Sample Company address }
        '''<summary>Adhoc Sample Invoice Mailing address</summary>
        AdhocAddressInv = 7    ' Adhoc Sample Invoice Mailing address }
        '''<summary>Vessel invoice mailing address</summary>
        VesselAddressMail = 8  ' Vessel invoice mailing address }
        '''<summary>Delivery (shipping) address</summary>
        DeliveryAddress = 9     ' Delivery (shipping) address }
        '''<summary>Site Address</summary>
        AddressSite = 10
        '''<summary>Collecting Officer Address</summary>
        AddressCollOfficer = 11
        '''<summary>Bank Address</summary>
        AddressBank = 12
        '''<summary>Invoice Mail address</summary>
        InvoiceMail = 13
        '''<summary>Invoice Shipping Address</summary>
        InvoiceShipping = 14
        '''<summary>GoldMine Address</summary>
        GoldmineAddress = 15
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Posible errors in an Email address 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum EmailErrors
        ''' <summary>No detectable errors</summary>
        None
        ''' <summary>@ sign missing</summary>
        NoAt
        ''' <summary>no . follwing @ symbol</summary>
        NoDomain
        '''<summary>no address provided </summary>
        NoAddress
        '''<summary>Preset invalid </summary>
        Invalid
        InvalidSet
    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Possible errors with a phone number
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum PhoneNoErrors
        ''' <summary>No Error</summary>
        NoError
        ''' <summary>international Dailling code present</summary>
        InTDialCodePresent
        ''' <summary>Illegal Character present</summary>
        IllegalCharacterPresent
    End Enum
#End Region

    '''<summary>EMail to Fax service provider</summary>
    Public Const GCST_EMAIL_To_Fax_provider As String = "starfax.co.uk"

    '''<summary>EMail to Fax service provider send string </summary>
    Public Const GCST_EMAIL_To_Fax_Send As String = "@" & GCST_EMAIL_To_Fax_provider

    Public Const GCST_BLAT_DIRECTORY As String = "\\corp.concateno.com\medscreen\common\lab programs\blat"

    '''<summary>Request for Card Details from ICP</summary>
    Public Const GCST_IFace_StatusCardDetails As String = "1"
    '''<summary>Request for an Address </summary>
    Public Const GCST_IFace_AddressRequst As String = "A"
    '''<summary>Task completed</summary>
    Public Const GCST_IFace_StatusCreated As String = "C"
    '''<summary>Request for a change</summary>
    Public Const GCST_IFace_StatuscHange As String = "H"
    '''<summary>Collection completed</summary>
    Public Const GCST_IFace_StatusJobMade As String = "J"
    '''<summary>Task Request</summary>
    Public Const GCST_IFace_TaskRequest As String = "K"
    '''<summary>Invoice Paid</summary>
    Public Const GCST_IFace_StatusPaid As String = "P"
    '''<summary>ICP Request</summary>
    Public Const GCST_IFace_StatusICPRequest As String = "M"
    '''<summary>Create a new collection</summary>
    Public Const GCST_IFace_StatusNew As String = "N"
    '''<summary>Temporary status</summary>
    Public Const GCST_IFace_StatusTemp As String = "T"
    '''<summary>Delete Collection</summary>
    Public Const GCST_IFace_StatusDelete As String = "X"
    '''<summary>Collection Invoiced</summary>
    Public Const GCST_IFace_StatusInvoiced As String = "I"
    '''<summary>Task failed</summary>
    Public Const GCST_IFace_StatusFailed As String = "F"
    '''<summary>Task Locked</summary>
    Public Const GCST_IFace_StatusLocked As String = "L"
    '''<summary>Request</summary>
    Public Const GCST_IFace_StatusRequest As String = "R"
    Public Const GCST_IFace_StatusAvailable As String = "V"

    '''<summary>Job Task</summary>
    Public Const GCST_IFace_TaskJob As String = "JOB"
    '''<summary>Pre pay invoice (Record payment)</summary>
    Public Const GCST_IFace_TaskPrepay As String = "TSKRCPY"
    '''<summary>Not used</summary>
    Public Const GCST_IFace_TaskSample As String = "SAMPLE"
    'Public Const GCST_IFace_TaskNewDate = "NEWDATE"
    ''Public Const GCST_IFace_TaskNewAddr = "TSKNEWADDRESS"
    ''Public Const GCST_IFace_TaskNewFixed = "TSKNEWF"
    ''Public Const GCST_IFace_TaskNewRoutine = "TSKNEWR"
    ''Public Const GCST_IFace_TaskNewCallOut = "TSKNEWC"
    '''<summary>Cancel Collection (Obsolete)</summary>
    Public Const GCST_IFace_TaskCancel As String = "TSKCANC"
    '''<summary>Confirm Collection</summary>
    Public Const GCST_IFace_TaskJobConfirm As String = "TSKCONFM"
    '''<summary>Task Job in Progress (Obsolete)</summary>
    Public Const GCST_IFace_TaskJobProgress As String = "TSKPROG"
    '''<summary>Task Send to Officer (Obsolete)</summary>
    Public Const GCST_IFace_TaskJobSendToCO As String = "TSKSEND"
    '''<summary>Task Job Cancel Invoice</summary>
    Public Const GCST_IFace_TaskJobInvCancel As String = "TSKINCN"
    '''<summary>Change Password Request</summary>
    Public Const GCST_IFACE_TaskPasswordChange As String = "CHPW"


    '''<summary>Standard Sample Cost (Don't use get from price list)</summary>
    Public Const cstStandardSampleCost As Double = 65
    '''<summary>Standard Fixed Site Cost (Don't use get from price list)</summary>
    Public Const cstStandardFixedSiteCost As Double = 85
    '''<summary>Standard minimum Collection Charge (Don't use get from price list)</summary>
    Public Const cstStandardMinCollectionCharge As Double = 325
    '''<summary>Standard Site charge (Don't use get from price list)</summary>
    Public Const gcstStandardSiteCharge As Double = 20

    '''<summary>Database Representation of TRUE</summary>
    Public Const C_dbTRUE As Char = "T"c
    '''<summary>Database Representation of FALSE</summary>
    Public Const C_dbFALSE As Char = "F"c
    '''<summary>Date representing NULL</summary>
    Public Const C_InitDate As Date = #1/1/1850#



    '''<summary>Collection Type Analysis Only</summary>
    Public Const GCST_JobTypeAnalysisOnly As String = "A" '
    '''<summary>Collection Type Call Out UK</summary>
    Public Const GCST_JobTypeCallout As String = "C"
    '''<summary>Collection Type Call Out Overseas</summary>
    Public Const GCST_JobTypeCalloutOverseas As String = "O"
    '''<summary>Collection Type UK Routine</summary>
    Public Const GCST_JobTypeWorkplace As String = "U"
    '''<summary>Collection Type Overseas Routine</summary>
    Public Const GCST_JobTypeOverseas As String = "W"
    '''<summary>Collection Type Fixed Site</summary>
    Public Const GCST_JobTypeFixed As String = "F"
    '''<summary>Collection Type Obsolete</summary>
    Public Const GCST_JobTypeObsolete As String = "X"
    '''<summary>Collection Type External</summary>
    Public Const GCST_JobTypeExternal As String = "E"


    'Certificate constants
    '''<summary>Esso Certificates</summary>
    Public Const GCST_ESSOCertifcate As String = "ESSO_CERT"
    '''<summary>LUL Certificates</summary>
    Public Const GCST_LULCertifcate As String = "LUL_CERT"
    '''<summary>Railtrack certificates</summary>
    Public Const GCST_RAILCertifcate As String = "RAIL_CERT"


    ' Collection status constants }
    '''<summary>Collection has been confirmed</summary>
    Public Const GCST_JobStatusConfirmed As String = "C"
    '''<summary>Collection has been created</summary>
    Public Const GCST_JobStatusCreated As String = "V"
    '''<summary>Collection has been assigned to collecting officer</summary>
    Public Const GCST_JobStatusAssigned As String = "W"
    '''<summary>Collection has been Cancelled</summary>
    Public Const GCST_JobStatusCancelled As String = "X"
    '''<summary>Collection has been done, Officer has informed</summary>
    Public Const GCST_JobStatusCollected As String = "D"
    '''<summary>Collection has Samples in LAB</summary>
    Public Const GCST_JobStatusReceived As String = "R"
    '''<summary>Collection has been sent to officer</summary>
    Public Const GCST_JobStatusSent As String = "P"
    '''<summary>Indicates that it is a payment psuedo collection</summary>
    Public Const GCST_JobStatusPay As String = "S"
    '''<summary>Collection has been Approved</summary>
    Public Const GCST_JobStatusApproved As String = "A"
    '''<summary>Collection has been Completed all samples processed and reported</summary>
    Public Const GCST_JobStatusCommitted As String = "M"
    '''<summary>Collection has been put on hold</summary>
    Public Const GCST_JobStatusOnHold As String = "H"
    '''<summary>Collection is not yet created</summary>
    Public Const GCST_JobStatusTemp As String = "T"
    Public Const GCST_JobStatusInterrupted As String = "I"

    'Job Status constants}
    '''<summary>Job is complete</summary>
    Public Const GCST_JobJobStatusCompleted As String = "C"
    '''<summary>Job is available for processing</summary>
    Public Const GCST_JobJobStatusAvailable As String = "V"
    '''<summary>Job has been authorised</summary>
    Public Const GCST_JobJobStatusAuthorised As String = "A"
    '''<summary>Job has been suspended</summary>
    Public Const GCST_JobJobStatusSuspended As String = "S"
    '''<summary>Job is Cancelled</summary>
    Public Const GCST_JobJobStatusCancelled As String = "X"
    '''<summary>Job has been rejected</summary>
    Public Const GCST_JobJobStatusRejected As String = "R"
    '''<summary>Job is undergoing inspection</summary>
    Public Const GCST_JobJobStatusInspection As String = "I"

    'Cancellation constants
    '''<summary>Cancellation status Don't invoice</summary>
    Public Const GCST_CANC_NoInvoice As String = "X"
    '''<summary>Cancellation status Waiting on Collector's time sheet</summary>
    Public Const GCST_CANC_WaitingTimeSheet As String = "W"
    '''<summary>Cancellation status Ready to be reviewed</summary>
    Public Const GCST_CANC_Ready As String = "R"
    '''<summary>Cancellation status Not cancelled</summary>
    Public Const GCST_CANC_NotCancelled As String = "N"
    '''<summary>Cancellation status Problem raised</summary>
    Public Const GCST_CANC_Problem As String = "P"
    '''<summary>Cancellation status Invoiced</summary>
    Public Const GCST_CANC_InvoiceD As String = "I"

    'Customer status constants 
    '''<summary>Customer Profile Status New</summary>
    Public Const GCST_Cust_Status_New As String = "N"
    '''<summary>Customer Profile Status Suspended</summary>
    Public Const GCST_Cust_Status_Suspended As String = "S"
    '''<summary>Customer Profile Status Active</summary>
    Public Const GCST_Cust_Status_Active As String = "A"
    '''<summary>Customer Profile Status Removed</summary>
    Public Const GCST_Cust_Status_Removed As String = "X"

    'Workflow
    '''<summary>Small box height</summary>
    Public Const BoxHeightSmall As Integer = 30
    '''<summary>Medium box height</summary>
    Public Const BoxHeightMedium As Integer = 45
    '''<summary>Large box height</summary>
    Public Const BoxHeightLarge As Integer = 60
    '''<summary>Gap between boxes</summary>
    Public Const BoxGap As Integer = 20

    'Address Constants}
    '''<summary>Address type main address</summary>
    Public Const GCST_Cust_Address_Main As Integer = 1
    '''<summary>Address type Shipping</summary>
    Public Const GCST_Cust_Address_Sec As Integer = 2
    '''<summary>Address type Invoicing</summary>
    Public Const GCST_Cust_Address_Inv As Integer = 3
    '''<summary>Address type Vessel Invoicing</summary>
    Public Const GCST_Vessel_Address_Inv As Integer = 4
    '''<summary>Table based address</summary>
    Public Const GCST_Table_Address As Integer = 5
    '''<summary>Adhoc address</summary>
    Public Const GCST_Adhoc_Address_Main As Integer = 6
    '''<summary>Adhoc address invoice</summary>
    Public Const GCST_Adhoc_Address_Inv As Integer = 7
    '''<summary>Vessel Address</summary>
    Public Const GCST_Vessel_Address_Mail As Integer = 8
    '''<summary>Delivery address</summary>
    Public Const GCST_Delivery_Address As Integer = 9
    '''<summary>Delivery Address</summary>
    Public Const GCST_Address_Site As Integer = 10
    '''<summary>Collecting Officer Address</summary>
    Public Const GCST_Address_CollOfficer As Integer = 11
    '''<summary>Bank Address</summary>
    Public Const GCST_Address_Bank As Integer = 12
    '''<summary>invoicing mail Address</summary>
    Public Const GCST_Address_InvMail As Integer = 13
    '''<summary>invoicing shipping Address</summary>
    Public Const GCST_Address_InvShip As Integer = 14

    ' Address type constants (semi-descriptive as used in address table) }
    '''<summary>Address Type Vessel Invoice</summary>
    Public Const GCST_AddrType_VesselInv As String = "VSIN"
    '''<summary>Address Type Vessel Mail</summary>
    Public Const GCST_AddrType_VesselMail As String = "VSML"
    '''<summary>Address Type Customer</summary>
    Public Const GCST_AddrType_Customer As String = "CUST"
    '''<summary>Address Type Centre</summary>
    Public Const GCST_AddrType_Centre As String = "CENT"

    'Server IP Addresses
    '''<summary>EM01 IP Address</summary>
    Public GCST_IP_EM01 As String = "172.16.0.8"
    '''<summary>MATT IP Address</summary>
    Public Const GCST_IP_Matt As String = "10.1.2.201"


    'UserInfo constants
    '''<summary>User Index by Domain and ID</summary>
    Public Const GCST_User_FullNetworkID As String = "WindowsIdentity+"
    '''<summary>User Index by network ID</summary>
    Public Const GCST_User_NetworkID As String = "WindowsIdentity"
    '''<summary>User Index by SMID</summary>
    Public Const GCST_User_Identity As String = "Identity"
    '''<summary>Return User Email</summary>
    Public Const GCST_User_Email As String = "Email"
    '''<summary>Return User Description (fuill name)</summary>
    Public Const GCST_User_Name As String = "Description"
    '''<summary>Return User's Manager</summary>
    Public Const GCST_User_Manager As String = "Manager_ID"
    '''<summary>Return User's Department</summary>
    Public Const GCST_User_Department As String = "Department_ID"
    '''<summary>Return User's Location</summary>
    Public Const GCST_User_Location As String = "Location_ID"


    'Payment Type Constants}

    '''<summary>
    ''' Types of payment
    ''' </summary>
    Public Enum PaymentTypes
        '''<summary>Payment by card</summary>
        PaymentCard = 1
        '''<summary>Payment by cheque</summary>
        PaymentCheque = 2
        '''<summary>Payment by Cash</summary>
        PaymentCash = 3
    End Enum

    '''<summary>
    ''' Customer Confirmation messages types
    ''' </summary>
    Public Enum ConfirmationType
        '''<summary>Confirmation on Arranging</summary>
        Arrange = Asc("A")
        '''<summary>Confirmation on Completion</summary>
        Confirm = Asc("C")      'Completed 
    End Enum


    '''<summary>
    ''' Methods of sending documents
    ''' </summary>
    Public Enum SendMethod
        '''<summary>Send by Email</summary>
        Email = Asc("E")
        '''<summary>Send by Fax</summary>
        Fax = Asc("F")
        '''<summary>Send as PDF</summary>
        PDF = Asc("A")
        '''<summary>Send to Printer</summary>
        Printer = Asc("P")
        '''<summary>Phone</summary>
        Verbal = Asc("V")
        '''<summary>Not defined</summary>
        NoConfirm = Asc("N")
        HTML = Asc("H")
        Excel = Asc("X")
        RTF = Asc("R")
    End Enum

    '''<summary>Payment by Credit card </summary>
    Public Const GCST_PaymentCard As String = "1"
    '''<summary>Payment by Cheque</summary>
    Public Const GCST_PaymentCheque As String = "2"
    '''<summary>Payment by Cash</summary>
    Public Const GCST_PaymentCash As String = "3"

    '''<summary>Starting abbreviation for a Collection</summary>
    Public Const GCST_Collection_Abbr As String = "CM"

    '''<summary>Send to Officer by (print out)</summary>
    Public Const GCST_CollOffSendPrinter As String = "PRINTER"
    '''<summary>Send to Officer by FAX (home)</summary>
    Public Const GCST_CollOffSendHomeFax As String = "Contact By HOME FAX"
    '''<summary>Send to Officer by FAX (work)</summary>
    Public Const GCST_CollOffSendWorkFax As String = "Contact By WORK FAX"
    '''<summary>Send to Officer by E-Mail (home)</summary>
    Public Const GCST_CollOffSendHomeEmail As String = "Contact By HOME E_MAIL"
    '''<summary>Send to Officer by E-Mail (Office)</summary>
    Public Const GCST_CollOffSendWorkEmail As String = "Contact By WORK E_MAIL"
    '''<summary>Contact Officer by Home Phone </summary>
    Public Const GCST_CollOffSendHomePhone As String = "Contact By HOME PHONE"
    '''<summary>Contact Officer by Work Phone</summary>
    Public Const GCST_CollOffSendWorkPhone As String = "Contact By WORK PHONE"

    ' public consts for Common Invoice Items }
    '''<summary>Old style sample charge</summary>
    Public Const GCST_LineItem_OSample As String = "OPS0003"     ' Old style sample charge }
    '''<summary>Old style minimum charge </summary>
    Public Const GCST_LineItem_OMinCharge As String = "OPS0002"  ' Old style minimum charge }
    '''<summary>Old style maximum charge</summary>
    Public Const GCST_LineItem_OMaxCharge As String = "OPS0008"  ' Old style maximum charge }
    '''<summary>Old style MRO charge</summary>
    Public Const GCST_LineItem_OMRO As String = "OPS0001"        ' Old style MRO charge }
    '''<summary>Old style collection charge</summary>
    Public Const GCST_LineItem_Comment As String = "OPS0004"     ' Old style collection charge }
    '''<summary>Old style Additional Test charge</summary>
    Public Const GCST_LineItem_OTest As String = "OPS0005"       ' Old style Additional Test charge }
    '''<summary>Old style Collection charge</summary>
    Public Const GCST_LineItem_OCollection As String = "OPS0006" ' Old style Collection charge }
    '''<summary>Old style Collection Time charge</summary>
    Public Const GCST_LineItem_OTime As String = "OPS0007"       ' Old style Collection Time charge }
    '''<summary>Annual retainer</summary>
    Public Const GCST_LineItem_Retainer As String = "RTNR001"    ' Annual retainer }
    '''<summary>Collection Cancellation Routine</summary>
    Public Const GCST_LineItem_Cancel_Routine As String = "COLS004"    ' Collection Cancellation }
    '''<summary>Collection Cancellation CallOut</summary>
    Public Const GCST_LineItem_Cancel_Oncall As String = "COLC004"    ' Collection Cancellation }
    '''<summary>Collection Cancellation Fixed Site</summary>
    Public Const GCST_LineItem_Cancel_Fixed As String = "COLF002"    ' Collection Cancellation }
    ' ENHA10286 Tim Moule 6/11/2003 Constants for cancelled LUL collections }
    '''<summary>Collection Cancellation LUL CallOut</summary>
    Public Const GCST_LineItem_Cancel_LULOncall As String = "COLL001"    ' Collection Cancellation }
    '''<summary>Collection Cancellation LUL InHours</summary>
    Public Const GCST_LineItem_Cancel_LULInhours As String = "COLL002"    ' Collection Cancellation }
    '''<summary>Collection Cancellation LUL Out of hours</summary>
    Public Const GCST_LineItem_Cancel_LULOuthours As String = "COLL003"    ' Collection Cancellation }

    '------------------------------------------------------------------------------}
    '------------------------- Collection Type Constants --------------------------}
    '------------------------------------------------------------------------------}
    ' Generic types used in invoicing, these are not used in CollectionType field }

    '''<summary>Collection Type Routine</summary>
    Public Const GCST_CollType_Routine As String = "S"     ' Site (routine) }
    '''<summary>Collection Type Call out UK</summary>
    Public Const GCST_CollType_CallOut As String = "C"     ' Call out }
    ' Types used both as generic types and as values for CollectionType field }
    '''<summary>Collection Type Fixed Site</summary>
    Public Const GCST_CollType_FixedSite As String = "F"   ' Fixed Site }
    '''<summary>Collection Type Analysis Only</summary>
    Public Const GCST_CollType_AnalOnly As String = "A"    ' Analysis only collection }
    '''<summary>Collection Type Not used</summary>
    Public Const GCST_CollType_Obsolete As String = "X"
    ' Specific types - These are only used for Collection.CollectionType }
    '''<summary>Collection Type Call out UK</summary>
    Public Const GCST_CollType_CallOutUK As String = "C"   ' UK Call outs }
    '''<summary>Collection Type Call out overseas</summary>
    Public Const GCST_CollType_CalloutOS As String = "O"   ' Overseas Call outs }
    '''<summary>Collection Type Routine</summary>
    Public Const GCST_CollType_RoutineUK As String = "U"   ' UK Routine (workplace) }
    '''<summary>Collection Type Routine Overseas</summary>
    Public Const GCST_CollType_RoutineOS As String = "W"   ' Overseas (Worldwide) Routine }


    '''<summary>Table header (start) in HTML</summary>
    Public Const cstTableHead As String = "<table class=fixwidnb><thead>"
    '''<summary>Table header element (HTML)</summary>
    Public Const cstColTableHead As String = "<th class=noborder width as string = "

    '''<summary>HTML CURRENCY indicator </summary>
    Public Const cstHTMLCurrency As Integer = 1

    '''<summary>
    ''' Collecting Officer send methods
    ''' </summary>
    Public Enum CollOfficerSend
        '''<summary>Send by Printer</summary>
        Printer
        '''<summary>Send To Fax at home</summary>
        HomeFax
        '''<summary>Send To Fax at work address</summary>
        WorkFax
        '''<summary>Send to Home Email address</summary>
        HomeEmail
        '''<summary>Sendto work Email </summary>
        WorkEmail
        '''<summary>Use Work Phone number</summary>
        WorkPhone
        '''<summary>Use Home Phone Number</summary>
        HomePhone
        '''<summary>Not Defined</summary>
        NotDefined
    End Enum

    '''<summary>
    ''' Report types
    ''' </summary>
    Public Enum ReportType
        '''<summary>Collection Request Report</summary>
        CollectionRequest
        '''<summary>Collection Arranged Report (to customer)</summary>
        CollectionArrangeConf
        '''<summary>Collection Completed Report </summary>
        CollectionCompleteConf
    End Enum

    '''<summary>
    ''' Capitilisation enumeration constants
    ''' </summary>
    Public Enum ChangeCase
        '''<summary>Force to upper case</summary>
        ChCUpper
        '''<summary>Force to lower case</summary>
        ChCLower
        '''<summary>Capitalise only the first letter</summary>
        ChCSentenace
        '''<summary>Capitalise all first letters</summary>
        ChCCapitalise
    End Enum

    'Comment type constants 

    '''<summary>Collection cancellation comment</summary>
    Public Const GCST_COMM_CollCancel As String = "CANC"
    '''<summary>Collection Invoice comment</summary>
    Public Const GCST_COMM_CollInvoice As String = "CINV"
    '''<summary>Collection cancellation invoice comment</summary>
    Public Const GCST_COMM_CollCancelInvoice As String = "CAIN"
    '''<summary>Collection confirmation comment</summary>
    Public Const GCST_COMM_CollConfirm As String = "CONF"
    '''<summary>Collection sent to Officer comment</summary>
    Public Const GCST_COMM_CollSendToOfficer As String = "SENT"
    '''<summary>Collection resent to Officer comment</summary>
    Public Const GCST_COMM_CollReSendToOfficer As String = "RSNT"
    '''<summary>Collection Information sent to Officer comment</summary>
    Public Const GCST_COMM_CollCOOfficerInfo As String = "INFO"
    '''<summary>Collection correction made to collection comment</summary>
    Public Const GCST_COMM_CCFix As String = "FIX"                  'Discrepency between number collected and booked fixed
    '''<summary>Collection error made comment</summary>
    Public Const GCST_COMM_ERROR As String = "ERR"
    '''<summary>Collection created comment</summary>
    Public Const GCST_COMM_NEWColl As String = "NEW"
    Public Const GCST_COMM_NEWInstrument As String = "INST"

    'Site type constants
    '''<summary>Customer Site</summary>
    Public Const GCSTCustSite As String = "CUSTSITE"
    '''<summary>Maritime Customer Site</summary>
    Public Const GCSTMaritimeSite As String = "MARTIMSITE"


    Public Sub New()

    End Sub

    Private Shared GCSTXDRIVE As String = "\\corp.concateno.com\medscreen\common"
    Private Shared GCSTBackupDrive As String = GCSTXDRIVE & "\limsfiles"
    Private Shared File_Server As String = "\\corp.concateno.com\medscreen\"
    Private Shared CollAddress As String = "toxeu.CollAdmin@alere.com"
    Private Shared myPasswordBodyFile As String = GCSTXDRIVE & "\Lab Programs\DBReports\Templates\EmailCover.txt"
    Private Shared GCSTCertsDirectory As String = "\certs\"
    Private Shared GCSTEmailResults As String = "\Dm_datafile\emails\"
    Private Shared GCSTWordoutput As String = "WordOutput"


    Public Shared Property WordOutput() As String
        Get
            Return MedConnection.Instance.ServerPath & GCSTWordoutput
        End Get
        Set(ByVal value As String)
            GCSTWordoutput = value
        End Set
    End Property

    Public Shared Property EmailResults() As String
        Get
            Return MedConnection.Instance.ServerPath & GCSTEmailResults
        End Get
        Set(ByVal value As String)
            GCSTEmailResults = value
        End Set
    End Property

    Public Shared Property CertsDirectory() As String
        Get
            Return MedConnection.Instance.ServerPath & GCSTCertsDirectory
        End Get
        Set(ByVal value As String)
            GCSTCertsDirectory = value
        End Set
    End Property

    Public Shared Property BackupFolder() As String
        Get
            Return GCSTBackupDrive
        End Get
        Set(ByVal value As String)
            GCSTBackupDrive = value
        End Set
    End Property

    ''' <summary>
    '''     Location of the X Drive
    ''' </summary>
    ''' <value>
    '''     <para>
    '''         
    '''     </para>
    ''' </value>
    ''' <remarks>
    '''     
    ''' </remarks>
    Public Shared Property GCST_X_DRIVE() As String
        Get
            Return GCSTXDRIVE
        End Get
        Set(ByVal Value As String)
            GCSTXDRIVE = Value
        End Set
    End Property


    Public Shared Property PasswordBodyFile() As String
        Get
            Return myPasswordBodyFile
        End Get
        Set(ByVal Value As String)
            myPasswordBodyFile = Value
        End Set
    End Property

    ''' <summary>
    '''     Location of the file server
    ''' </summary>
    ''' <value>
    '''     <para>
    '''         
    '''     </para>
    ''' </value>
    ''' <remarks>
    '''     
    ''' </remarks>
    Public Shared Property FileServer() As String
        Get
            Return File_Server
        End Get
        Set(ByVal Value As String)
            File_Server = Value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    ''' <summary>
    '''     Email Address to send Collections from  using OUTPUT_FROM command
    ''' </summary>
    Public Shared Property CollectionEmailAddress() As String
        Get
            Return CollAddress
        End Get
        Set(ByVal Value As String)
            CollAddress = Value
        End Set
    End Property

    ''' <summary>
    ''' Created by taylor on ANDREW at 25/01/2007 07:11:17
    '''     Create reporter email from address
    ''' </summary>
    ''' <returns>
    '''     A System.String value...
    ''' </returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 25/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Public Shared Function ReporterEmailFrom() As String
        Return "OUTPUT_FROM=" & CollectionEmailAddress
    End Function

End Class
#End Region

#Region "Table Field Types"
#Region "Boolean Field"


'''<summary>
''' A field that represents a boolean value within the database
''' </summary>
''' <remarks>A Boolean field is represented in the database by a varchar2(1) field,
''' this usually has the values of 'T' for TRUE and 'F' for FALSE, but it may have other values
''' </remarks>
Public Class BooleanField
    Inherits TableField
    Private blnNVL As Char
    Private chTrueChar As Char
    Private chFalseChar As Char


    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Value to use if Null in the database</param>
    ''' <param name='TrueChar'>Character that indicates a TRUE value defaults to 'T'</param>
    ''' <param name='FalseChar'>Character that indicates a FALSE value defaults to 'F'</param>
    ''' <param name='isIdentity'>Indicates that it forms part of the primary key, defaults to false</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Char, _
    Optional ByVal TrueChar As Char = "T"c, Optional ByVal FalseChar As Char = "F"c, _
    Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, isIdentity)
        chTrueChar = TrueChar
        chFalseChar = FalseChar
    End Sub

    '''<summary>
    ''' Value property for Boolean field 
    ''' </summary>
    ''' <remarks> An override on the base class to give a boolean value field
    ''' </remarks>
    Public Shadows Property Value() As Boolean
        Get
            If TypeOf MyBase.Value Is Boolean Then
                Return CBool(MyBase.Value)
            Else
                Return (CStr(MyBase.Value) = chTrueChar)
            End If
        End Get
        Set(ByVal Value As Boolean)
            MyBase.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Indicates what is being used to indate TRUE
    ''' </summary>
    Public Property TrueChar() As Char
        Get
            Return chTrueChar
        End Get
        Set(ByVal Value As Char)
            chTrueChar = Value
        End Set
    End Property

    '''<summary>
    ''' Indicates what is being used to indate FALSE
    ''' </summary>
    Public Property FalseChar() As Char
        Get
            Return chFalseChar
        End Get
        Set(ByVal Value As Char)
            chFalseChar = Value
        End Set
    End Property

    '''<summary>
    ''' Convert Field to an XML representation of the field
    ''' </summary>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""boolean"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If MyBase.Value Is DBNull.Value Then
        Else
            If Me.Value Then
                strRet += "TRUE"
            Else
                strRet += "FALSE"
            End If
        End If
        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf
        Return strRet

    End Function

    'Overide for dealing with updating Boolean fields 
    '''<summary>
    ''' convert the current value into a form suitable for use in the database
    ''' </summary>
    ''' <param name='Sep'></param>
    Public Overrides Function ValueToDbString(Optional ByVal Sep As Char = "'"c, _
        Optional ByVal ChangeID As Boolean = False) As String
        Dim strTemp As String = "F"
        Dim strRet As String = ""

        If Value Then
            strTemp = Me.chTrueChar
        Else
            strTemp = Me.chFalseChar
        End If

        strRet = Sep & strTemp & Sep

        Return strRet

    End Function

    '''<summary>
    ''' Tells whether field has changed
    ''' </summary>
    ''' <returns>True if field has changed value</returns>
    Public Overrides Function Changed() As Boolean
        If OldValue Is Nothing Then
            Return True
            Exit Function
        End If
        If MyBase.Value Is Nothing Then
            Value = (CStr(NVL) = TrueChar)
        End If

        If TypeOf OldValue Is Char Then OldValue = (CStr(OldValue) = TrueChar)

        Return (CStr(OldValue) <> Value)
    End Function


End Class
#End Region

'''<summary>
''' A field that represents a Sample Manager User in the database
''' </summary>
''' <remarks>
''' This type of field though primarily intended for use in maintenace fields can be used anywhere where the data represents a Personnel.identity field
''' </remarks>
Public Class UserField
    Inherits StringField

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    Public Sub New(ByVal FieldName As String)
        MyBase.New(FieldName, "", 10, False)
    End Sub

    '''<summary>
    ''' Always will return true
    ''' </summary>
    ''' <returns>True if field has changed value</returns>
    Public Overrides Function Changed() As Boolean
        Return False        ' Timestamps will change if the any of the other data is changed
    End Function
End Class

'''<summary>
''' A field that represents a time stamp field in the database
''' </summary>
''' <remarks>
''' This type of field though primarily intended for use in maintenace fields can be used anywhere where the data represents a timestamp<para/>
''' This field will usually have the name Modified_on and will be stored as a date
''' </remarks>
Public Class TimeStampField
    Inherits DateField
    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    Public Sub New(ByVal FieldName As String)
        MyBase.New(FieldName, DateField.ZeroDate, False)
    End Sub

    '''<summary>
    ''' Provides the update information, uses Sysdate
    ''' </summary>
    ''' <returns>Code to se tthe field to sysdate</returns>
    Public Overrides Function UpdateString() As String
        Return Me.FieldName & " = SYSDATE + 0.005 "
    End Function

    '''<summary>
    ''' Always will return true
    ''' </summary>
    ''' <returns>True if field has changed value</returns>
    Public Overrides Function Changed() As Boolean
        Return False         ' Timestamps will change if the any of the other data is changed
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' For timestamp fields we should write sysdate into the field.
    ''' </summary>
    ''' <param name="Sep"></param>
    ''' <param name="ChangeID"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [01/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overrides Function ValueToDbString(Optional ByVal Sep As Char = "'", Optional ByVal ChangeID As Boolean = False) As String
        Return "SYSDATE"
    End Function
End Class

'''<summary>
''' A field that represents a date in the database
''' </summary>
Public Class DateField
    Inherits TableField

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Date, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, isIdentity)
    End Sub

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    Public Sub New(ByVal FieldName As String)
        MyBase.New(FieldName, DateField.ZeroDate, False)
    End Sub

    '''<summary>
    ''' Override of XML function to deal with dates
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""date"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If MyBase.Value Is DBNull.Value Then
        Else
            If Date.Compare(Me.Value, DateField.ZeroDate) = 0 Then
            Else
                strRet += Me.Value.ToString("dd-MMM-yyyy HH:mm")
            End If
        End If
        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf
        Return strRet
    End Function

    '''<summary>
    ''' this date is used to represent a null date
    ''' </summary>
    Public Const DummyDate As Date = #1/1/1850#

    '''<summary>
    ''' this date is used to represent a null date
    ''' </summary>
    ''' <returns>Dummy date</returns>
    Public Shared Function ZeroDate() As Date
        Return DateSerial(1850, 1, 1)
    End Function

    '''<summary>
    ''' the value of this field as a date
    ''' </summary>
    Public Shadows Property Value() As Date
        Get
            If MyBase.Value Is DBNull.Value Then
                Return MyBase.Value
            Else
                Return CType(MyBase.Value, Date)
            End If
        End Get
        Set(ByVal Value As Date)
            MyBase.Value = Value
        End Set
    End Property


  
End Class


'''<summary>
''' represents integer fields in the database i.e type NUMBER or NUMBER(*)
''' </summary>
Public Class IntegerField
    Inherits TableField

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Integer, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, isIdentity)
    End Sub

    '''<summary>
    ''' Override of XML function to deal with dates
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""integer"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If Me.Value Is DBNull.Value Then
        ElseIf Me.Value Is Nothing Then
        Else
            strRet += CStr(CInt(Me.Value))
        End If

        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf

        Return strRet

    End Function


End Class

'''<summary>
''' represents a tableset in the database
''' </summary>
Public Class TableSetField
    Inherits IntegerField
    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Integer, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, isIdentity)
    End Sub

    '''<summary>
    ''' Override of XML function to deal with dates
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower & vbCrLf
        strRet += " type=""integer"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">" & vbCrLf
        Else
            strRet += ">" & vbCrLf
        End If
        If Me.Value Is DBNull.Value Then
            strRet += vbCrLf
        Else
            If (Me.Value = -1) Or (Me.Value = 0) Then
                strRet += vbCrLf
            Else
                strRet += Me.Value & vbCrLf
            End If
        End If

        strRet += "</" & Me.FieldName.ToLower & ">"

        Return strRet

    End Function
End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : ROWIDField
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Handles RowIDs a pseudocolumn in the database
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [04/07/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class ROWIDField
    Inherits StringField

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new Rowid field
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/07/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New("ROWID", "", 20)
        Me.ReadOnlyField = True
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicate whether RowId is set 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/07/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsSet() As Boolean
        Return (CStr(Me.Value).Length > 0)
    End Function
End Class

'''<summary>
''' represents a string VARCHAR2(*) or CHAR
''' </summary>
Public Class StringField
    Inherits TableField

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='Length'>Length of string</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As String, _
        ByVal Length As Integer, Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, Length, isIdentity)
    End Sub

    '''<summary>
    ''' Override of XML function to deal with strings
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date
        Dim intPos As Integer
        Dim strTemp As String

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""string"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If Me.Value Is DBNull.Value Then
        Else
            If Me.Value = Nothing Then
            ElseIf Me.Value.trim.length = 0 Then
            Else
                strTemp = Me.Value
                intPos = InStr(strTemp, "&")
                While intPos <> 0
                    strTemp = Mid(strTemp, 1, intPos) & "amp;" & Mid(strTemp, intPos + 1)
                    intPos = InStr(intPos + 1, strTemp, "&")
                End While
                strRet += strTemp
            End If
        End If

        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf

        Return strRet

    End Function

    '''<summary>
    ''' separate property to return value as a string
    ''' </summary>
    Public Property StringValue() As String
        Get
            Return MyBase.Value
        End Get
        Set(ByVal Value As String)
            MyBase.Value = Value
        End Set
    End Property


End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : StringPhraseField
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' String field linked to a phrase
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [08/02/2006]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class StringPhraseField
    Inherits StringField
    Private myPhrase As String = ""
    Private myPhraseObject As Glossary.Phrase

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create object 
    ''' </summary>
    ''' <param name="FieldName"></param>
    ''' <param name="NVL"></param>
    ''' <param name="Length"></param>
    ''' <param name="Phrase"></param>
    ''' <param name="isIdentity"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal FieldName As String, ByVal NVL As String, _
        ByVal Length As Integer, ByVal Phrase As String, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, Length, isIdentity)

        myPhrase = Phrase
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the phrase used 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Phrase() As String
        Get
            Return Me.myPhrase
        End Get
        Set(ByVal Value As String)
            Me.myPhrase = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the underlying phrase
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[Taylor]</Author><date> [15/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseObject() As Glossary.Phrase
        Get
            If myPhraseObject Is Nothing Then
                myPhraseObject = New Glossary.Phrase(Me.Phrase, Me.Value)
            End If
            Return myPhraseObject
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return the phrase description associated with the phrase 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	27/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property PhraseValue() As String
        Get
            Dim strRet As String = ""
            If Me.myPhrase.Trim.Length > 0 Then
                If Me.StringValue.Trim.Length > 0 Then
                    'Define parameter collection and fill it 
                    Dim pColl As New Collection()
                    pColl.Add(Me.myPhrase)
                    pColl.Add(Me.StringValue)

                    strRet = CConnection.PackageStringList("LIB_UTILS.decodePhrase", pColl)

                End If
            End If
            Return strRet
        End Get

    End Property

    Public Sub SetPhraseByOrderNum(ByVal OrdNum As Integer)
        Dim oCmd As New OleDb.OleDbCommand()
        Dim objReturn As Object
        Dim strPhraseId As String = ""
        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "Select phrase_id from phrase where phrase_type = ? and order_num = ? "
            oCmd.Parameters.Add(CConnection.StringParameter("Type", Me.myPhrase, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("ID", CStr(OrdNum).PadLeft(10), 10))

            If CConnection.ConnOpen Then            'Attempt to open reader
                objReturn = oCmd.ExecuteScalar
                If Not objReturn Is Nothing Then
                    strPhraseId = CStr(objReturn)
                End If
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "StringPhraseField-OrderNumber-4855")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try
        Me.Value = strPhraseId
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get Order number of phrase
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	23/07/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function OrderNumber() As Integer

        Dim oCmd As New OleDb.OleDbCommand()
        Dim objReturn As Object
        Dim intOrdNum As Integer = 0
        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.CommandText = "Select order_num from phrase where phrase_type = ? and phrase_id = ? "
            oCmd.Parameters.Add(CConnection.StringParameter("Type", Me.myPhrase, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("ID", Me.StringValue, 10))

            If CConnection.ConnOpen Then            'Attempt to open reader
                objReturn = oCmd.ExecuteScalar
                If Not objReturn Is Nothing Then
                    intOrdNum = CInt(objReturn)
                End If
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "StringPhraseField-OrderNumber-4855")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try
        Return intOrdNum
    End Function
End Class

Public Class BooleanPhraseField
    Inherits StringPhraseField

    Private blnNVL As Char
    Private chTrueChar As Char
    Private chFalseChar As Char


    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Value to use if Null in the database</param>
    ''' <param name='TrueChar'>Character that indicates a TRUE value defaults to 'T'</param>
    ''' <param name='FalseChar'>Character that indicates a FALSE value defaults to 'F'</param>
    ''' <param name='isIdentity'>Indicates that it forms part of the primary key, defaults to false</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Char, ByVal Phrase As String, _
    Optional ByVal TrueChar As Char = "T", Optional ByVal FalseChar As Char = "F", _
    Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, 1, Phrase, isIdentity)
        chTrueChar = TrueChar
        chFalseChar = FalseChar
    End Sub

    '''<summary>
    ''' Value property for Boolean field 
    ''' </summary>
    ''' <remarks> An override on the base class to give a boolean value field
    ''' </remarks>
    Public Shadows Property Value() As Boolean
        Get
            If TypeOf MyBase.Value Is Boolean Then
                Return MyBase.Value
            Else
                Return (MyBase.Value = chTrueChar)
            End If
        End Get
        Set(ByVal Value As Boolean)
            MyBase.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Indicates what is being used to indate TRUE
    ''' </summary>
    Public Property TrueChar() As Char
        Get
            Return chTrueChar
        End Get
        Set(ByVal Value As Char)
            chTrueChar = Value
        End Set
    End Property

    '''<summary>
    ''' Indicates what is being used to indate FALSE
    ''' </summary>
    Public Property FalseChar() As Char
        Get
            Return chFalseChar
        End Get
        Set(ByVal Value As Char)
            chFalseChar = Value
        End Set
    End Property

    '''<summary>
    ''' Convert Field to an XML representation of the field
    ''' </summary>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""boolean"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If MyBase.Value Is DBNull.Value Then
        Else
            If Me.Value Then
                strRet += "TRUE"
            Else
                strRet += "FALSE"
            End If
        End If
        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf
        Return strRet

    End Function

    'Overide for dealing with updating Boolean fields 
    '''<summary>
    ''' convert the current value into a form suitable for use in the database
    ''' </summary>
    ''' <param name='Sep'></param>
    Public Overrides Function ValueToDbString(Optional ByVal Sep As Char = "'", _
        Optional ByVal ChangeID As Boolean = False) As String
        Dim strTemp As String = "F"
        Dim strRet As String = ""

        If Value Then
            strTemp = Me.chTrueChar
        Else
            strTemp = Me.chFalseChar
        End If

        strRet = Sep & strTemp & Sep

        Return strRet

    End Function

    '''<summary>
    ''' Tells whether field has changed
    ''' </summary>
    ''' <returns>True if field has changed value</returns>
    Public Overrides Function Changed() As Boolean
        If OldValue Is Nothing Then
            Return True
            Exit Function
        End If
        If MyBase.Value Is Nothing Then
            Value = (NVL = TrueChar)
        End If

        If TypeOf OldValue Is Char Then OldValue = (OldValue = TrueChar)

        Return (OldValue <> Value)
    End Function


End Class

'''<summary>
''' specialised form of string field to deal with set based strings
''' </summary>
Public Class StringRangeField
    Inherits StringField
    Private myRange As String = ""
    Private blnError As Boolean = False

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='Length'>Length of string</param>
    ''' <param name='Range'>Range of possible values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As String, _
        ByVal Length As Integer, ByVal Range As String, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, Length, isIdentity)

        myRange = Range
    End Sub

    '''<summary>
    ''' override of value property to do error checking
    ''' </summary>
    Public Shadows Property Value() As String
        Get
            Return MyBase.Value
        End Get
        Set(ByVal Value As String)
            If InStr(myRange, Value) <> 0 Then
                MyBase.Value = Value
                blnError = False
            Else
                MyBase.Value = DBNull.Value
                blnError = True
            End If
        End Set
    End Property

    '''<summary>
    ''' Error check
    ''' </summary>
    ''' <returns>TRUE if error</returns>
    Public Function HasError() As Boolean
        Return blnError
    End Function
End Class

'''<summary>
''' Representation of a real in the database
''' </summary>
Public Class DoubleField
    Inherits TableField

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Double, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.New(FieldName, NVL, isIdentity)
    End Sub

    '''<summary>
    ''' Override of XML function to deal with numbers
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    Public Overrides Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date

        strRet += "<" & Me.FieldName.ToLower
        strRet += " type=""double"""
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        If Me.Value Is DBNull.Value Then
        ElseIf Me.Value Is Nothing Then
        Else
            strRet += CStr(CDbl(Me.Value))
        End If

        strRet += "</" & Me.FieldName.ToLower & ">" & vbCrLf

        Return strRet
    End Function

    '''<summary>
    ''' Override of XSDSchema function to deal with numbers
    ''' </summary>
    ''' <returns>XML representation of field structure</returns>
    Public Overrides Function XSDSchema() As String
        Return "<xs:element name=""" & Me.FieldName & """ type=""xs:decimal""  minOccurs=""0""/>"

    End Function
End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : RangeOption
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' object to handle range options replacing database fields
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	01/06/2007	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class RangeOption
    Private strDefault As String = ""
    Private strValue As String = ""
    Private strOldValue As String = ""
    Private strOption As String = ""
    Private strClient As String = ""
    Private strRowID As String = ""
    Private strType As String

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new range option 
    ''' </summary>
    ''' <param name="OptionName"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal OptionName As String)
        MyBase.new()
        strOption = OptionName

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get default value 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function GetDefault() As Object
        'now get default value
        Dim oCmd As New OleDb.OleDbCommand()
        oCmd.CommandText = "Select OPTION_DEFAULT from options where identity = ?"
        Dim objReturn As Object = Nothing

        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.Parameters.Add(CConnection.StringParameter("Optionid", strOption, 10))
            If CConnection.ConnOpen Then            'Attempt to open reader
                objReturn = oCmd.ExecuteScalar
                strDefault = CStr(objReturn)
                oCmd.CommandText = "Select OPTION_TYPE from options where identity = ?"
                objReturn = oCmd.ExecuteScalar
                strType = CStr(objReturn)
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-New-5042")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try
        'We should have the default value for this option 
        Return objReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Load value from database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function Load() As Boolean
        Dim blnRet As Boolean = True

        'get actual value 
        Dim oColl As New Collection()
        Dim oCmd As New OleDb.OleDbCommand()
        Try  ' Protecting
            GetDefault()
            'Add parameters
            oColl.Add(strClient)
            oColl.Add(strOption)


            oColl.Add(strDefault)


            'Library routine takes care of start and end dates
            Dim strReturn As String = CConnection.PackageStringList("lib_customer.GetOptionDefault", oColl)
            strValue = strReturn.Trim
            strOldValue = strValue
            oCmd.CommandText = "Select rowId from customer_options where customer_id = ? and option_id = ?"
            oCmd.Parameters.Add(CConnection.StringParameter("clientid", strClient, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("optionid", strOption, 10))
            oCmd.Connection = CConnection.DbConnection
            If CConnection.ConnOpen Then
                Dim objReturn As Object = oCmd.ExecuteScalar
                If Not objReturn Is Nothing Then
                    strRowID = CStr(objReturn)
                End If
            End If
        Catch ex As Exception
            blnRet = False
            MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-Load-5084")
        Finally
            CConnection.SetConnClosed()
        End Try
        Return blnRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the value property
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Value() As String
        Get
            Return Me.strValue
        End Get
        Set(ByVal Value As String)
            Me.strValue = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Roll data back to base value 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub RollBack()
        strValue = strOldValue
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get set customer, set operation loads
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Customer() As String
        Get
            Return Me.strClient
        End Get
        Set(ByVal Value As String)
            Me.strClient = Value
            Load()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update value to database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Update() As Boolean

        Dim blnReturn As Boolean

        Dim oCmd As New OleDb.OleDbCommand()
        If Me.strClient.Trim.Length = 0 Then Exit Function
        If strRowID.Trim.Length = 0 Then
            oCmd.CommandText = "Insert into customer_options (option_id,customer_id,OPTION_VALUE_RANGE,Operator) values(?,?,?,?)"
            oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strOption, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("Customer_id", Me.strClient, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("Value", strValue, 100))

            oCmd.Parameters.Add(CConnection.StringParameter("operator", Glossary.Glossary.CurrentSMUser.Identity, 10))
            Try ' Protecting 
                oCmd.Connection = CConnection.DbConnection
                If CConnection.ConnOpen Then            'Attempt to open reader
                    Dim intReturn As Integer = oCmd.ExecuteNonQuery
                    blnReturn = (intReturn = 1)
                    If blnReturn Then
                        Return blnReturn
                        Exit Function
                    End If
                End If
            Catch ex As OleDb.OleDbException
                If InStr(ex.Message, "ORA-00001") <> 0 Then 'Need to update not insert
                    oCmd.CommandText = "update customer_options set option_value_bool = ? where option_id = ? and customer_id = ?"
                    oCmd.Parameters.Clear()
                    oCmd.Parameters.Add(CConnection.StringParameter("Value", strValue, 100))
                    oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strOption, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("Customer_id", Me.strClient, 10))
                    Dim intReturn As Integer = oCmd.ExecuteNonQuery

                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-Update-5139")
            Finally
                CConnection.SetConnClosed()             'Close connection
            End Try
        Else
            Try
                oCmd.CommandText = "update customer_options set option_value_bool = ? where rowid = ? "
                oCmd.Parameters.Clear()

                oCmd.Parameters.Add(CConnection.StringParameter("Value", strValue, 100))
                oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strRowID, 40))
                Dim intReturn As Integer = oCmd.ExecuteNonQuery
            Catch ex As Exception
            End Try
        End If
    End Function

End Class
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : BooleanOption
''' 
''' -----------------------------------------------------------------------------
''' <summary>
'''     a database access object tied to a customer option
''' </summary>
''' <remarks>
''' </remarks>
''' <history>
''' 	[taylor]	01/06/2007	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class BooleanOption
    Private blnDefault As Boolean
    Private blnValue As Boolean
    Private blnOldValue As Boolean
    Private strOption As String = ""
    Private strClient As String = ""
    Private strRowID As String = ""

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create object 
    ''' </summary>
    ''' <param name="OptionName"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal OptionName As String)
        MyBase.new()
        strOption = OptionName

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get the default value for this option
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function GetDefault() As Object
        'now get default value
        Dim oCmd As New OleDb.OleDbCommand()
        oCmd.CommandText = "Select OPTION_DEFAULT from options where identity = ?"
        Dim objReturn As Object = Nothing

        Try ' Protecting 
            oCmd.Connection = CConnection.DbConnection
            oCmd.Parameters.Add(CConnection.StringParameter("Optionid", strOption, 10))
            If CConnection.ConnOpen Then            'Attempt to open reader
                objReturn = oCmd.ExecuteScalar
                If objReturn Is Nothing Then
                    blnDefault = False
                Else
                    blnDefault = (CStr(objReturn).Trim = "T")
                End If

            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-New-5042")
        Finally
            CConnection.SetConnClosed()             'Close connection
        End Try
        'We should have the default value for this option 
        Return objReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get value of option
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Private Function Load() As Boolean
        Dim blnRet As Boolean = True

        'get actual value 
        Dim oColl As New Collection()
        Dim oCmd As New OleDb.OleDbCommand()
        Try  ' Protecting
            GetDefault()
            'Add parameters
            oColl.Add(strClient)
            oColl.Add(strOption)

            If blnDefault Then
                oColl.Add("T")
            Else
                oColl.Add("F")
            End If

            'Library routine takes care of start and end dates
            Dim strReturn As String = CConnection.PackageStringList("lib_customer.GetOptionDefault", oColl)
            blnValue = (strReturn.Trim = "T")
            blnOldValue = blnValue
            oCmd.CommandText = "Select rowId from customer_options where customer_id = ? and option_id = ?"
            oCmd.Parameters.Add(CConnection.StringParameter("clientid", strClient, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("optionid", strOption, 10))
            oCmd.Connection = CConnection.DbConnection
            If CConnection.ConnOpen Then
                Dim objReturn As Object = oCmd.ExecuteScalar
                If Not objReturn Is Nothing Then
                    strRowID = CStr(objReturn)
                End If
            End If
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-Load-5084")
            blnRet = False
        Finally

        End Try
        Return blnRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rollback to base value 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub RollBack()
        Me.blnValue = Me.blnOldValue
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose the value property
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Value() As Boolean
        Get
            Return Me.blnValue
        End Get
        Set(ByVal Value As Boolean)
            Me.blnValue = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get set customer, set operation loads
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	01/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Property Customer() As String
        Get
            Return Me.strClient
        End Get
        Set(ByVal Value As String)
            Me.strClient = Value
            Load()
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Update value to database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	04/06/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function Update() As Boolean

        Dim blnReturn As Boolean

        Dim oCmd As New OleDb.OleDbCommand()

        'see if we have a customer id
        If Me.strClient.Trim.Length = 0 Then Exit Function
        If strRowID.Trim.Length = 0 Then
            oCmd.CommandText = "Insert into customer_options (option_id,customer_id,OPTION_VALUE_BOOL,Operator) values(?,?,?,?)"
            oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strOption, 10))
            oCmd.Parameters.Add(CConnection.StringParameter("Customer_id", Me.strClient, 10))
            If blnValue Then
                oCmd.Parameters.Add(CConnection.StringParameter("Value", "T", 1))
            Else
                oCmd.Parameters.Add(CConnection.StringParameter("Value", "F", 1))
            End If
            oCmd.Parameters.Add(CConnection.StringParameter("operator", Glossary.Glossary.CurrentSMUser.Identity, 10))
            Try ' Protecting 
                oCmd.Connection = CConnection.DbConnection
                If CConnection.ConnOpen Then            'Attempt to open reader
                    Dim intReturn As Integer = oCmd.ExecuteNonQuery
                    blnReturn = (intReturn = 1)
                    If blnReturn Then
                        Return blnReturn
                        Exit Function
                    End If
                End If
            Catch ex As OleDb.OleDbException
                If InStr(ex.Message, "ORA-00001") <> 0 Then 'Need to update not insert
                    oCmd.CommandText = "update customer_options set option_value_bool = ? where option_id = ? and customer_id = ?"
                    oCmd.Parameters.Clear()
                    If blnValue Then
                        oCmd.Parameters.Add(CConnection.StringParameter("Value", "T", 1))
                    Else
                        oCmd.Parameters.Add(CConnection.StringParameter("Value", "F", 1))
                    End If
                    oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strOption, 10))
                    oCmd.Parameters.Add(CConnection.StringParameter("Customer_id", Me.strClient, 10))
                    Dim intReturn As Integer = oCmd.ExecuteNonQuery

                End If

            Catch ex As Exception
                MedscreenLib.Medscreen.LogError(ex, , "BooleanOption-Update-5139")
            Finally
                CConnection.SetConnClosed()             'Close connection
            End Try
        Else
            Try
                oCmd.CommandText = "update customer_options set option_value_bool = ? where rowid = ? "
                oCmd.Parameters.Clear()
                If blnValue Then
                    oCmd.Parameters.Add(CConnection.StringParameter("Value", "T", 1))
                Else
                    oCmd.Parameters.Add(CConnection.StringParameter("Value", "F", 1))
                End If
                oCmd.Parameters.Add(CConnection.StringParameter("Optid", Me.strRowID, 40))
                Dim intReturn As Integer = oCmd.ExecuteNonQuery
            Catch ex As Exception
            End Try
        End If
    End Function
End Class

'''<summary>
''' A field that value within the database
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Read only property Added</Action></revision>
''' </revisionHistory>

Public Class TableField
    Private blnIdentity As Boolean
    Private strFieldName As String
    Private CurrentValue As Object
    Private OldVValue As Object
    Private blnChanged As Boolean
    Private intFieldLen As Integer = -1
    Private objNVL As Object
    Private blnNull As Boolean = True
    Private blnReadOnly As Boolean = False

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether this field is read only or not
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReadOnlyField() As Boolean
        Get
            Return Me.blnReadOnly
        End Get
        Set(ByVal Value As Boolean)

            Me.blnReadOnly = Value
        End Set
    End Property

    '''<summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()


    End Sub

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Object, Optional ByVal isIdentity As Boolean = False)
        MyBase.new()
        strFieldName = FieldName
        blnIdentity = isIdentity
        objNVL = NVL

    End Sub

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='FieldLen'>The defined length of the field (Strings Only)</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, ByVal NVL As Object, ByVal FieldLen As Integer, _
        Optional ByVal isIdentity As Boolean = False)
        MyBase.new()
        strFieldName = FieldName
        intFieldLen = FieldLen
        blnIdentity = isIdentity
        objNVL = NVL
    End Sub

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='FieldLen'>The defined length of the field (Strings Only)</param>
    ''' <param name='value'>The value of the field</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, _
        ByVal NVL As Object, _
        ByVal FieldLen As Integer, ByVal value As Object, _
           Optional ByVal isIdentity As Boolean = False)
        MyBase.new()
        strFieldName = FieldName
        intFieldLen = FieldLen
        CurrentValue = value
        blnIdentity = isIdentity
        objNVL = NVL

    End Sub

    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='FieldName'>Name of the field in the database</param>
    ''' <param name='NVL'>Representation of NULL values</param>
    ''' <param name='value'>The value of the field</param>
    ''' <param name='isIdentity'>Indicates that the field is part of the primary key</param>
    Public Sub New(ByVal FieldName As String, _
        ByVal NVL As Object, ByVal value As Object, _
          Optional ByVal isIdentity As Boolean = False)
        MyBase.new()
        strFieldName = FieldName
        blnIdentity = isIdentity
        CurrentValue = value
        objNVL = NVL

    End Sub

    '''<summary>
    ''' NVL value
    ''' </summary>
    Public Property NVL() As Object
        Get
            Return objNVL
        End Get
        Set(ByVal Value As Object)
            objNVL = Value
        End Set
    End Property


    Private myTableName As String
    ''' <summary>
    ''' Name of the table field is contained in
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Property TableName() As String
        Get
            Return myTableName
        End Get
        Set(ByVal value As String)
            myTableName = value
        End Set
    End Property



    '''<summary>
    ''' Set the field to null
    ''' </summary>
    Public Sub SetNull()
        CurrentValue = System.DBNull.Value
    End Sub

    '''<summary>
    ''' String to update the field in the database
    ''' </summary>
    ''' <returns>String that will update this field</returns>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Added support for readonly fields</Action></revision>
    ''' </revisionHistory>
    Public Overridable Function UpdateString() As String
        If Me.ReadOnlyField Then
            Return ""
        Else
            Return Me.FieldName & " = " & Me.ValueToDbString
        End If
    End Function

    '''<summary>
    ''' The value of this field
    ''' </summary>
    Public Overridable Property Value() As Object
        Get
            If CurrentValue Is Nothing OrElse CurrentValue Is System.DBNull.Value Then
                Return Me.NVL
            Else
                If TypeOf CurrentValue Is String Then   'If it is a string trim it to get rid of SM Mess
                    Return CStr(CurrentValue).Trim
                Else
                    Return CurrentValue
                End If

            End If

        End Get
        Set(ByVal Value As Object)
            CurrentValue = Value
            Me.blnNull = False
        End Set
    End Property

    '''<summary>
    ''' Tells whether field has changed
    ''' </summary>
    ''' <returns>True if field has changed value</returns>
    Public Overridable Function Changed() As Boolean
        'If Me.IsIdentity Then 'Identity fields will never change after creation
        'Return False
        'Else
        If OldValue Is Nothing Then
            If TypeOf OldValue Is Boolean Then
                OldValue = False
            Else
                OldValue = NVL
            End If
        End If
        If CurrentValue Is Nothing Then CurrentValue = NVL
        If TypeOf CurrentValue Is String Then
            Return (CType(OldVValue, String).Trim <> CType(CurrentValue, String).Trim)
        ElseIf CurrentValue Is System.DBNull.Value Then
            Return Not (OldValue Is System.DBNull.Value)
        Else
            Try
                Return (OldVValue <> CurrentValue)
            Catch
                Return True
            End Try
        End If
        'End If
    End Function

    '''<summary>
    ''' The original value read from the database
    ''' </summary>
    Public Property OldValue() As Object
        Get
            Return OldVValue
        End Get
        Set(ByVal Value As Object)
            OldVValue = Value
        End Set
    End Property

    '''<summary>
    ''' Is the field part of the primary key
    ''' </summary>
    Public Property IsIdentity() As Boolean
        Get
            Return blnIdentity
        End Get
        Set(ByVal Value As Boolean)
            blnIdentity = Value
        End Set
    End Property

    Public ReadOnly Property IsNull() As Boolean
        Get
            Return blnNull
        End Get
    End Property


    Public WriteOnly Property SetIsNull() As Boolean
        Set(ByVal Value As Boolean)
            blnNull = Value
            If blnNull Then SetNull()
        End Set
    End Property
    '''<summary>
    ''' Name of field in the database
    ''' </summary>
    Public Property FieldName() As String
        Get
            Return strFieldName
        End Get
        Set(ByVal Value As String)
            strFieldName = Value
        End Set
    End Property

    '''<summary>
    ''' Length of field
    ''' </summary>
    Public Property FieldLength() As Integer
        Get
            Return intFieldLen
        End Get
        Set(ByVal Value As Integer)
            intFieldLen = Value
        End Set
    End Property



    '''<summary>
    ''' Convert Value to a string that can be stored in the database
    ''' </summary>
    ''' <param name='Sep'>Separator charcter default is '</param>
    ''' <param name='ChangeID'>Allow the ID field to be changed</param>
    ''' <returns>Formatte dstring</returns>
    Public Overridable Function ValueToDbString(Optional ByVal Sep As Char = "'", _
    Optional ByVal ChangeID As Boolean = False) As String
        Dim strTemp As String = ""
        Dim strRet As String = ""
        Dim datTemp As Date
        Dim ValToChange As Object

        If ChangeID And Me.IsIdentity Then
            ValToChange = OldValue
        ElseIf CurrentValue Is System.DBNull.Value Then
            ValToChange = ""
        Else
            ValToChange = CurrentValue
        End If


        If TypeOf ValToChange Is String Or TypeOf ValToChange Is Char Then
            strTemp = ValToChange
            If strTemp.Length > intFieldLen AndAlso intFieldLen > 0 Then
                strTemp = Mid(strTemp, 1, intFieldLen)
            End If
            If strTemp.Trim.Length = 0 Then
                strRet = "NULL"
            Else
                strTemp = Medscreen.FixQuotes(strTemp)
                strRet = Sep & strTemp & Sep
            End If
        ElseIf TypeOf ValToChange Is Date Then
            datTemp = ValToChange
            If datTemp = DateField.ZeroDate Then        'If the date value is zero date then return a null
                strRet = "NULL"
            Else
                If Sep = "'" Then
                    strRet = "TO_DATE('" & datTemp.ToString("ddMMyyyy HHmm") & "','DDMMYYYY HH24mi')"
                Else
                    strRet = datTemp.ToString("dd-MMM-yyyy HH:mm")
                End If
            End If
        ElseIf TypeOf ValToChange Is Boolean Then 'Should never get here should be hit by override
            strTemp = "F"
            If CType(ValToChange, Boolean) Then strTemp = "T"
            strRet = Sep & strTemp & Sep
        Else
            If ValToChange Is System.DBNull.Value Then
                strRet = "NULL"
            ElseIf ValToChange Is Nothing Then
                strRet = "NULL"
            Else
                strRet = ValToChange
            End If
        End If
        Return strRet

    End Function

    '''<summary>
    ''' Provide XSD information
    ''' </summary>
    ''' <returns>Schema info</returns>
    Public Overridable Function XSDSchema() As String
        Return "<xs:element name=""" & Me.FieldName & """ type=""xs:string""  minOccurs=""0""/>"
    End Function

    '''<summary>
    ''' return value as XML
    ''' </summary>
    ''' <returns>XML form of the contents of field</returns>
    Public Overridable Function XMLValue() As String
        'Dim intPos As Integer
        Dim strRet As String = ""
        Dim strTemp As String

        If Me.Value Is DBNull.Value Then
        ElseIf Me.Value Is Nothing Then
        Else
            If TypeOf CurrentValue Is String Then
                If Me.Value.trim.length = 0 Then
                Else
                    strTemp = Medscreen.FixXML(Me.Value)
                    'intPos = InStr(strTemp, "&")
                    'While intPos <> 0
                    '    strTemp = Mid(strTemp, 1, intPos) & "amp;" & Mid(strTemp, intPos + 1)
                    '    intPos = InStr(intPos + 1, strTemp, "&")
                    'End While
                    Return strTemp.Trim
                End If
            ElseIf TypeOf CurrentValue Is Date Then
                If CurrentValue > DateField.ZeroDate Then
                    strRet += CType(Me.Value, Date).ToString("dd-MMM-yyyy HH:mm")
                End If
            ElseIf TypeOf CurrentValue Is Boolean Then
                If CType(Me.Value, Boolean) Then
                    strRet += "TRUE"
                Else
                    strRet += "FALSE"
                End If
            ElseIf (TypeOf CurrentValue Is Integer) Or _
                (TypeOf CurrentValue Is Decimal) Then
                strRet += CStr(Me.Value)
            ElseIf TypeOf CurrentValue Is Double Then
                strRet += CStr(Me.Value)
            End If
        End If

        Return strRet
    End Function

    '''<summary>
    ''' return value as CSV
    ''' </summary>
    ''' <returns>CSV form of the contents of field</returns>
    Public Overridable Function CSVValue() As String
        'Dim intPos As Integer
        Dim strRet As String = ""
        Dim strTemp As String

        If Me.Value Is DBNull.Value Then
        ElseIf Me.Value Is Nothing Then
        Else
            If TypeOf CurrentValue Is String Then
                If Me.Value.trim.length = 0 Then
                Else
                    strTemp = Me.Value
                    Return Chr(34) & strTemp.Trim & Chr(34)
                End If
            ElseIf TypeOf CurrentValue Is Date Then
                If CurrentValue > DateField.ZeroDate Then
                    strRet += Chr(34) & CType(Me.Value, Date).ToString("dd-MMM-yyyy HH:mm") & Chr(34)
                End If
            ElseIf TypeOf CurrentValue Is Boolean Then
                If CType(Me.Value, Boolean) Then
                    strRet += "TRUE"
                Else
                    strRet += "FALSE"
                End If
            ElseIf (TypeOf CurrentValue Is Integer) Or _
                (TypeOf CurrentValue Is Decimal) Then
                strRet += CStr(Me.Value)
            ElseIf TypeOf CurrentValue Is Double Then
                strRet += CStr(Me.Value)
            End If
        End If

        Return strRet
    End Function


    '''<summary>
    ''' return value as CSV
    ''' </summary>
    ''' <returns>CSV form of the contents of field</returns>
    Public Overridable Function ToCSV() As String
        Dim strRet As String = ""
        'Dim datTemp As Date
        'Dim intPos As Integer

        strRet = Me.CSVValue

        Return strRet

    End Function


    '''<summary>
    ''' return value as XML
    ''' </summary>
    ''' <returns>XML form of the contents of field</returns>
    Public Overridable Function ToXML() As String
        Dim strRet As String = ""
        'Dim datTemp As Date
        'Dim intPos As Integer

        strRet = "<" & Me.FieldName.ToLower
        'ADDRESS_TYPE" type="xs:string" />
        If TypeOf CurrentValue Is String Then
            strRet += " type=""string"""
        ElseIf TypeOf CurrentValue Is Date Then
            strRet += " type=""date"""
        ElseIf TypeOf CurrentValue Is Boolean Then
            strRet += " type=""boolean"""
        ElseIf TypeOf CurrentValue Is Integer Then
            strRet += " type=""integer"""
        ElseIf TypeOf CurrentValue Is Double Then
            strRet += " type=""double"""
        End If
        If Not Me.IsIdentity Then
            strRet += " minOccurs=""0""" & ">"
        Else
            strRet += ">"
        End If
        strRet += Me.XMLValue
        strRet += "</" & Me.FieldName.ToLower & ">"
        Return strRet

    End Function

    '''<summary>
    ''' return value as XML
    ''' </summary>
    ''' <param name='Highlight'>Hightlight this field using string</param>
    ''' <returns>XML form of the contents of field</returns>
    Public Overridable Function ToXMLSchema(Optional ByVal HighLight As String = "") As String
        Dim strRet As String = ""
        'Dim datTemp As Date
        'Dim intPos As Integer

        strRet = "<" & Me.FieldName.ToLower & ">"
        'ADDRESS_TYPE" type="xs:string" />
        strRet += Medscreen.FixGreaterThen(Medscreen.FixLessThan(Me.XMLValue)) & HighLight
        strRet += "</" & Me.FieldName.ToLower & ">"
        Return strRet

    End Function



End Class

'''<summary>
''' specialised set of fields to match basic address fields in tables
''' </summary>
Public Class AddressFields
    Inherits TableFields
    Private objAddrLine1 As StringField
    Private objAddrLine2 As StringField
    Private objAddrLine3 As StringField
    Private objCity As StringField
    Private objDistrict As StringField
    Private objPostCode As StringField
    Private objPhone As StringField
    Private objFax As StringField
    Private objCountry As StringField

    Private strWhere As String = ""
    Private intDefaulrFieldLen As Integer = 20

    '''<summary>
    ''' constructor
    ''' </summary>
    ''' <param name='strTableName'>Table in which these fields can be found</param>
    ''' <param name='defaultFieldLen'>normal length of fields</param>
    ''' <param name='Addr1FieldName'>Line1 field name</param>
    ''' <param name='Addr2FieldName'>Line2 field name</param>
    ''' <param name='Addr3FieldName'>Line3 field name</param>
    ''' <param name='AddrCityName'>City field name</param>
    ''' <param name='addrDistrictName'>County field name</param>
    ''' <param name='adrPostcodeName'>Postcode field name</param>
    ''' <param name='addrCountry'>Country Field Name defaults to ""</param>
    ''' <param name='addrPhone'>Phone field Name defaults to ""</param>
    ''' <param name='addrFax'>Fax field defaults to ""</param>
    ''' <remarks>If a fieldname is a null string it is not added</remarks>
    Public Sub New(ByVal strTableName As String, ByVal defaultFieldLen As Integer, _
        ByVal Addr1FieldName As String, ByVal Addr2FieldName As String, _
        ByVal Addr3FieldName As String, ByVal AddrCityName As String, _
        ByVal addrDistrictName As String, ByVal adrPostcodeName As String, _
        Optional ByVal addrCountry As String = "", _
        Optional ByVal addrPhone As String = "", Optional ByVal AddrFax As String = "")

        MyBase.New(strTableName)
        intDefaulrFieldLen = defaultFieldLen
        Me.objAddrLine1 = New StringField(Addr1FieldName, "", defaultFieldLen)
        Me.Add(Me.objAddrLine1)
        Me.objAddrLine2 = New StringField(Addr2FieldName, "", defaultFieldLen)
        Me.Add(Me.objAddrLine2)
        Me.objAddrLine3 = New StringField(Addr3FieldName, "", defaultFieldLen)
        Me.Add(Me.objAddrLine3)
        Me.objCity = New StringField(AddrCityName, "", defaultFieldLen)
        Me.Add(Me.objCity)
        Me.objDistrict = New StringField(addrDistrictName, "", defaultFieldLen)
        Me.Add(Me.objDistrict)
        Me.objPostCode = New StringField(adrPostcodeName, "", defaultFieldLen)
        Me.Add(Me.objPostCode)

        'Optionals 
        If addrCountry.Length > 0 Then
            Me.objCountry = New StringField(addrCountry, "", defaultFieldLen)
            Me.Add(objCountry)
        End If
        If addrPhone.Length > 0 Then
            Me.objPhone = New StringField(addrPhone, "", defaultFieldLen)
            Me.Add(objPhone)
        End If
        If AddrFax.Length > 0 Then
            Me.objFax = New StringField(AddrFax, "", defaultFieldLen)
            Me.Add(objFax)
        End If
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/12/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Clone() As AddressFields
        Dim NewFields As New AddressFields(MyBase.Table, intDefaulrFieldLen, objAddrLine1.FieldName, objAddrLine2.FieldName, objAddrLine3.FieldName, objCity.FieldName, _
                                           objDistrict.FieldName, objPostCode.FieldName)


        Return NewFields
    End Function

    '''<summary>
    ''' Override of XML function to deal with addresses
    ''' </summary>
    ''' <returns>XML representation of field value</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [07/12/2006]</date><Action>fixed bug in xml production</Action></revision>
    ''' </revisionHistory>
    Public Overloads Function ToXML() As String
        Dim strRet As String
        strRet = "<" & Me.XMLHeader & ">" & vbCrLf
        If Me.objAddrLine1.XMLValue Is Nothing Then
            strRet += "<line1/>"
        ElseIf objAddrLine1.XMLValue.Trim.Length = 0 Then
            strRet += "<line1/>"
        Else
            strRet += "<line1>" & Medscreen.Capitalise(Me.objAddrLine1.XMLValue) & "</line1>"
        End If
        If Me.objAddrLine2.XMLValue Is Nothing Then
            strRet += "<line2/>"
        ElseIf objAddrLine2.XMLValue.Trim.Length = 0 Then
            strRet += "<line2/>"
        Else
            strRet += "<line2>" & Medscreen.Capitalise(Me.objAddrLine2.XMLValue) & "</line2>"
        End If
        If Me.objAddrLine3.XMLValue Is Nothing Then
            strRet += "<line3/>"
        ElseIf objAddrLine3.XMLValue.Trim.Length = 0 Then
            strRet += "<line3/>"
        Else
            strRet += "<line3>" & Medscreen.Capitalise(Me.objAddrLine3.XMLValue) & "</line3>"
        End If

        If Me.objCity.XMLValue Is Nothing Then
            strRet += "<city/>"
        ElseIf objCity.XMLValue.Trim.Length = 0 Then
            strRet += "<city/>"
        Else
            strRet += "<city>" & Medscreen.Capitalise(Me.objCity.XMLValue) & "</city>"
        End If
        If Me.objDistrict.XMLValue Is Nothing Then
            strRet += "<district/>"
        ElseIf objDistrict.XMLValue.Trim.Length = 0 Then
            strRet += "<district/>"
        Else
            strRet += "<district>" & Me.objDistrict.XMLValue & "</district>"
        End If
        If Me.objPostCode.XMLValue Is Nothing Then
            strRet += "<postcode/>"
        ElseIf objPostCode.XMLValue.Trim.Length = 0 Then
            strRet += "<postcode/>"
        Else
            strRet += "<postcode>" & Me.objPostCode.XMLValue.ToUpper & "</postcode>"
        End If
        If Not Me.objCountry Is Nothing Then
            If Me.objCountry.XMLValue Is Nothing Then
                strRet += "<country/>"
            ElseIf objCountry.XMLValue.Trim.Length = 0 Then
                strRet += "<country/>"
            Else
                strRet += "<country>" & Me.objCountry.XMLValue & "</country>"
            End If
        End If
        If Not Me.objPhone Is Nothing Then
            If Me.objPhone.XMLValue Is Nothing Then
                strRet += "<phone/>"
            ElseIf objPhone.XMLValue.Trim.Length = 0 Then
                strRet += "<phone/>"
            Else
                strRet += "<phone>" & Me.objPhone.XMLValue & "</phone>"
            End If
        End If
        If Not Me.objFax Is Nothing Then
            If Me.objFax.XMLValue Is Nothing Then
                strRet += "<fax/>"
            ElseIf objFax.XMLValue.Trim.Length = 0 Then
                strRet += "<fax/>"
            Else
                strRet += "<fax>" & Me.objFax.XMLValue & "</fax>"
            End If
        End If

        strRet += "</" & Me.XMLHeader & ">" & vbCrLf
        Return strRet
    End Function

    '''<summary>
    ''' Override of Update function to deal with addresses
    ''' </summary>
    ''' <param name='Conn'>Oledb data connector</param>
    ''' <param name='WhereString'>Where clause</param>
    ''' <returns>TRUE if succesful</returns>
    Public Overloads Function Update(ByVal Conn As OleDb.OleDbConnection, ByVal WhereString As String) As Boolean
        Dim blnRet As Boolean = False
        strWhere = WhereString
        Dim ocmd As New OleDb.OleDbCommand()
        Try

            If MyBase.UpdateString.Trim.Length > 0 Then
                ocmd.CommandText = MyBase.UpdateString & " " & strWhere
                ocmd.Connection = Conn
                Conn.Open()
                Dim iRet As Integer = ocmd.ExecuteNonQuery
                blnRet = (iRet = 1)
            End If
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            Conn.Close()
        End Try
    End Function

    '''<summary>
    ''' First line of address
    ''' </summary>
    Public Property AddLine1() As StringField
        Get
            Return Me.objAddrLine1
        End Get
        Set(ByVal Value As StringField)
            objAddrLine1 = Value
        End Set
    End Property

    '''<summary>
    ''' Second line of address
    ''' </summary>
    Public Property AddLine2() As StringField
        Get
            Return Me.objAddrLine2
        End Get
        Set(ByVal Value As StringField)
            objAddrLine2 = Value
        End Set
    End Property

    '''<summary>
    '''Third line of address
    ''' </summary>
    Public Property AddLine3() As StringField
        Get
            Return Me.objAddrLine3
        End Get
        Set(ByVal Value As StringField)
            objAddrLine3 = Value
        End Set
    End Property

    '''<summary>
    ''' City of address
    ''' </summary>
    Public Property AddrCity() As StringField
        Get
            Return Me.objCity
        End Get
        Set(ByVal Value As StringField)
            Me.objCity = Value
        End Set
    End Property
    Public Property AddrDistrict() As StringField
        Get
            Return Me.objDistrict
        End Get
        Set(ByVal Value As StringField)
            Me.objDistrict = Value
        End Set
    End Property

    '''<summary>
    ''' Postcode
    ''' </summary>
    Public Property AddrPostcode() As StringField
        Get
            Return Me.objPostCode
        End Get
        Set(ByVal Value As StringField)
            Me.objPostCode = Value
        End Set
    End Property

    '''<summary>
    ''' Phone
    ''' </summary>
    Public Property AddrPhone() As StringField
        Get
            Return Me.objPhone
        End Get
        Set(ByVal Value As StringField)
            objPhone = Value
        End Set
    End Property

    '''<summary>
    ''' Country
    ''' </summary>
    Public Property AddrCountry() As StringField
        Get
            Return Me.objCountry
        End Get
        Set(ByVal Value As StringField)
            objCountry = Value
        End Set
    End Property

    '''<summary>
    ''' Fax
    ''' </summary>
    Public Property AddrFax() As StringField
        Get
            Return Me.objFax
        End Get
        Set(ByVal Value As StringField)
            objFax = Value
        End Set
    End Property
    'Public Overrides Function SelectString(Optional ByVal AddTable As Boolean = True, Optional ByVal AddSelect As Boolean = True) As String
    '    Return MyBase.SelectString(False, False)
    'End Function
End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : TableFields
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Representation of a table within the database
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Readonly property Added</Action></revision>
''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action>Timestamp properties added</Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class TableFields
    Inherits CollectionBase
    Private strTableName As String
    Private strHeader As String
    Private blnLoaded As Boolean = False
    Private datLoaded As Date = #1/1/1850#
    Private datUpdated As Date = #1/1/1850#
    Private datAccessed As Date = #1/1/1850#
    Private blnReadonly As Boolean = False
    Private objRowID As ROWIDField = New ROWIDField()

    '''<summary>
    ''' The XMLHeader value
    ''' </summary>
    Public Property XMLHeader() As String
        Get
            Return strHeader
        End Get
        Set(ByVal Value As String)
            strHeader = Value
        End Set
    End Property

    '''<summary>
    ''' Indicates whether the data have changed or not
    ''' </summary>
    ''' <returns>TRUE if the data have  changed</returns>
    Public Function Changed() As Boolean
        Dim blnChanged As Boolean = False
        Dim i As Integer
        Dim objTf As TableField

        For i = 0 To Me.Count - 1
            objTf = Me.Item(i)
            If objTf.Changed Then
                blnChanged = True
                Exit For
            End If
        Next

        Return blnChanged
    End Function

    Public Sub CopyTo(ByVal InFields As TableFields)
        Dim objSourceTf As TableField
        Dim objDestTf As TableField
        For Each objSourceTf In InFields
            If (Not objSourceTf.IsIdentity) And (Not TypeOf objSourceTf Is ROWIDField) Then
                objDestTf = Me.Item(objSourceTf.FieldName)
                objDestTf.Value = objSourceTf.Value
            End If
        Next
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Get a datarow from the table by its row id
    ''' </summary>
    ''' <param name="Row_ID"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	04/11/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Function LoadByRowID(ByVal Row_ID As String) As Boolean
        'Dim intRes As Integer
        Dim Dr As OleDb.OleDbDataReader = Nothing
        Dim ocmd As New OleDb.OleDbCommand()
        Dim blnRet As Boolean = False
        Try
            ocmd.Connection = CConnection.DbConnection
            ocmd.CommandText = Me.FullRowSelect & " where rowid = ?"
            ocmd.Parameters.Add(CConnection.StringParameter("RowID", Row_ID, 60))
            If CConnection.ConnOpen Then
                Dr = ocmd.ExecuteReader
                If Dr.Read Then 'Read the data and populate the field collection
                    readfields(Dr)
                    blnRet = True
                    Me.Loaded = True
                    Me.RowID = Row_ID
                Else
                    blnRet = False
                    Me.Loaded = False
                End If
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , ocmd.CommandText) 'Log any errors
            blnRet = False
        Finally
            'oledconn.Close() don't close it is too expensive opening a connection
            If Not Dr Is Nothing Then
                If Not Dr.IsClosed Then Dr.Close()
                Dr = Nothing
            End If
            CConnection.SetConnClosed()
            ocmd = Nothing

        End Try
        Return blnRet
    End Function

    '''<summary>
    ''' load rows into the table
    ''' </summary>
    ''' <param name='Conn'>Oledb Data connector</param>
    ''' <param name='WhereClause'>Optional where cluase restricting data </param>
    ''' <returns>TRUE if the data have been loaded</returns>
    Public Overloads Function Load(ByVal Conn As OleDb.OleDbConnection, _
    Optional ByVal WhereClause As String = "", _
    Optional ByVal ParameterCollection As Collection = Nothing, _
    Optional ByVal FullTable As Boolean = False) As Boolean
        'Dim intRes As Integer
        Dim Dr As OleDb.OleDbDataReader = Nothing
        Dim ocmd As New OleDb.OleDbCommand()
        Dim blnRet As Boolean = False

        Try
            ocmd.Connection = Conn 'Get current connection and ensure it's open 
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            'Set the command object with the correct select string 
            If WhereClause.Trim.Length > 0 Then
                If FullTable Then
                    ocmd.CommandText = Me.FullRowSelect & " where " & WhereClause
                Else
                    ocmd.CommandText = Me.SelectString & " where " & WhereClause

                End If
                If Not ParameterCollection Is Nothing Then      'Do we have parameters?
                    Dim k As Integer
                    For k = 1 To ParameterCollection.Count      'If so go through and add them 
                        ocmd.Parameters.Add(ParameterCollection.Item(k))
                    Next
                End If
            Else
                If FullTable Then
                    ocmd.CommandText = Me.FullRowSelect
                Else
                    ocmd.CommandText = Me.SelectString
                End If
            End If
            CConnection.SetConnOpen()
            Dr = ocmd.ExecuteReader
            If Dr.Read Then 'Read the data and populate the field collection
                readfields(Dr)
                blnRet = True
                Me.Loaded = True
            Else
                blnRet = False
                Me.Loaded = False
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , ocmd.CommandText) 'Log any errors
            blnRet = False
        Finally
            'oledconn.Close() don't close it is too expensive opening a connection
            If Not Dr Is Nothing Then
                If Not Dr.IsClosed Then Dr.Close()
                Dr = Nothing
            End If
            Conn.Close()
            ocmd = Nothing

        End Try
        Return blnRet
    End Function

    '''<summary>
    ''' Delete rows from the table
    ''' </summary>
    ''' <param name='Conn'>Oledb Data connector</param>
    ''' <returns>TRUE if the data have been changed</returns>
    Public Overloads Function Delete(ByVal Conn As OleDb.OleDbConnection) As Boolean
        Dim intRes As Integer
        Dim oCmd As New OleDb.OleDbCommand()

        Try
            oCmd.Connection = Conn
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            If Me.RowID.Trim.Length > 0 Then
                oCmd.CommandText = "Delete from " & Me.Table & " where rowid = '" & Me.RowID & "'"
            Else
                Me.GetRowId()
                oCmd.CommandText = "Delete from " & Me.Table & " where rowid = '" & Me.RowID & "'"
            End If
            Debug.WriteLine(oCmd.CommandText)
            intRes = oCmd.ExecuteNonQuery
            'Me.Commit()
            Return (intRes = 1)
        Catch ex As Exception
            Medscreen.LogError(ex)
            Conn = Nothing
            Return False
        Finally
            oCmd = Nothing
            'oledconn.Close()
        End Try

    End Function

    '''<summary>
    ''' Insert rows into the table
    ''' </summary>
    ''' <param name='Conn'>Oledb Data connector</param>
    ''' <returns>TRUE if the data have been changed</returns>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Read only checks added</Action></revision>
    ''' </revisionHistory>
    Public Overloads Function Insert(ByVal Conn As OleDb.OleDbConnection) As Boolean
        Dim intRes As Integer
        Dim oCmd As New OleDb.OleDbCommand()

        If Me.ReadOnlyTable Then                        'Check to see if read only 
            Return False
            Exit Function
        End If

        Try
            oCmd.Connection = Conn
            oCmd.CommandText = Me.InsertString
            If CConnection.ConnOpen Then


                intRes = oCmd.ExecuteNonQuery
                'Me.Commit()
                Me.LastModifyDate = Now
                Me.LastAccessDate = Now
                Me.Loaded = (intRes = 1)
                If Loaded Then
                    Me.GetRowId()

                End If
            End If

            Return (intRes = 1)
        Catch ex As OleDb.OleDbException
            Dim strError As String = ex.Message
            Dim intPos As Integer = InStr(strError, "ORA-00001")
            If intPos = 0 Then
                Medscreen.LogError(ex, , "Into " & Me.Table & " - " & Me.InsertString)
            End If
        Catch ex As Exception
            Dim strError As String = ex.Message
            Dim intPos As Integer = InStr(strError, "ORA-00001")
            If intPos = 0 Then
                Medscreen.LogError(ex, , "Into " & Me.Table & " - " & Me.InsertString)
            End If
            Conn = Nothing
            Return False
        Finally
            oCmd = Nothing
            'oledconn.Close()
        End Try


    End Function

    '''<summary>
    ''' Update rows in the table
    ''' </summary>
    ''' <param name='Conn'>Oledb Data connector</param>
    ''' <param name='ChangeID'>If TRUE ID values can be changed</param>
    ''' <returns>TRUE if the data have been changed</returns>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Read only checks added</Action></revision>
    ''' </revisionHistory>
    Public Overridable Overloads Function Update(ByVal Conn As OleDb.OleDbConnection, _
    Optional ByVal ChangeID As Boolean = False, Optional ByVal Force As Boolean = False) As Boolean
        Dim intRes As Integer
        Dim oCmd As New OleDb.OleDbCommand()

        'If nothing has been updated then skip straight out 
        Me.LastAccessDate = Now

        If Not Force And (Not Me.Changed Or Me.ReadOnlyTable) Then  'Or the table is set to be readonly

            Return True         'Though no update done it is a success
            Exit Function
        End If

        Try
            oCmd.Connection = Conn  'set up connection
            Dim changed As Boolean = False
            'Open it if necessary
            oCmd.CommandText = Me.UpdateString(oCmd, changed, ChangeID)          'Set up Update string
            If CConnection.ConnOpen Then
                If Me.Changed Or Force Then intRes = oCmd.ExecuteNonQuery '         'Do update
                If intRes = 1 Then                                      'If one row changed 
                    Me.Commit()                                         'do a commit
                    Me.Loaded = True 'Ensure loaded flag is set
                End If
            End If
            Me.LastModifyDate = Now
            Return (intRes = 1)
        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "ORA-00001") <> 0 Then
                Throw New MedscreenExceptions.MedscreenException("Attempt to create duplicate entry")
            Else
                Medscreen.LogError(ex, , "update " & Me.UpdateString)
            End If
        Catch ex As Exception
            Medscreen.LogError(ex, , "update " & Me.UpdateString)
            'Conn.Close()
            Return False
        Finally
            'oledconn.Close()
            Try
                Conn.Close()

            Catch ex As Exception

            End Try
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' GetRowId from database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/01/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function GetRowId() As String
        If Not Me.objRowID.IsSet Then
            Try
                Dim oCMd As New OleDb.OleDbCommand()
                oCMd.Connection = MedConnection.Connection
                oCMd.CommandText = Me.RowIdSelect(True)


                If CConnection.ConnOpen Then
                    Me.RowID = oCMd.ExecuteScalar
                End If
            Catch ex As Exception
                Medscreen.LogError(ex, , "GetRowID")
            Finally
                CConnection.SetConnClosed()
            End Try
        Else
        End If
        Return Me.RowID
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Property to control whether this table set is read only or not 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ReadOnlyTable() As Boolean
        Get
            Return Me.blnReadonly
        End Get
        Set(ByVal Value As Boolean)
            Me.blnReadonly = Value
        End Set
    End Property


    '''<summary>
    ''' Constructor
    ''' </summary>
    ''' <param name='TableName'>Name of table in the database</param>
    Public Sub New(ByVal TableName As String)
        MyBase.New()
        strTableName = TableName
        strHeader = strTableName
        Me.List.Add(Me.objRowID)

    End Sub


    Public Property RowID() As String
        Get
            Return Me.objRowID.Value
        End Get
        Set(ByVal Value As String)
            Me.objRowID.Value = Value
        End Set
    End Property
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' find the position of the field by name
    ''' </summary>
    ''' <param name="FieldName">Name of field</param>
    ''' <returns>Position of field</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IndexOf(ByVal FieldName As String) As Integer
        Dim objtf As TableField
        Dim i As Integer = 0
        For Each objtf In Me.List
            If objtf.FieldName.ToUpper = FieldName.ToUpper Then
                Exit For
            End If
            i += 1
        Next
        Return i
    End Function

    '''<summary>
    ''' Fill the table
    ''' </summary>
    ''' <param name='inFields'>Fill this table from these fields</param>
    ''' <returns>TRUE if succesful</returns>
    Public Function Fill(ByVal inFields As TableFields) As Boolean
        Dim objTf As TableField
        Dim objTme As TableField
        Dim i As Integer
        Dim blnRet As Boolean = False

        If inFields.Count <> Me.Count Then Exit Function

        For i = 0 To inFields.Count - 1
            objTf = inFields.Item(i)
            If Not objTf.IsIdentity Then
                objTme = Me.Item(i)
                If objTme.FieldName.ToUpper = objTf.FieldName.ToUpper Then
                    objTme.Value = objTf.Value
                    objTme.OldValue = objTf.OldValue
                Else ' Field mismatch
                    Return False
                    Exit For
                End If
            End If
        Next
        blnRet = True
        Return blnRet
    End Function

    '''<summary>
    ''' transform these data in HTML using a stylsheet
    ''' </summary>
    ''' <param name='StyleSheetID'>Stylesheet to use</param>
    ''' <param name='inFileName'>Source file optional</param>
    ''' <returns>TRUE if succesful</returns>
    Public Function DoTransform(Optional ByVal StyleSheetID As String = "test4.xsl", _
    Optional ByVal inFileName As String = "") As String

        Dim xslt As Xsl.XslTransform = New Xsl.XslTransform()
        Dim strTemp As String

        Dim StyleSheet As String = MedscreenLib.Constants.GCST_X_DRIVE & "\Lab Programs\Transforms\XSL\"
        Dim resolver As XmlUrlResolver = New XmlUrlResolver()

        StyleSheet += StyleSheetID
        Dim fileName As String = "C:\temp\" & Now.ToString("yyyyMMddhhmmss").Trim & ".xml"
        Dim fileNameX As String = "C:\temp\" & Now.ToString("yyyyMMddhhmmss").Trim & ".html"

        'Create two temporary file names 

        If inFileName.Trim.Length = 0 Then
            'Export data to XML File 
            Me.XmlToFile(fileName, True)
        Else
            fileName = inFileName 'Supplied prepared XMLFile
        End If
        'Load as resolver using passed stylesheet
        xslt.Load(StyleSheet, resolver)

        'Create an XPATh document from the XML we have written
        Dim doc As XPath.XPathDocument = New XPath.XPathDocument(fileName)
        'And a writer to use for output purposes
        Dim writer As XmlTextWriter = New XmlTextWriter(fileNameX, System.Text.Encoding.UTF8)

        'Transform the sourde document to an output file, which will be XML
        xslt.Transform(doc, Nothing, writer)

        'Get rid of the source XML file 
        doc = Nothing
        writer.Flush()
        writer.Close()

        xslt = Nothing
        System.IO.File.Delete(fileName)

        'Read in the HTML so we can emit it as atring 
        Dim Readr As New IO.StreamReader(fileNameX)

        strTemp = Readr.ReadToEnd()
        Readr.Close()
        System.Threading.Thread.Sleep(200)
        'Now delete it we have no need of it any longer.
        System.IO.File.Delete(fileNameX)

        Return strTemp


    End Function

    '''<summary>
    ''' table reffered to
    ''' </summary>
    Public Property Table() As String
        Get
            Return Me.strTableName
        End Get
        Set(ByVal Value As String)
            Me.strTableName = Value
        End Set
    End Property

    '''<summary>
    ''' Field within the database row, indexed by Field position
    ''' </summary>
    ''' <param name='index'>Index into the table</param>
    Default Public Property Item(ByVal index As Integer) As TableField
        Get
            Return CType(MyBase.List.Item(index), TableField)
        End Get
        Set(ByVal Value As TableField)
            MyBase.List.Item(index) = Value
        End Set
    End Property

    '''<summary>
    ''' Field within the database table, indexed by Field name
    ''' </summary>
    ''' <param name='index'>Index as string</param>
    Default Public ReadOnly Property Item(ByVal index As String) As TableField
        Get

            Dim objTF As TableField = Nothing
            Dim I As Integer
            For I = 0 To Me.Count - 1
                objTF = CType(MyBase.List.Item(I), TableField)
                If objTF.FieldName.ToUpper = index.ToUpper Then
                    Exit For
                End If
                objTF = Nothing
            Next
            Return objTF
        End Get
    End Property

    '''<summary>
    ''' Add a field to the list of fields
    ''' </summary>
    ''' <param name='item'>Field to add</param>
    ''' <returns> Index of field</returns>
    Public Function Add(ByVal item As TableField) As Integer
        Dim aItem As TableField
        aItem = Me.Item(item.FieldName)
        If aItem Is Nothing Then
            Return MyBase.List.Add(item)
            item.TableName = Me.Table
        Else
            Return MyBase.List.IndexOf(aItem)
        End If
    End Function

    '''<summary>
    ''' Count of the number of fields in the database
    ''' </summary>
    ''' <returns>Count of the number of fields</returns>
    Public Shadows Function Count() As Integer
        Return MyBase.Count
    End Function

    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <param name='Format'>Format for the XML</param>
    ''' <param name='HeaderText'>Top level entity name</param>
    ''' <returns>XML representation of row</returns>
    ''' <remarks>Format property will allow easy variations on the theme, it is not yet encoded</remarks>
    Public Overloads Function ToXML(ByVal Format As Integer, _
    ByVal HeaderText As String) As String
        'Format property will allow easy variations on the theme

        Dim strRet As String = "<" & HeaderText & ">"

        Dim i As Integer
        Dim objTf As TableField

        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToXMLSchema
            Next
        Catch ex As Exception
        End Try

        strRet += "</" & HeaderText & ">"
        Return strRet
    End Function

    '''<summary>
    ''' Convert this row to CSV Comma Separated Variables
    ''' </summary>
    ''' <returns>This data row converted to CSV</returns>
    Public Overloads Function ToCSV() As String
        Dim strRet As String = ""

        Dim i As Integer
        Dim objTf As TableField



        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToCSV & ","
            Next
        Catch ex As Exception
        End Try

        Return strRet

    End Function

    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <param name='blnNoHeader'>Don't add a W3C XML header</param>
    ''' <param name='StyleSheet'>XSL StyleSheet to use</param>
    ''' <param name='Header'>Top level entity name Default = "", means don't override</param>
    ''' <param name='blnDontCloseHeader'>Don't close top level entity default is FALSE</param>
    ''' <returns>XML representation of row</returns>
    ''' <remarks>Use the ToXMLSchema versions in preference</remarks>
    Public Overloads Function ToXML(Optional ByVal blnNoHeader As Boolean = False, _
    Optional ByVal StyleSheet As String = "test2.xsl", _
    Optional ByVal Header As String = "", _
    Optional ByVal blnDontCloseHeader As Boolean = False) As String
        Dim strRet As String = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf

        Dim i As Integer
        Dim objTf As TableField

        If blnNoHeader Then
            strRet = ""
        Else
            strRet += "<?xml-stylesheet type=""text/xsl"" href=""" & StyleSheet & """?>" & vbCrLf
        End If

        If Header.Trim.Length > 0 Then
            strRet += "<" & Header & ">" & vbCrLf
        End If
        strRet += "<" & Me.strHeader & ">" & vbCrLf
        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToXML
            Next
        Catch ex As Exception
        End Try

        strRet += "</" & Me.strHeader & ">"
        If Header.Trim.Length > 0 Then
            If Not blnDontCloseHeader Then
                strRet += "</" & Header & ">" & vbCrLf
            End If
        End If
        Return strRet

    End Function

    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <param name='blnNoHeader'>Don't add a W3C XML header</param>
    ''' <param name='StyleSheet'>XSL StyleSheet to use</param>
    ''' <param name='Header'>Top level entity name Default = "", means don't override</param>
    ''' <param name='blnDontCloseHeader'>Don't close top level entity default is FALSE</param>
    ''' <returns>XML representation of row</returns>
    Public Overloads Function ToXMLSchema(ByVal blnNoHeader As Boolean, _
    Optional ByVal StyleSheet As String = "test2.xsl", _
    Optional ByVal Header As String = "", _
    Optional ByVal blnDontCloseHeader As Boolean = False) As String
        Dim strRet As String = "<?xml version=""1.0"" standalone=""yes"" ?>" & vbCrLf

        Dim i As Integer
        Dim objTf As TableField

        If blnNoHeader Then
            strRet = ""
        Else
            strRet += "<?xml-stylesheet type=""text/xsl"" href=""" & StyleSheet & """?>" & vbCrLf
        End If

        If Header.Trim.Length > 0 Then
            strRet += "<" & Header & ">" & vbCrLf
        End If
        strRet += "<" & Me.strHeader & ">" & vbCrLf
        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToXMLSchema & vbCrLf
            Next
        Catch ex As Exception
        End Try

        strRet += "</" & Me.strHeader & ">"
        If Header.Trim.Length > 0 Then
            If Not blnDontCloseHeader Then
                strRet += "</" & Header & ">" & vbCrLf
            End If
        End If
        Return strRet

    End Function

    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <returns>XML representation of row</returns>
    Public Overloads Function ToXMLSchema() As String
        Dim strRet As String = "" & vbCrLf

        Dim i As Integer
        Dim objTf As TableField


        If Me.strHeader.Trim.Length <> 0 Then strRet += "<" & Me.strHeader & ">" & vbCrLf
        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToXMLSchema
                If i Mod 8 = 0 Then strRet += vbCrLf
            Next
        Catch ex As Exception
        End Try

        If Me.strHeader.Trim.Length <> 0 Then strRet += "</" & Me.strHeader & ">"

        Return strRet

    End Function

    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <param name='UseOption'>Don't add a W3C XML header</param>
    ''' <returns>XML representation of row</returns>
    Public Overloads Function ToXMLSchema(ByVal UseOption As Integer) As String
        Dim strRet As String = "" & vbCrLf

        Dim i As Integer
        Dim objTf As TableField


        If UseOption = 0 Then strRet += "<" & Me.strHeader & ">" & vbCrLf
        Try
            For i = 0 To Me.List.Count - 1
                objTf = Me.Item(i)
                strRet += objTf.ToXMLSchema
                If i Mod 8 = 0 Then strRet += vbCrLf
            Next
        Catch ex As Exception
        End Try

        If UseOption = 0 Then strRet += "</" & Me.strHeader & ">"

        Return strRet

    End Function



    '''<summary>
    ''' Produce a CSV File Header, e.g. the titles of the fields
    ''' </summary>
    ''' <returns>CSV representation of row header</returns>
    Public Function CSVFileHeader() As String
        'Dim w As IO.StreamWriter
        Dim i As Integer
        Dim strOut As String
        Dim dQuote As String = """"
        Dim objTf As TableField

        Try
            strOut = ""
            For i = 0 To list.Count - 1
                objTf = Item(i)
                strOut += objTf.FieldName & ","
            Next
            Return strOut
        Catch ex As Exception
            Return ""
        End Try

    End Function

    '''<summary>
    ''' Produce a CSV File Data row
    ''' </summary>
    ''' <returns>CSV representation of row </returns>
    Public Function CSVFileRow() As String
        'Dim w As IO.StreamWriter
        Dim i As Integer
        Dim strOut As String
        Dim dQuote As String = """"
        Dim objTf As TableField

        Try
            strOut = ""
            For i = 0 To list.Count - 1
                objTf = Item(i)
                strOut += objTf.ValueToDbString(dQuote) & ","
            Next
            Return strOut
        Catch ex As Exception
            Return ""
        End Try

    End Function



    '''<summary>
    ''' Convert this row to XML
    ''' </summary>
    ''' <param name='Filename'>Name of file to store data in</param>
    ''' <param name='blnNoHeader'>Don't add a W3C XML header</param>
    ''' <param name='StyleSheet'>XSL StyleSheet to use</param>
    ''' <param name='Header'>Top level entity name Default = "", means don't override</param>
    ''' <param name='blnDontCloseHeader'>Don't close top level entity default is FALSE</param>
    Public Sub XmlToFile(ByVal Filename As String, _
       Optional ByVal blnNoHeader As Boolean = False, _
       Optional ByVal StyleSheet As String = "test4.xsl", _
       Optional ByVal Header As String = "", _
       Optional ByVal blnDontCloseHeader As Boolean = False)
        Dim w As IO.StreamWriter

        Try
            w = New IO.StreamWriter(Filename, False)
            w.Write(Me.ToXML(blnNoHeader, StyleSheet, Header, blnDontCloseHeader))
            w.Flush()
            w.Close()

        Catch ex As Exception
        End Try

    End Sub

    '''<summary>
    ''' produce a string which can be used as command to insert these data into a database
    ''' </summary>
    ''' <returns>Insert String</returns>
    Public Overridable Function InsertString() As String
        Dim strRet As String = ""
        Dim objtf As TableField
        Dim blnFirst As Boolean

        strRet = "insert into " & strTableName & "("
        blnFirst = True
        For Each objtf In Me.List
            If Not TypeOf objtf Is ROWIDField Then
                If Not blnFirst Then strRet += ","
                If objtf.FieldName <> "ROWID" Then strRet += objtf.FieldName
                blnFirst = False
            End If
        Next

        strRet += ") Values("
        blnFirst = True
        For Each objtf In Me.List
            If Not TypeOf objtf Is ROWIDField Then
                If Not blnFirst Then strRet += ","
                'deal with user fields 
                If TypeOf objtf Is UserField Then
                    If Not Glossary.Glossary.CurrentSMUser Is Nothing Then
                        strRet += "'" & Glossary.Glossary.CurrentSMUser.Identity & "'"
                    Else
                        strRet += "NULL"
                    End If
                    'deal with timestamps
                ElseIf TypeOf objtf Is TimeStampField Then
                    strRet += "SYSDATE"
                Else
                    If objtf.FieldName <> "ROWID" Then strRet += objtf.ValueToDbString
                    blnFirst = False
                End If
            End If
        Next
        strRet += ")"
        Return strRet

    End Function


    '''<summary>
    ''' Indicates whether the data has been retrieved from the database or is a new record
    ''' </summary>
    ''' <returns>TRUE = has been read from database</returns>
    Public Property Loaded() As Boolean
        Get
            If Me.RowID.Trim.Length > 0 Then blnLoaded = True
            Return blnLoaded
        End Get
        Set(ByVal Value As Boolean)
            blnLoaded = Value
        End Set
    End Property

    '''<summary>
    ''' produce a string which can be used as command to update these data in a database
    ''' </summary>
    ''' <param name='ChangeID'>IF TRUE the primary keys can be modified</param>
    ''' <returns>Insert String</returns>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action>Read Only tables dealt with</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overridable Overloads Function UpdateString(Optional ByVal ChangeID As Boolean = False) As String
        Dim strWhere As String = " where "
        Dim objtf As TableField

        If Me.ReadOnlyTable Then
            Return ""
            Exit Function
        End If

        Dim strRet As String = "Update " & strTableName & " set "
        Dim strRet1 As String = "Update " & strTableName & " set "

        'See if we can update by RowId 
        Dim blnByRowId As Boolean = False
        If Me.RowID.Trim.Length > 0 Then
            strWhere = " Where rowid = '" & Me.RowID.Trim & "'"
            blnByRowId = True
        End If

        For Each objtf In Me.List
            If objtf.IsIdentity AndAlso Not blnByRowId Then

                If strWhere.Length > 7 Then strWhere += " and "
                strWhere += objtf.FieldName & " = " & objtf.ValueToDbString(, ChangeID)
                If objtf.Changed Then
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "
                    End If
                    'strRet += objtf.FieldName & " = " & objtf.ValueToDbString
                    strRet += objtf.UpdateString
                End If
            ElseIf TypeOf objtf Is UserField Then  'If this is a user stamp field 
                If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then  'If we have a user
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "              'Do we need to add a comma
                    End If
                    strRet += objtf.FieldName & " = '" & MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity.Trim & "'"

                End If
            Else
                'ensure that timestamp fields get updated
                If objtf.Changed Or _
                    TypeOf objtf Is BooleanField Or _
                    TypeOf objtf Is TimeStampField Then
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "
                    End If
                    'strRet += objtf.FieldName & " = " & objtf.ValueToDbString
                    strRet += objtf.UpdateString
                End If
            End If
        Next

        If strRet = strRet1 Then ' Nothing to Update
            strRet = ""
        Else
            strRet += " " & strWhere
        End If

        Return strRet

    End Function

    Public Overridable Overloads Function UpdateString(ByRef Cmd As OleDb.OleDbCommand, ByRef DataChanged As Boolean, Optional ByVal ChangeID As Boolean = False) As String
        Dim strWhere As String = " where "
        Dim objtf As TableField

        If Me.ReadOnlyTable Then
            Return ""
            Exit Function
        End If

        Dim strRet As String = "Update " & strTableName & " set "
        Dim strRet1 As String = "Update " & strTableName & " set "
        Dim setParam As OleDb.OleDbParameter        'Define parameter for string 
        Dim WhereColl As Collection = New Collection()
        For Each objtf In Me.List
            If objtf.IsIdentity Then

                If strWhere.Length > 7 Then strWhere += " and "
                strWhere += objtf.FieldName & " = ?" '& objtf.ValueToDbString(, ChangeID)
                If ChangeID Then
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "
                    End If
                    If strRet1.Length > Len("Update " & strTableName & " set ") Then
                        strRet1 += ", "
                    End If
                    strRet += objtf.FieldName & " = ? " ' & objtf.ValueToDbString
                    strRet1 += objtf.FieldName & " =  " & objtf.ValueToDbString
                    DataChanged = DataChanged Or objtf.Changed          'See if the value has changed
                    'strRet += objtf.UpdateString
                    setParam = SetUpParameter(objtf)
                    Cmd.Parameters.Add(setParam)
                End If
                setParam = SetUpParameter(objtf)

                WhereColl.Add(setParam)
                'End If
            ElseIf TypeOf objtf Is UserField Then  'If this is a user stamp field 
                If Not MedscreenLib.Glossary.Glossary.CurrentSMUser Is Nothing Then  'If we have a user
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "              'Do we need to add a comma
                    End If
                    strRet += objtf.FieldName & " = ? " ' & MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity.Trim & "'"
                    setParam = New OleDb.OleDbParameter(objtf.FieldName, MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity.Trim)
                    setParam.DbType = DbType.String
                    setParam.Size = objtf.FieldLength
                    Cmd.Parameters.Add(setParam)
                End If
                'Timestampfields
            ElseIf TypeOf objtf Is TimeStampField Then
                strRet += "," & objtf.FieldName & " = sysdate "
            Else
                If objtf.FieldName <> "ROWID" Then
                    'ensure that timestamp fields get updated
                    'If objtf.Changed Or _
                    '              TypeOf objtf Is BooleanField Or _
                    '              TypeOf objtf Is TimeStampField Then
                    DataChanged = DataChanged Or objtf.Changed  'See if anything has changed
                    If strRet.Length > Len("Update " & strTableName & " set ") Then
                        strRet += ", "
                    End If
                    If strRet1.Length > Len("Update " & strTableName & " set ") Then
                        strRet1 += ", "
                    End If
                    strRet += objtf.FieldName & " = ?" '& objtf.ValueToDbString
                    strRet1 += objtf.FieldName & "=" & objtf.ValueToDbString
                    setParam = SetUpParameter(objtf)
                    Cmd.Parameters.Add(setParam)
                    'strRet += objtf.UpdateString
                End If
            End If
        Next

        If Me.objRowID.IsSet Then
            'We can use the rowid to update the data
            strWhere = " Where rowid = ?"
            Cmd.Parameters.Add(CConnection.StringParameter("RowID", Me.objRowID.Value, 20))
        Else
            Dim k As Integer
            For k = 1 To WhereColl.Count
                Cmd.Parameters.Add(WhereColl.Item(k))
            Next
        End If

        If strRet = strRet1 Then ' Nothing to Update
            strRet = ""
        Else
            strRet += " " & strWhere
        End If

        Return strRet

    End Function

    Friend Function SetUpParameter(ByVal objtf As TableField, Optional ByVal prefix As String = "") As OleDb.OleDbParameter
        Dim setParam As OleDb.OleDbParameter = New OleDb.OleDbParameter(prefix & objtf.FieldName, objtf.Value)
        If TypeOf objtf Is StringField Then
            setParam.DbType = DbType.String
            setParam.Size = objtf.FieldLength
        ElseIf TypeOf objtf Is TimeStampField Then
            setParam.DbType = DbType.DateTime
            objtf.SetIsNull = False                 'Ensure that it is not null
            setParam.Value = Now
        ElseIf TypeOf objtf Is DateField Then
            setParam.DbType = DbType.DateTime
            If objtf.Value = DateField.ZeroDate Then
                setParam.Value = DBNull.Value
            Else
                setParam.Value = objtf.Value
            End If
        ElseIf TypeOf objtf Is IntegerField Then
            setParam.DbType = DbType.Int32
        ElseIf TypeOf objtf Is DoubleField Then
            setParam.DbType = DbType.Double
        ElseIf TypeOf objtf Is BooleanField Then
            setParam.DbType = DbType.String
            setParam.Size = 1
            If CType(objtf, BooleanField).Value Then
                setParam.Value = CType(objtf, BooleanField).TrueChar
            Else
                setParam.Value = CType(objtf, BooleanField).FalseChar
            End If
        End If
        If objtf.IsNull AndAlso Not (TypeOf objtf Is BooleanField) AndAlso objtf.Value = objtf.NVL Then _
            setParam.Value = System.DBNull.Value
        Return setParam
    End Function

    Public Overloads Function readfields(ByVal oRead As Xml.XmlElement) As Boolean
        Dim objTf As TableField
        Dim objBf As BooleanField
        Dim obj As Object
        Me.InitialiseOldValues()
        Try
            For Each objTf In Me.List
                If Not objTf.IsIdentity Then
                    Dim blnIsNull As Boolean
                    obj = MedscreenLib.Medscreen.ReadValue(oRead, objTf.FieldName, objTf.NVL, blnIsNull)
                    objTf.SetIsNull = blnIsNull
                    If Not obj Is Nothing Then
                        If TypeOf objTf Is BooleanField Then
                            objBf = objTf
                            objBf.Value = (CType(objTf, BooleanField).TrueChar = obj)
                            objBf.OldValue = objBf.Value
                            objTf = objBf
                        Else
                            objTf.Value = obj
                            objTf.OldValue = objTf.Value
                        End If
                    Else
                        objTf.Value = objTf.NVL
                        objTf.OldValue = objTf.NVL
                    End If
                End If
            Next
            blnLoaded = False 'Loading from a node list not a table so not loaded 
            Me.datLoaded = Now

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' readfield data from an xmlnodelist 
    ''' </summary>
    ''' <param name="oread"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [09/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Function ReadFields(ByVal oread As Xml.XmlNodeList) As Boolean
        Dim objTf As TableField
        Dim objBf As BooleanField
        Dim obj As Object

        Try
            For Each objTf In Me.List
                Dim blnIsNull As Boolean
                obj = MedscreenLib.Medscreen.ReadValue(oread, objTf.FieldName, objTf.NVL, blnIsNull)
                objTf.SetIsNull = blnIsNull
                If Not obj Is Nothing Then
                    If TypeOf objTf Is BooleanField Then
                        objBf = objTf
                        objBf.Value = (CType(objTf, BooleanField).TrueChar = obj)
                        objBf.OldValue = objBf.Value
                        objTf = objBf
                    Else
                        objTf.Value = obj
                        objTf.OldValue = objTf.Value
                    End If
                Else
                    objTf.Value = objTf.NVL
                    objTf.OldValue = objTf.NVL
                End If

            Next
            blnLoaded = False   'Loading from a node list not a table so not loaded
            Me.datLoaded = Now
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    '''<summary>
    ''' read a row in the database and populate contents of collection (OLEDB Version)
    ''' </summary>
    ''' <param name='oRead'>Data reader</param>
    ''' <returns>TRUE if succesful</returns>
    Public Overloads Function ReadFields(ByVal oRead As OleDb.OleDbDataReader) As Boolean
        Dim obj As Object
        Dim objTf As TableField
        Dim objBf As BooleanField

        Try
            For Each objTf In Me.List

                Dim blnIsNull As Boolean
                obj = MedscreenLib.Medscreen.ReadValue(oRead, objTf.FieldName, objTf.NVL, blnIsNull)
                objTf.SetIsNull = blnIsNull
                If Not obj Is Nothing Then
                    If TypeOf objTf Is BooleanField Then
                        objBf = objTf
                        objBf.Value = (CType(objTf, BooleanField).TrueChar = obj)
                        objBf.OldValue = objBf.Value
                        objTf = objBf
                    Else
                        objTf.Value = obj
                        objTf.OldValue = objTf.Value
                    End If
                Else
                    objTf.Value = objTf.NVL
                    objTf.OldValue = objTf.NVL
                End If
            Next
            blnLoaded = True
            Me.datLoaded = Now
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    '''<summary>
    ''' read a row in the database and populate contents of collection (DataRow Version)
    ''' </summary>
    ''' <param name='oRead'>Data reader</param>
    ''' <returns>TRUE if succesful</returns>
    ''' <revisionHistory><revision><modified>28-Nov-2011 07:45 altered to deal with sql server bit field type</modified><Author>CONCATENO\Taylor</Author></revision></revisionHistory>
    Public Overloads Function ReadFields(ByVal oRead As DataRow) As Boolean
        Dim obj As Object
        Dim objTf As TableField
        Dim objBf As BooleanField

        If oRead Is Nothing Then Exit Function

        Dim oTable As DataTable = oRead.Table                   'Get the parent table
        If oTable Is Nothing Then Exit Function

        Dim oColumns As DataColumnCollection = oTable.Columns   'Get its schema
        If oColumns Is Nothing Then Exit Function

        Dim IColumn As Integer
        Try
            For Each objTf In Me.List
                IColumn = oColumns.IndexOf(objTf.FieldName) ' Find field index
                If IColumn >= 0 Then
                    obj = oRead.Item(IColumn)               ' Get value 
                    If oRead.IsNull(IColumn) Then
                        obj = objTf.NVL                     ' Set = to null value 
                    End If

                    If Not obj Is Nothing Then              'If we have a value 
                        If TypeOf objTf Is BooleanField Then
                            objBf = objTf
                            If TypeOf obj Is Boolean Then   'If the data type is boolean in the database then use that direct
                                objbf.value = obj
                            Else                            'Else use a string representation of the data 
                                objBf.Value = (objBf.TrueChar = obj)

                            End If

                            objBf.OldValue = objBf.Value
                            objTf = objBf
                        Else
                            objTf.Value = obj
                            objTf.OldValue = objTf.Value
                        End If
                    Else
                        objTf.Value = objTf.NVL
                        objTf.OldValue = objTf.NVL
                    End If
                End If
            Next
            blnLoaded = True
            Me.datLoaded = Now

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

  
    '''<summary>
    ''' read a row in the database and populate contents of collection (SQL Server Client Version)
    ''' </summary>
    ''' <param name='oRead'>Data reader</param>
    ''' <returns>TRUE if succesful</returns>
    Public Overloads Function ReadFields(ByVal oRead As SqlClient.SqlDataReader) As Boolean
        Dim obj As Object
        Dim objTf As TableField
        Dim objBf As BooleanField

        Try
            For Each objTf In Me.List
                Dim blnIsNull As Boolean
                obj = MedscreenLib.Medscreen.ReadValue(oRead, objTf.FieldName, objTf.NVL, blnIsNull)
                objTf.SetIsNull = blnIsNull
                If Not obj Is Nothing Then
                    If TypeOf objTf Is BooleanField Then
                        objBf = objTf
                        objBf.Value = (objBf.TrueChar = obj)
                        objBf.OldValue = objBf.Value
                        objTf = objBf
                    Else
                        objTf.Value = obj
                        objTf.OldValue = objTf.Value
                    End If
                Else
                    objTf.Value = objTf.NVL
                    objTf.OldValue = objTf.NVL
                End If
            Next
            blnLoaded = True
            Me.datLoaded = Now

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function



    '

    '''<summary>
    ''' Commit the values i.e. update the old (read values) values with the new ones 
    ''' so it will now appear unchanged
    ''' </summary>
    Public Sub Commit()
        Dim objTf As TableField

        For Each objTf In Me.List
            If objTf.Changed Then
                objTf.OldValue = objTf.Value
            End If
        Next

    End Sub



    '''<summary>
    ''' Initialise values to the NVL value
    ''' </summary>
    Public Sub Initialise()
        Dim objTf As TableField

        For Each objTf In Me.List
            If Not objTf.IsIdentity Then
                objTf.OldValue = objTf.NVL
                objTf.Value = objTf.NVL
            End If
        Next
        Me.datAccessed = Now


    End Sub

    '''<summary>
    ''' Initialise Old (read values) values to the NVL value
    ''' </summary>
    Public Sub InitialiseOldValues()
        Dim objTf As TableField

        For Each objTf In Me.List
            If Not objTf.IsIdentity Then
                objTf.OldValue = objTf.NVL
            End If
        Next

    End Sub

    '''<summary>
    ''' Produce an XSD description of these data
    ''' </summary>
    ''' <returns>an XSD description of these data</returns>
    Public Function XSDSchema() As String
        Dim strRet As String = "<?xml version=""1.0"" ?>"
        strRet += "<xs:schema id=""" & _
            Me.strTableName & """ targetNamespace=""http://tempuri.org/" & _
            Me.strTableName & "1.xsd"" xmlns:mstns=""http://tempuri.org/" & _
            Me.strTableName & "1.xsd"" xmlns=""http://tempuri.org/" & _
            Me.strTableName & "1.xsd"" xmlns:xs=""http://www.w3.org/2001/XMLSchema""" & _
            " xmlns:msdata=""urn:schemas-microsoft-com:xml-msdata"" attributeFormDefault=""" & _
            "qualified"" elementFormDefault=""qualified""> "
        strRet += " <xs:element name=""" & Me.strTableName & """ msdata:IsDataSet=""true""" & _
            " msdata:Locale=""en-GB"" msdata:EnforceConstraints=""False"">" & vbCrLf

        strRet += "  <xs:complexType>" & vbCrLf
        strRet += "   <xs:choice maxOccurs=""unbounded"">" & vbCrLf
        strRet += "    <xs:element name=""" & Me.strTableName.ToLower & """>" & vbCrLf
        strRet += "     <xs:complexType>" & vbCrLf
        strRet += "      <xs:sequence>" & vbCrLf
        Dim objTf As TableField

        For Each objTf In Me.List
            strRet += objTf.XSDSchema & vbCrLf
        Next

        strRet += "      </xs:sequence>" & vbCrLf
        strRet += "     </xs:complexType>" & vbCrLf
        strRet += "    </xs:element>" & vbCrLf
        strRet += "   </xs:choice>" & vbCrLf
        strRet += "  </xs:complexType>" & vbCrLf
        strRet += " </xs:element>" & vbCrLf
        strRet += "</xs:schema>" & vbCrLf
        Return strRet

    End Function

    '''<summary>
    ''' Produce a command string to delete from database
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    Public Overridable Function DeleteString() As String
        Dim strRet As String = ""
        Dim strWhere As String = " where "
        Dim objtf As TableField

        strRet = "Delete from  " & strTableName
        For Each objtf In Me.List
            If objtf.IsIdentity Then
                If strWhere.Length > 7 Then strWhere += " and "
                strWhere += objtf.FieldName & " = " & objtf.ValueToDbString
            End If
        Next
        strRet += " " & strWhere
        Return strRet


    End Function

    '''<summary>
    ''' Rollback data by replacing Current values with Read or original values
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    Public Function Rollback() As Boolean
        Dim objTf As TableField


        For Each objTf In Me.List
            If objTf.Changed Then
                objTf.Value = objTf.OldValue
            End If
        Next

        Return True
    End Function

    Public Function RowIdSelect(Optional ByVal WhereClause As Boolean = False, Optional ByVal oCmd As OleDb.OleDbCommand = Nothing) As String
        Dim strReturn As String = "Select a.ROWID from " & Me.strTableName & " a "
        If WhereClause Then
            Dim objTf As TableField
            Dim strWhere As String = " where "
            For Each objTf In Me.List                           'Go through each field
                If (objTf.IsIdentity) Then                      'If it is an identity field
                    If strWhere.Length > 7 Then                 'See if we have added anything to where clause
                        strWhere += " and "                     'If so add an and 
                    End If
                    '                                           'Add where clause item
                    If oCmd Is Nothing Then
                        strWhere += objTf.FieldName & " = " & objTf.ValueToDbString
                    Else
                        strWhere += objTf.FieldName & " = ? "
                        If TypeOf objTf Is DateField Then
                            Dim objD As DateField = objTf
                            oCmd.Parameters.Add(CConnection.StringParameter(objTf.FieldName, "TO_DATE('" & objD.Value.ToString("ddMMyyyy HHmm") & "','DDMMYYYY HH24mi')", 100))
                        Else

                            oCmd.Parameters.Add(CConnection.StringParameter(objTf.FieldName, objTf.Value, 20))
                        End If
                    End If
                End If
            Next
            If strWhere.Length > 7 Then                         'Did we construct a where clause 
                strReturn += strWhere
            End If
        End If
        Return strReturn
    End Function


    Public Function FullRowSelect(Optional ByVal WhereClause As Boolean = False, Optional ByVal oCmd As OleDb.OleDbCommand = Nothing) As String
        Dim strReturn As String = "Select a.*,a.ROWID from " & Me.strTableName & " a "
        If WhereClause Then
            Dim objTf As TableField
            Dim strWhere As String = " where "
            For Each objTf In Me.List                           'Go through each field
                If (objTf.IsIdentity) Then                      'If it is an identity field
                    If strWhere.Length > 7 Then                 'See if we have added anything to where clause
                        strWhere += " and "                     'If so add an and 
                    End If
                    '                                           'Add where clause item
                    If oCmd Is Nothing Then
                        strWhere += objTf.FieldName & " = " & objTf.ValueToDbString
                    Else
                        strWhere += objTf.FieldName & " = ? "
                        oCmd.Parameters.Add(CConnection.StringParameter(objTf.FieldName, objTf.Value, 20))
                    End If
                End If
            Next
            If strWhere.Length > 7 Then                         'Did we construct a where clause 
                strReturn += strWhere
            End If
        End If
        Return strReturn
    End Function

    '''<summary>
    ''' Produce a command string to select these data
    ''' </summary>
    ''' <param name='AddTable'>TRUE = FROM Table<para/>FALSE don't add from clause</param>
    ''' <param name='AddSelect'>TRUE = prefix with SELECT (default is TRUE)</param>
    ''' <param name='SkipIdentityFields'>TRUE = Don't include Primary Key fields (default = FALSE)</param>
    '''<returns>Return a select command string</returns>
    '''<revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [15/11/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    Public Overridable Function SelectString(Optional ByVal AddTable As Boolean = True, _
    Optional ByVal AddSelect As Boolean = True, _
    Optional ByVal SkipIdentityFields As Boolean = False, _
    Optional ByVal AddWhereClause As Boolean = False) As String
        Dim objTf As TableField
        Dim strRet As String = ""

        If AddSelect Then strRet = "Select "
        For Each objTf In Me.List
            If Not (objTf.IsIdentity And SkipIdentityFields) Then
                If strRet <> "Select " Then
                    strRet += ", "
                End If
                strRet += objTf.FieldName
            End If
        Next
        If AddTable Then strRet += " from " & Me.strTableName
        'Deal with adding a where clause automatically
        Dim strWhere As String = " Where "
        If AddSelect And AddTable And AddWhereClause Then       'We are going to construct a full query
            For Each objTf In Me.List                           'Go through each field
                If (objTf.IsIdentity) Then                      'If it is an identity field
                    If strWhere.Length > 7 Then                 'See if we have added anything to where clause
                        strWhere += " and "                     'If so add an and 
                    End If
                    '                                           'Add where clause item
                    strWhere += objTf.FieldName & " = " & objTf.ValueToDbString

                End If
            Next
            If strWhere.Length > 7 Then                         'Did we construct a where clause 
                strRet += strWhere
            End If
        End If
        Return strRet
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Last access Date 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LastAccessDate() As Date
        Get
            Return Me.datAccessed
        End Get
        Set(ByVal Value As Date)
            datAccessed = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Last Load date 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LastLoadDate() As Date
        Get
            Return Me.datLoaded
        End Get
        Set(ByVal Value As Date)
            Me.datLoaded = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Last change date
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [13/10/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LastModifyDate() As Date
        Get
            Return Me.datUpdated
        End Get
        Set(ByVal Value As Date)
            Me.datUpdated = Value
        End Set
    End Property
End Class

#End Region

#Region "Interface to Tables or processes"

'''<summary>
''' Interface to the reporter process
''' </summary>
''' <remarks>
''' The reporter is used to send out custom reports via a number of different delivery channels<para/>
''' it works by using command files that are parameterised
''' </remarks>
Public Class ReporterInterfaceX

#Region "Declarations"


    Private strFilename As String = ""
    Private strOutputType As String = ""
    Private strOutputAddress As String = ""
    Private strReportTemplate As String = ""
    Private strAddress As String
    Private strSubject As String = ""
    Private strCompany As String = ""
    Private strCompanyID As String = ""
    Private strContact As String = ""

    Private RepType As Constants.ReportType
    Private RepMethod As Constants.CollOfficerSend

    Private w As IO.StreamWriter
#End Region

    '''<summary>
    ''' Customer's contact (recipient)
    ''' </summary>
    Public Property CustomerContact() As String
        Get
            Return strContact
        End Get
        Set(ByVal Value As String)
            strContact = Value
        End Set
    End Property

    '''<summary>
    ''' Full name of the client
    ''' </summary>
    Public Property ClientFullName() As String
        Get
            Return Me.strCompany
        End Get
        Set(ByVal Value As String)
            strCompany = Value
        End Set
    End Property

    '''<summary>
    ''' Client's SMID
    ''' </summary>
    Public Property ClientId() As String
        Get
            Return Me.strCompanyID
        End Get
        Set(ByVal Value As String)
            strCompanyID = Value
        End Set
    End Property

    '''<summary>
    ''' Subject of message
    ''' </summary>
    Public Property Subject() As String
        Get
            Return strSubject
        End Get
        Set(ByVal Value As String)
            strSubject = Value
        End Set
    End Property


    '''<summary>
    ''' Address to send to 
    ''' </summary>
    Public Property ReportAddress() As String
        Get
            Return strAddress
        End Get
        Set(ByVal Value As String)
            strAddress = Value
        End Set
    End Property

    '''<summary>
    ''' Type of report
    ''' </summary>
    Public Property ReportType() As Constants.ReportType
        Get
            Return RepType
        End Get
        Set(ByVal Value As Constants.ReportType)
            RepType = Value
        End Set
    End Property

    '''<summary>
    ''' How to send report
    ''' </summary>
    Public Property ReportMethod() As Constants.CollOfficerSend
        Get
            Return RepMethod
        End Get
        Set(ByVal Value As Constants.CollOfficerSend)
            RepMethod = Value
        End Set
    End Property

    '''<summary>
    ''' Create a filename for the report
    ''' </summary>
    ''' <param name='strName'>Filename to use</param>
    Public Function CreateFilename(ByVal strName As String) As String
        'Dim strRet As String
        Select Case RepType
            Case Constants.ReportType.CollectionRequest
                strFilename = getNextFileName(Medscreen.LiveRoot & "wordoutput\" & "Collection_Request-" & _
                strName, "out")
            Case Constants.ReportType.CollectionArrangeConf
                strFilename = getNextFileName(Medscreen.LiveRoot & "wordoutput\" & "CollArrConf" & "-" & Mid(strName, 1, 1) & _
                    strName.Substring(strName.Length - 4, 4), "out")
            Case Constants.ReportType.CollectionCompleteConf
                strFilename = getNextFileName(Medscreen.LiveRoot & "wordoutput\" & "CollCompConf" & "-" & Mid(strName, 1, 1) & _
                    strName.Substring(strName.Length - 4, 4), "out")
        End Select
        Return strFilename
    End Function

    Private Function getNextFileName(ByVal newname As String, ByVal suffix As String, _
    Optional ByVal StartAt As Integer = 0) As String
        Dim ext As Integer
        Dim strFileToFind As String

        ext = StartAt
        strFileToFind = newname & "-" & CStr(ext).Trim & "." & CStr(suffix).Trim
        While File.Exists(strFileToFind)
            ext += 1
            strFileToFind = newname & "-" & CStr(ext).Trim & "." & CStr(suffix).Trim
        End While
        Return strFileToFind
    End Function


    '''<summary>
    ''' Write out report
    ''' </summary>
    Public Overridable Function WriteReport() As Boolean
    End Function

    '''<summary>
    '''Name of file to use
    ''' </summary>
    Public ReadOnly Property FileName() As String
        Get
            Return strFilename
        End Get
    End Property

    '''<summary>
    '''Stream in use
    ''' </summary>
    Protected Property Stream() As StreamWriter
        Get
            Return w
        End Get
        Set(ByVal Value As StreamWriter)
            w = Value
        End Set
    End Property

    '''<summary>
    '''Constructor
    ''' </summary>
    ''' <param name='ReportType'>Type of report to create</param>
    Public Sub New(ByVal ReportType As Constants.ReportType)
        RepType = ReportType
    End Sub
End Class


#End Region
'Class to manage crystal reports
'An adaption has been made to allow generic report handling, this will use menu types to deal with these 
'additional types
'''<summary>
''' Class to manage automated reports
''' </summary>
''' <remarks>
'''An adaption has been made to allow generic report handling, this will use menu types to deal with these 
'''additional types<para/> 
''' This class handles report definition parameters, <see cref="MedscreenLib.CRFormulaItem"/> which are held in <see cref="MedscreenLib.CrFormulaItems"/>.  
''' These can be used to direct the report controlling code to locations where the parameters may be found.
''' <para/>This object was originally designed for creating Windows Menus for displaying reports, it can 
''' still be used that way, but additional <see cref="CRMenuItem.ReportTypes"/> have been defined to make the system more flexible.
''' At a minimum the report needs a <see cref="CRMenuItem.MenuIdentity "/> and a <see cref="CRMenuItem.MenuPath"/> defined, though not absolutely essential 
''' the <see cref="CRMenuItem.ReportType"/> should also be defined.
'''</remarks>
''' <seealso cref="CRFormulaItems"/>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/09/2005]</date><Action>Enumeration for report types Added, 
''' property returning enumeration added</Action></revision>
''' </revisionHistory>
Public Class CRMenuItem
#Region "Declarations"


    Private objIdentity As StringField
    Private objParent As StringField
    Private objText As StringField
    Private objPath As StringField
    Private objOrder As IntegerField
    Private objType As StringField
    Private objInstance As StringField = New StringField("INSTANCE", "", 100)
    Private objUserName As StringField = New StringField("USERNAME", "", 40)
    Private objPassword As StringField = New StringField("PASSWORD", "", 40)

    Private myHandle As System.IntPtr

    Private myFields As TableFields
    Private myFormulae As CrFormulaItems

#End Region

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enumeration of the different values that the Menu (report) type field can take
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/09/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum ReportTypes
        ''' <summary>Menu divider</summary>
        BAR
        ''' <summary>HTML report, calls a web site</summary>
        HTML
        ''' <summary>A Menu</summary>
        MNU
        ''' <summary>A Routine in a library for VGL the report 
        ''' path field should be a fully qualified routine name, parameters should be qualified in the Parameter collection</summary>
        PROC
        ''' <summary>Run a report</summary>
        RPT
        ''' <summary>Run an extended report</summary>
        RPTX

        SOAP
        ''' <summary>USe a Soap Service</summary>
        ''' 
        RPTP
    End Enum

#Region "Declarations"

#End Region

#Region "Public Instance"

#Region "Functions"

    Public Function DoUpdate() As Boolean
        Fields.GetRowId()
        If Fields.RowID.Trim.Length = 0 Then
            Fields.Insert(CConnection.DbConnection)
        Else
            Fields.Update(CConnection.DbConnection, , True)
        End If
    End Function
    ''' <developer></developer>
    ''' <summary>
    ''' Create a copy of the report
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>02-May-2012 10:56</created><Author>CONCATENO\taylor</Author></revision></revisionHistory>
    Public Function CloneReport() As CRMenuItem
        Dim newReport As CRMenuItem = CRMenuItems.CreateMenuItem
        With Me
            newReport.Instance = .Instance
            newReport.MenuOrder = .MenuOrder
            newReport.MenuParent = .MenuParent
            newReport.MenuPath = .MenuPath
            newReport.MenuType = .MenuType
            newReport.Password = .Password
            newReport.Username = .Username
            DoUpdate()
        End With
        Return newReport
    End Function
#End Region

#Region "Procedures"

#End Region

#Region "Properties"

#End Region
#End Region


    '''<summary>
    ''' Constructor
    ''' </summary>
    Public Sub New()
        SetupFields()
    End Sub

    Public Sub New(ByVal menuid As String)
        MyClass.New()
        Me.MenuIdentity = menuid
        Me.Fields.Load(CConnection.DbConnection, "identity = '" & Me.MenuIdentity & "'")
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Output Crystal report info as XML 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [03/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function ToXML() As String
        Dim strRet As String = "<CrystalReport>"
        Me.Fields.XMLHeader = ""
        strRet += Me.Fields.ToXMLSchema()
        strRet += Me.Formulae.ToXML
        'Do Bound reports
        Dim strParaRet As String = CConnection.PackageStringList("lib_ccTool.FindCustomerSchedules", Me.MenuIdentity)
        If Not strParaRet Is Nothing AndAlso strParaRet.Trim.Length > 0 Then
            strRet += "<Usage>"
            Dim strParas As String() = strParaRet.Split(New Char() {","})
            Dim i As Integer
            For i = 0 To strParas.Length - 1
                strRet += "<CustRep>" & strParas.GetValue(i) & "</CustRep>"
            Next
            strRet += "</Usage>"
        End If
        strRet += "</CrystalReport>"
        Return strRet
    End Function

    '''<summary>
    ''' A collection of Parameters for the report
    ''' </summary>
    Public Property Formulae() As CrFormulaItems
        Get
            If myFormulae Is Nothing Then
                myFormulae = New CrFormulaItems(Me.objIdentity.Value)
                myFormulae.Load(MedConnection.Connection)
            End If
            Return myFormulae
        End Get
        Set(ByVal Value As CrFormulaItems)
            myFormulae = Value
        End Set
    End Property

    Private Sub SetupFields()
        myFields = New TableFields("CRYSTAL_REPORTS")

        objIdentity = New StringField("IDENTITY", "", 10, True)
        myFields.Add(objIdentity)

        objParent = New StringField("PARENT", "", 10)
        myFields.Add(objParent)
        objText = New StringField("MenuText", "", 20)
        myFields.Add(objText)
        objPath = New StringField("REPORTPATH", "", 100)
        myFields.Add(objPath)
        objOrder = New IntegerField("MENUORDER", 0)
        myFields.Add(objOrder)
        objType = New StringField("MENUTYPE", "", 4)
        myFields.Add(objType)
        myFields.Add(objInstance)
        myFields.Add(objUserName)
        myFields.Add(objPassword)

    End Sub
    Public Property Instance() As String
        Get
            Return objInstance.Value
        End Get
        Set(ByVal value As String)
            objInstance.Value = value
        End Set
    End Property


    Public Property Username() As String
        Get
            Return objUserName.Value
        End Get
        Set(ByVal value As String)
            objUserName.Value = value
        End Set
    End Property


    Public Property Password() As String
        Get
            Return objPassword.Value
        End Get
        Set(ByVal value As String)
            objPassword.Value = value
        End Set
    End Property
    '''<summary>
    ''' Access to the data row
    ''' </summary>
    Public Property Fields() As TableFields
        Get
            Return myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    '''<summary>
    ''' Handle (to Crystal Report)
    ''' </summary>
    Public Property Handle() As System.IntPtr
        Get
            Return myHandle
        End Get
        Set(ByVal Value As System.IntPtr)
            myHandle = Value
        End Set
    End Property


    Private myIsFormula As Boolean = True
    Public Property IsFormula() As Boolean
        Get
            Return myIsFormula
        End Get
        Set(ByVal value As Boolean)
            myIsFormula = value
        End Set
    End Property



    '''<summary>
    ''' Text that appears on the menu
    ''' </summary>
    Public Property MenuText() As String
        Get
            Return objText.Value
        End Get
        Set(ByVal Value As String)
            objText.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Position of the menu in the menu list 
    ''' </summary>
    Public Property MenuOrder() As Integer
        Get
            Return Me.objOrder.Value
        End Get
        Set(ByVal Value As Integer)
            objOrder.Value = Value
        End Set
    End Property

    '''<summary>
    ''' menu which is the parent of this menu (NULL if no parent)
    ''' </summary>
    Public Property MenuParent() As String
        Get
            Return Me.objParent.Value
        End Get
        Set(ByVal Value As String)
            objParent.Value = Value
        End Set
    End Property

    '''<summary>
    ''' ID of this automated report
    ''' </summary>
    Public Property MenuIdentity() As String
        Get
            Return Me.objIdentity.Value
        End Get
        Set(ByVal Value As String)
            objIdentity.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Type of Automated Report
    ''' </summary>
    Public Property MenuType() As String
        Get
            Return Me.objType.Value
        End Get
        Set(ByVal Value As String)
            objType.Value = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Property to return Report Type as an Enumeration value
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property ReportType() As ReportTypes
        Get
            If MenuType.Trim.Length > 0 Then
                Return [Enum].Parse(GetType(ReportTypes), Me.MenuType)
            Else
                Return 0
            End If
        End Get
    End Property

    '''<summary>
    ''' Path to report template.
    ''' </summary>
    ''' <remarks>The function of this field id dependent on the menu type setting</remarks>
    Public Property MenuPath() As String
        Get
            Return Me.objPath.Value
        End Get
        Set(ByVal Value As String)
            objPath.Value = Value
        End Set
    End Property

    Public Overloads Sub FillFormula(ByRef ParameterList As String, Optional ByVal PhraseId As String = "")
        Medscreen.LogAction("Loading formulae")
        Dim vntRet As Object
        Dim I As Integer
        Dim Crf As MedscreenLib.CRFormulaItem
        Dim strVal As String

        If Me.Formulae.Count <> 0 Then
            For I = 0 To Me.Formulae.Count - 1
                Medscreen.LogAction("Getting parameter " & I)
                Crf = Me.Formulae.Item(I)
                Select Case Crf.ParamType.ToUpper
                    Case "DATE"
                        Dim tmpDate As Date = DateSerial(Today.Year, Today.Month, 1)
                        If Crf.Value.ToUpper = "YEARSTART" Then
                            tmpDate = DateSerial(Today.Year, 1, 1)
                        ElseIf Crf.Value.ToUpper = "YEAREND" Then
                            tmpDate = DateSerial(Today.Year, 12, 31)
                        ElseIf Crf.Value.ToUpper = "MONTHEND" Then
                            tmpDate = DateSerial(Today.Year, Today.Month, 1).AddMonths(1).AddDays(-1)
                        ElseIf Crf.Value.ToUpper = "PREVMONTH" Then
                            tmpDate = DateSerial(Today.Year, Today.Month, 1).AddMonths(-1)

                        End If
                        vntRet = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typDate, Crf.Formula, , tmpDate)
                    Case "STRING"
                        vntRet = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typString, Crf.Formula)
                    Case "INTEGER"
                        vntRet = MedscreenLib.Medscreen.GetParameter(MedscreenLib.Medscreen.MyTypes.typeInteger, Crf.Formula)
                    Case "PHRASEID"
                        vntRet = PhraseId
                    Case "SMID"
                        vntRet = PhraseId
                    Case Else
                        vntRet = Crf.ParamType
                End Select
                If vntRet Is Nothing Then
                    Dim ex As Exception = New Exception("No Parameter Entered")
                    Throw ex
                Else
                    strVal = vntRet
                End If
                ParameterList += " " & Crf.Formula.ToUpper & "=" & strVal
                'FillFormulaDet(cr, Crf, strVal)
            Next

        End If

    End Sub
    ''' <developer></developer>
    ''' <summary>
    ''' Log in and obtain as copy of the report
    ''' </summary>
    ''' <param name="ReportSource"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision>Moved to Crystal Support Library</revision></revisionHistory>
    <CLSCompliant(False)> _
    Public Function LogInCrystal(ByVal ReportSource As String, Optional ByVal noLogin As Boolean = False) As CrystalDecisions.CrystalReports.Engine.ReportDocument
        'Dim objTab As CrystalDecisions.CrystalReports.Engine.Table
        Dim Cr As CrystalDecisions.CrystalReports.Engine.ReportDocument = Nothing
        Cr = CrystalSupport.LogInCrystal(ReportSource, Me.MenuPath, Me.Username, Me.Password, Me.Instance, noLogin)


        Return Cr
    End Function


    ''' <developer></developer>
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><modified>03-Apr-2012 15:02</modified><Author>CONCATENO\taylor</Author>
    ''' Last working day and first working day options added</revision></revisionHistory>
    <CLSCompliant(False)> _
Public Function FormulaList(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument) As String
        Medscreen.LogAction("Loading formulae")
        Dim vntRet As Object
        Dim I As Integer
        Dim Crf As MedscreenLib.CRFormulaItem
        Dim strVal As String
        Dim strRet As String = ""
        If Me.Formulae.Count <> 0 Then
            For I = 0 To Me.Formulae.Count - 1

                Medscreen.LogAction("Getting parameter " & I)
                Crf = Me.Formulae.Item(I)
                If Crf.Value.Trim.Length > 0 Then
                    Select Case Crf.ParamType
                        Case "DATE"
                            'default to start of month
                            Dim tmpDate As Date = DateSerial(Today.Year, Today.Month, 1)
                            If Crf.Value.ToUpper = "YEARSTART" Then
                                tmpDate = DateSerial(Today.Year, 1, 1)
                            ElseIf Crf.Value.ToUpper = "YEAREND" Then
                                tmpDate = DateSerial(Today.Year, 12, 31)
                            ElseIf Crf.Value.ToUpper = "PREVMONTH" Then
                                tmpDate = tmpDate.AddMonths(-1)
                            ElseIf Crf.Value.ToUpper = "NEXTMONTH" Then
                                tmpDate = tmpDate.AddMonths(+1)
                            ElseIf Crf.Value.ToUpper = "MONTHEND" Then
                                tmpDate = DateSerial(Today.Year, Today.Month, 1).AddMonths(1).AddDays(-1)
                            ElseIf Crf.Value.ToUpper = "LASTWEEK" Then
                                tmpDate = tmpDate.AddDays(-7)
                            ElseIf Crf.Value.ToUpper = "NEXTWEEK" Then
                                tmpDate = tmpDate.AddDays(7)
                            ElseIf Crf.Value.ToUpper = "PREVFORTNT" Then
                                tmpDate = tmpDate.AddDays(-14)
                            ElseIf Crf.Value.ToUpper = "NEXTFORTNT" Then
                                tmpDate = tmpDate.AddDays(14)
                            ElseIf Crf.Value.ToUpper = "LASTWRKDAY" Then
                                tmpDate = tmpDate.AddMonths(+1).AddDays(-1)
                                While tmpDate.DayOfWeek = DayOfWeek.Saturday Or tmpDate.DayOfWeek = DayOfWeek.Sunday
                                    tmpDate = tmpDate.AddDays(-1)
                                End While
                            ElseIf Crf.Value.ToUpper = "FRSTWRKDAY" Then
                                tmpDate = tmpDate.AddMonths(+1)
                                While tmpDate.DayOfWeek = DayOfWeek.Saturday Or tmpDate.DayOfWeek = DayOfWeek.Sunday
                                    tmpDate = tmpDate.AddDays(1)
                                End While
                            End If
                            vntRet = tmpDate
                        Case "STRING"
                            vntRet = Crf.Value
                        Case "INTEGER"
                            vntRet = Crf.Value
                        Case "VALUE"
                            vntRet = Crf.Value
                        Case Else
                            vntRet = Crf.Value
                    End Select
                    If vntRet Is Nothing Then
                        Dim ex As Exception = New Exception("No Parameter Entered")
                        Throw ex
                    Else
                        strVal = vntRet
                    End If
                    strRet += Crf.Formula & "=" & vntRet & " "
                End If
            Next

        End If
        Return strRet

    End Function

    <CLSCompliant(False)> _
Public Sub FillFormulaFromValues(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        Medscreen.LogAction("Loading formulae")
        Dim vntRet As Object
        Dim I As Integer
        Dim Crf As MedscreenLib.CRFormulaItem
        Dim strVal As String

        If Me.Formulae.Count <> 0 Then
            For I = 0 To Me.Formulae.Count - 1

                Medscreen.LogAction("Getting parameter " & I)
                Crf = Me.Formulae.Item(I)
                If Crf.Value.Trim.Length > 0 Then
                    Select Case Crf.ParamType
                        Case "DATE"
                            'default to start of month
                            Dim tmpDate As Date = DateSerial(Today.Year, Today.Month, 1)
                            If Crf.Value.ToUpper = "YEARSTART" Then
                                tmpDate = DateSerial(Today.Year, 1, 1)
                            ElseIf Crf.Value.ToUpper = "YEAREND" Then
                                tmpDate = DateSerial(Today.Year, 12, 31)
                            ElseIf Crf.Value.ToUpper = "PREVMONTH" Then
                                tmpDate = tmpDate.AddMonths(-1)
                            ElseIf Crf.Value.ToUpper = "NEXTMONTH" Then
                                tmpDate = tmpDate.AddMonths(+1)
                            ElseIf Crf.Value.ToUpper = "MONTHEND" Then
                                tmpDate = DateSerial(Today.Year, Today.Month, 1).AddMonths(1).AddDays(-1)
                            ElseIf Crf.Value.ToUpper = "LASTWEEK" Then
                                tmpDate = tmpDate.AddDays(-7)
                            ElseIf Crf.Value.ToUpper = "NEXTWEEK" Then
                                tmpDate = tmpDate.AddDays(7)
                            ElseIf Crf.Value.ToUpper = "PREVFORTNT" Then
                                tmpDate = tmpDate.AddDays(-14)
                            ElseIf Crf.Value.ToUpper = "NEXTFORTNT" Then
                                tmpDate = tmpDate.AddDays(14)
                            End If
                            vntRet = tmpDate
                        Case "STRING"
                            vntRet = Crf.Value
                        Case "INTEGER"
                            vntRet = Crf.Value
                        Case "VALUE"
                            vntRet = Crf.Value
                        Case Else
                            vntRet = Crf.Value
                    End Select
                    If vntRet Is Nothing Then
                        Dim ex As Exception = New Exception("No Parameter Entered")
                        Throw ex
                    Else
                        strVal = vntRet
                    End If
                    IsFormula = (Me.MenuType <> "RPTP")
                    If IsFormula Then
                        FillFormulaDet(cr, Crf, strVal)
                    Else
                        FillParamDet(cr, Crf, strVal)
                    End If
                End If
            Next

        End If

    End Sub

    <CLSCompliant(False)> _
    Public Sub FillFormulaFromValue(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument)
        Medscreen.LogAction("Loading formulae")
        Dim vntRet As Object
        Dim I As Integer
        Dim Crf As MedscreenLib.CRFormulaItem
        Dim strVal As String

        If Me.Formulae.Count <> 0 Then
            For I = 0 To Me.Formulae.Count - 1

                Medscreen.LogAction("Getting parameter " & I)
                Crf = Me.Formulae.Item(I)
                If Crf.Value.Trim.Length > 0 Then
                    Select Case Crf.ParamType
                        Case "DATE"
                            Dim tmpDate As Date = DateSerial(Today.Year, Today.Month, 1)
                            If Crf.Value.ToUpper = "YEARSTART" Then
                                tmpDate = DateSerial(Today.Year, 1, 1)
                            ElseIf Crf.Value.ToUpper = "YEAREND" Then
                                tmpDate = DateSerial(Today.Year, 12, 31)
                            ElseIf Crf.Value.ToUpper = "PREVMONTH" Then
                                tmpDate = tmpDate.AddMonths(-1)
                            ElseIf Crf.Value.ToUpper = "NEXTMONTH" Then
                                tmpDate = tmpDate.AddMonths(+1)
                            ElseIf Crf.Value.ToUpper = "MONTHEND" Then
                                tmpDate = DateSerial(Today.Year, Today.Month, 1).AddMonths(1).AddDays(-1)
                            ElseIf Crf.Value.ToUpper = "LASTWEEK" Then
                                tmpDate = tmpDate.AddDays(-7)
                            ElseIf Crf.Value.ToUpper = "NEXTWEEK" Then
                                tmpDate = tmpDate.AddDays(7)
                            ElseIf Crf.Value.ToUpper = "LASTFORTNT" Then
                                tmpDate = tmpDate.AddDays(-14)
                            ElseIf Crf.Value.ToUpper = "NEXTFORTNT" Then
                                tmpDate = tmpDate.AddDays(14)
                            End If
                            vntRet = tmpDate
                        Case "STRING"
                            vntRet = Crf.Value
                        Case "INTEGER"
                            vntRet = Crf.Value
                        Case Else
                            vntRet = Crf.Value
                    End Select
                    If vntRet Is Nothing Then
                        Dim ex As Exception = New Exception("No Parameter Entered")
                        Throw ex
                    Else
                        strVal = vntRet
                    End If
                    'check to see if by parameter or formula
                    If Me.IsFormula Then
                        FillFormulaDet(cr, Crf, strVal)
                    Else
                        FillParamDet(cr, Crf, strVal)
                    End If
                End If
            Next

        End If

    End Sub


    Private Sub FillFormulaDet(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
      ByVal Crf As MedscreenLib.CRFormulaItem, ByVal strVal As String)
        Dim FF As CrystalDecisions.CrystalReports.Engine.FormulaFieldDefinition
        For Each FF In cr.DataDefinition.FormulaFields
            If FF.FormulaName.ToUpper = "{@" & Crf.Formula.ToUpper & "}" Then
                FF.Text = """" & strVal & """"
            End If
        Next

    End Sub

    Private Sub FillParamDet(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, _
      ByVal Crf As MedscreenLib.CRFormulaItem, ByVal strVal As String)
        Dim FF As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
        For Each FF In cr.DataDefinition.ParameterFields
            If Not FF.IsLinked And (FF.Name = Crf.Formula OrElse FF.Name = Crf.FieldName) Then
                cr.SetParameterValue(FF.Name, Crf.Value)

            End If
        Next

    End Sub
End Class

'''<summary>
''' a Collection of automated reports
''' </summary>
Public Class CRMenuItems
    Inherits CollectionBase
#Region "Declarations"
    Private objIdentity As StringField
    Private objParent As StringField
    Private objText As StringField
    Private objPath As StringField
    Private objOrder As IntegerField
    Private objType As StringField
    Private objInstance As StringField = New StringField("INSTANCE", "", 100)
    Private objUserName As StringField = New StringField("USERNAME", "", 40)
    Private objPassword As StringField = New StringField("PASSWORD", "", 40)

    Private myFields As TableFields
#End Region
    ''' <developer></developer>
    ''' <summary>
    ''' Create a new menu item
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Shared Function CreateMenuItem() As CRMenuItem
        Dim objRet As CRMenuItem = New CRMenuItem()
        Dim newId As String = CConnection.NextSequence("SEQ_CRYSTAL_REPORTS")
        objRet.MenuIdentity = newId
        objRet.Fields.Insert(CConnection.DbConnection)
        Return objRet

    End Function

    '''<summary>
    ''' Create a new Automated report collection
    ''' </summary>
    Public Sub New()
        MyBase.New()
        SetupFields()
    End Sub

    '''<summary>
    ''' Add a Windows Menu
    ''' </summary>
    ''' <param name='MnuItem'>Menu Item to add</param>
    ''' <param name='Handler'>Windows Event Handler</param>
    ''' <param name='CSVHandler'>Windows Event Handler for CSV reports</param>
    Public Sub AddMenus(ByVal MnuItem As Windows.Forms.MenuItem, _
        ByVal Handler As System.EventHandler, _
    ByVal CSVHandler As System.EventHandler, Optional ByVal Parent As String = ".")
        Medscreen.LogAction("Building Menus")
        Dim I As Integer
        Dim crMenu As CRMenuItem
        Dim crParent As CRMenuItem
        Dim mnuX As Windows.Forms.MenuItem

        For I = 0 To Me.Count - 1
            crMenu = Me.Item(I)
            If crMenu.MenuType = "MNU" AndAlso crMenu.MenuParent = Parent Then
                mnuX = MnuItem.MenuItems.Add(crMenu.MenuText)
                crMenu.Handle = mnuX.Handle
            End If
            If crMenu.MenuType = "BAR" Then
                crParent = Me.Item(crMenu.MenuIdentity)
                If Not crParent Is Nothing Then
                    mnuX = FindMenu(crParent.Handle, MnuItem)
                    If Not mnuX Is Nothing Then
                        mnuX = mnuX.MenuItems.Add("-")
                        crMenu.Handle = mnuX.Handle
                    End If
                End If

            ElseIf crMenu.MenuType = "RPT" Then
                crParent = Me.Item(crMenu.MenuParent)
                If Not crParent Is Nothing Then
                    mnuX = FindMenu(crParent.Handle, MnuItem)

                    If Not mnuX Is Nothing Then
                        mnuX = mnuX.MenuItems.Add(crMenu.MenuText, Handler)
                        'Medscreen.LogAction("menu handler added - " & crMenu.MenuText)
                        crMenu.Handle = mnuX.Handle
                    End If
                End If
            ElseIf (crMenu.MenuType = "RPTX" Or crMenu.MenuType = "CSV") Then
                crParent = Me.Item(crMenu.MenuParent)
                If Not crParent Is Nothing Then
                    mnuX = FindMenu(crParent.Handle, MnuItem)

                    If Not mnuX Is Nothing Then
                        mnuX = mnuX.MenuItems.Add(crMenu.MenuText, CSVHandler)
                        crMenu.Handle = mnuX.Handle
                    End If
                End If
            End If

        Next

    End Sub


    '''<summary>
    ''' Find a menu
    ''' </summary>
    ''' <param name='objHandle'>Windows handle of the menu to find</param>
    ''' <param name='Item'>MenuItem to find</param>
    ''' <returns>Menu Item</returns>
    Private Function FindMenu(ByVal objHandle As System.IntPtr, ByVal Item As Windows.Forms.MenuItem) As Windows.Forms.MenuItem
        Dim mnuX As Windows.Forms.MenuItem = Nothing

        For Each mnuX In Item.MenuItems
            If mnuX.Handle.Equals(objHandle) Then
                Exit For
            End If
            If mnuX.IsParent Then
                mnuX = FindMenu(objHandle, mnuX)
                If Not mnuX Is Nothing Then
                    Exit For
                End If
            End If
            mnuX = Nothing
        Next
        Return mnuX
    End Function


    '''<summary>
    ''' Add a menu item
    ''' </summary>
    ''' <param name='Item'>Item to be added</param>
    ''' <returns>Position in List</returns>
    Public Function Add(ByVal Item As CRMenuItem) As Integer
        Return MyBase.List.Add(Item)
    End Function

    '''<summary>
    ''' get menu item
    ''' </summary>
    ''' <param name='index'>Handle of item to be found</param>
    ''' <returns>MenuItem</returns>
    Public Overloads Function Item(ByVal index As System.IntPtr) As CRMenuItem
        Dim i As Integer
        Dim objCr As CRMenuItem = Nothing

        Medscreen.LogAction(Me.Count)
        For i = 0 To Count - 1
            objCr = Me.List(i)
            Medscreen.LogAction(objCr.Handle.ToString & " - " & index.ToString)
            If objCr.Handle.Equals(index) Then
                Exit For

            End If
            objCr = Nothing
        Next
        Return objCr

    End Function

    '''<summary>
    ''' get menu item
    ''' </summary>
    ''' <param name='index'>Identity of item to be found</param>
    ''' <returns>MenuItem</returns>
    Public Overloads Function Item(ByVal index As String, Optional ByVal by As Integer = 0) As CRMenuItem
        Dim i As Integer
        Dim objCr As CRMenuItem = Nothing

        For i = 0 To Count - 1
            objCr = Me.List(i)
            If by = 0 Then
                If objCr.MenuIdentity = index Then
                    Exit For
                End If
            ElseIf by = 1 Then
                Medscreen.LogAction(objCr.MenuText & " - " & index)
                If objCr.MenuText = index Then
                    Exit For

                End If
            End If
            objCr = Nothing
        Next
        Return objCr

    End Function

    '''<summary>
    ''' get menu item
    ''' </summary>
    ''' <param name='index'>Position of item to be found</param>
    ''' <returns>MenuItem</returns>
    Public Overloads Function Item(ByVal index As Integer) As CRMenuItem
        Return CType(MyBase.List.Item(index), CRMenuItem)
    End Function

    Private Sub SetupFields()
        myFields = New TableFields("CRYSTAL_REPORTS")

        objIdentity = New StringField("IDENTITY", "", 10, True)
        myFields.Add(objIdentity)
        objParent = New StringField("PARENT", "", 10)
        myFields.Add(objParent)
        objText = New StringField("MenuText", "", 20)
        myFields.Add(objText)
        objPath = New StringField("REPORTPATH", "", 100)
        myFields.Add(objPath)
        objOrder = New IntegerField("MENUORDER", 0)
        myFields.Add(objOrder)
        objType = New StringField("MENUTYPE", "", 4)
        myFields.Add(objType)
        myFields.Add(objInstance)
        myFields.Add(objUserName)
        myFields.Add(objPassword)

    End Sub

    Public Function LoadXML(ByVal Filename As String) As Boolean
        Dim ds As New DataSet()
        Me.Clear()
        ds.ReadXml(Filename)
        'Dim menuTable As ReportMenus.rowDataTable = ds.Tables(0)
        Dim r As DataRow
        For Each r In ds.Tables(0).Rows
            Dim crMenu As New CRMenuItem()
            crMenu.Fields.readfields(r)
            Me.Add(crMenu)
        Next r

    End Function

    '''<summary>
    ''' Load Automated reports from Database
    ''' </summary>
    ''' <returns>TRUE if succesful</returns>
    Public Function Load() As Boolean
        Dim oCmd As New OleDb.OleDbCommand()
        Dim oRead As OleDb.OleDbDataReader = Nothing
        Dim Cr As CRMenuItem


        Try
            oCmd.CommandText = "Select * from Crystal_reports order by parent,menuorder"
            oCmd.Connection = CConnection.DbConnection
            If CConnection.ConnOpen Then


                oRead = oCmd.ExecuteReader
                While oRead.Read
                    myFields.readfields(oRead)
                    Cr = New CRMenuItem()
                    Cr.MenuIdentity = myFields.Item("IDENTITY").Value
                    Cr.Fields.Fill(myFields)
                    Me.Add(Cr)
                End While
            End If

        Catch ex As OleDb.OleDbException
            If InStr(ex.Message, "ORA-125") Then
                'MsgBox("oracle is not available this application will terminate" & vbCrLf & _
                '    ex.Message, MsgBoxStyle.Critical)
                Throw New MedscreenExceptions.OracleFailure("Oracle is not available this application will terminate" & vbCrLf & _
                    ex.Message)
            End If
            Medscreen.LogError(ex)
        Catch ex As Exception
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            CConnection.SetConnClosed()
        End Try

    End Function
End Class

'''<summary>
''' Wrapper to Net Chat
''' </summary>
Public Class NetChat
#Region "Declarations"


    Private userIni As IniFile.IniFiles
    Private users As IniFile.IniCollection
    Private w As IO.StreamWriter
#End Region

    '''<summary>
    ''' Create New Net Chat Message
    ''' </summary>
    ''' <param name='Message'>Message to send</param>
    ''' <param name='UserList'>List of user to send message to</param>
    ''' <param name='Top'>Send to top N users</param>
    Public Sub New(ByVal Message As String, ByVal UserList As String, _
    ByVal Top As Integer)
        Dim v As IniFile.IniElement
        userIni = New IniFile.IniFiles()
        userIni.FileName = Medscreen.LiveRoot & "inifiles\spawnchat.ini"
        users = userIni.ReadSectionCollection(UserList)

        w = New IO.StreamWriter(Medscreen.LiveRoot & "userout\NC-" & Now.ToString("yyyyMMdd") & ".NCM")
        w.WriteLine(Message)
        w.WriteLine()
        w.WriteLine("[UserList]=" & Top.ToString.Trim)
        For Each v In users
            w.WriteLine(v.Header)
        Next

        w.Flush()
        w.Close()



    End Sub

End Class

'''<summary>
''' Class to manage parameters for automated reports
''' </summary>
Public Class CrFormulaItems
    Inherits CollectionBase
#Region "Declarations"

    Private objIdentity As StringField
    Private objFormula As StringField
    Private objParamType As StringField
    Private objFieldName As StringField
    Private objFieldValue As StringField

    Private myFields As TableFields
    Private myIdentity As String
#End Region

    '''<summary>
    ''' Create a new parameter list 
    ''' </summary>
    ''' <param name='Identity'>Identity of Automated report to create parameter list for</param>
    Public Sub New(ByVal Identity As String)
        MyBase.New()

        If Identity.Trim.Length < 10 Then
            Identity = Identity.PadLeft(10, "0")
        End If
        myIdentity = Identity
        myFields = New TableFields("crystal_formula")
        objIdentity = New StringField("IDENTITY", "", 10, True)
        myFields.Add(objIdentity)

        objFormula = New StringField("FORMULANAME", "", 30, True)
        myFields.Add(objFormula)

        objParamType = New StringField("PARAMTYPE", "", 10)
        myFields.Add(objParamType)

        objFieldName = New StringField("COLUMNNAME", "", 30)
        myFields.Add(objFieldName)
        objFieldValue = New StringField("VALUE", "", 30)
        myFields.Add(objFieldValue)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Expose index of method
    ''' </summary>
    ''' <param name="Item"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IndexOf(ByVal Item As CRFormulaItem) As Integer
        Return MyBase.List.IndexOf(Item)
    End Function

    '''<summary>
    ''' add new parameter to list 
    ''' </summary>
    ''' <param name='CR'>Parameter to add</param>
    ''' <returns>Position in list</returns>
    Public Function Add(ByVal CR As CRFormulaItem) As Integer
        Return MyBase.List.Add(CR)
    End Function

    '''<summary>
    ''' get Item from list  
    ''' </summary>
    ''' <param name='index'>Position in list to get</param>
    ''' <returns>Parameter</returns>
    Public Function Item(ByVal index As Integer) As CRFormulaItem
        Return CType(MyBase.List.Item(index), CRFormulaItem)
    End Function

    '''<summary>
    ''' Load parameter list 
    ''' </summary>
    ''' <param name='oConn'>OLEDB connector</param>
    ''' <returns>TRUE if succesful</returns>
    Public Function Load(ByVal oConn As OleDb.OleDbConnection) As Boolean
        Dim oCmd As New OleDb.OleDbCommand()
        Dim oRead As OleDb.OleDbDataReader
        Dim Cr As CRFormulaItem

        If Me.Count > 0 Then Exit Function
        Try
            oCmd.CommandText = myFields.FullRowSelect & " where identity = ?"
            oCmd.Parameters.Add(CConnection.StringParameter("Identity", myIdentity, 10))
            oCmd.Connection = oConn
            oConn.Open()

            oRead = oCmd.ExecuteReader
            While oRead.Read
                myFields.readfields(oRead)
                Cr = New CRFormulaItem()
                Cr.Identity = myFields.Item("IDENTITY").Value
                Cr.Formula = myFields.Item("FORMULANAME").Value
                Cr.ParamType = myFields.Item("PARAMTYPE").Value
                Cr.Fields.readfields(oRead)
                Me.Add(Cr)
            End While
            oRead.Close()
        Catch ex As Exception
        Finally
            oConn.Close()
        End Try

    End Function

    Public Function ExportPath() As String
        Dim strPath As String = ""
        For Each CFForm As CRFormulaItem In MyBase.InnerList
            If CFForm.Formula = "PATH" Then
                strPath = CFForm.Value
                Exit For

            End If
        Next
        Return strPath
    End Function


    Public Function ToXMLElement() As XElement
        Dim Element As New XElement("Parameters")
        For Each par As CRFormulaItem In MyBase.InnerList
            Element.Add(par.ToXMLElement)
        Next
        Return Element
    End Function

    Public Function ToXMLElementFull() As XElement
        Dim Element As XElement = ToXMLElement()
        Return Element
    End Function


    '''<summary>
    ''' Converts parameters into XML 
    ''' </summary>
    ''' <returns>Parameter list a sXML</returns>
    Public Function ToXML() As String
        Dim strXML As String = ""
        Dim oCRFrmIt As CRFormulaItem

        Dim i As Integer
        For i = 0 To Count - 1
            oCRFrmIt = Item(i)
            strXML += oCRFrmIt.ToXML()

        Next
        Return strXML

    End Function

End Class

' 

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : CRFormulaItem
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' An object to deal with a single parameter for a report
''' </summary>
''' <remarks>Firstly there is no restriction on the use these fields can be put to. The current use is as follows :- <para/>
''' <see cref="CRFormulaItem.ParamType"/> this is used to desc what type of parameter this is, it is available via an enumeration, if you 
''' want to add new values, please ensure the enumeration gets updated.<para/>
''' <see cref="CRFormulaItem.Formula"/>Originally this pointed to a Crystal Report Formula name, 
''' again being consistent with the formulae names is advantageous.  It may also be 
''' used as a HTML report parameter name.<para/>
''' <see cref="CRFormulaItem.FieldName"/> The name of the field that this parameter can be found 
''' in, this may be fully qualified i.e. TABLE.FIELD<para/>
''' <see cref="CRFormulaItem.Value"/>The value field contains anything suitable for the parameter.
''' 
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [30/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class CRFormulaItem

#Region "Declarations"
    Private objIdentity As StringField
    Private objFormula As StringField
    Private objParamType As StringField
    Private objFieldName As StringField
    Private objFieldValue As StringField

    Private myFields As TableFields
#End Region

#Region "Enumerations"
    '''<summary>Types of Parameter possible</summary>
    Public Enum ParameterTypes
        '''<summary>Value will refer to a number of days</summary>
        typDAYS
        '''<summary>Field property will refer to a primary ID</summary>
        typID
        '''<summary>Types of Parameter possible</summary>
        typBOOLEAN
        '''<summary>Value property will be TRUE or FALSE</summary>
        typDATE
        '''<summary>Value property will refer to a physical or logical date</summary>
        typEMAIL
        '''<summary>Value property will be a filename, Formula will be a file type</summary>
        typFILENAME

        typPhraseId

    End Enum

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Types possible for the Formula name field
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [30/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum FormulaTypesx
        ''' <summary>Parameter is a location that the email address can be found</summary>
        EMAIL
        ''' <summary>Parameter gives the location where the end date for report can be found</summary>
        ENDDATE
        ''' <summary>Relates to a grouping parameter gives the value e.g. Group or not group, or way of grouping, e.g. Parent child, Sage ID</summary>
        GROUP
        ''' <summary>Location of the ID for the report</summary>
        ID
        ''' <summary>A parameter relating to repeat frequencies</summary>
        REPEAT
        ''' <summary>Typically relates to an XSL stylesheet, could also be a word template</summary>
        STYLESHEET
        ''' <summary>Gives the location where the start date can be found</summary>
        STARTDATE
        ''' <summary>Customer's SMID held in value field</summary>
        SMID
        ''' <summary>How report will be detailed</summary>
        DETAIL
        ''' <summary>How report will be sent</summary>
        SENDTYPE

        PATH
    End Enum

    Public Enum FieldTypes
        CUSTOMER_ID
        NEXT_REPORT
        PREV_REPORT
        RECIPIENTS
        VALUE
        SMID
        SENDTYPE
    End Enum
#End Region
    '''<summary>
    ''' Create a new Parameter Item
    '''</summary>
    Public Sub New()
        myFields = New TableFields("crystal_formula")
        objIdentity = New StringField("IDENTITY", "", 10, True)
        myFields.Add(objIdentity)

        objFormula = New StringField("FORMULANAME", "", 30, True)
        myFields.Add(objFormula)
        objParamType = New StringField("PARAMTYPE", "", 10)
        myFields.Add(objParamType)

        objFieldName = New StringField("COLUMNNAME", "", 30)
        myFields.Add(objFieldName)
        objFieldValue = New StringField("VALUE", "", 30)
        myFields.Add(objFieldValue)

    End Sub


    '''<summary>
    ''' Field Name parameter refers to 
    '''</summary>
    Public Property FieldName() As String
        Get
            Return Me.objFieldName.Value
        End Get
        Set(ByVal Value As String)
            Me.objFieldName.Value = Value
        End Set
    End Property


    Friend Property Fields() As TableFields
        Get
            Return Me.myFields
        End Get
        Set(ByVal Value As TableFields)
            Me.myFields = Value
        End Set
    End Property

    '''<summary>
    ''' Formula used 
    '''</summary>
    ''' <remarks>In crystal reports parameters are named formulae in the report</remarks>
    Public Property Formula() As String
        Get
            Return objFormula.Value
        End Get
        Set(ByVal Value As String)
            objFormula.Value = Value
        End Set
    End Property


    '''<summary>
    ''' Identity of Parent report
    '''</summary>
    Public Property Identity() As String
        Get
            Return Me.objIdentity.Value
        End Get
        Set(ByVal Value As String)
            Me.objIdentity.Value = Value
        End Set
    End Property

    Private cr As CrystalDecisions.CrystalReports.Engine.ReportDocument
    <CLSCompliant(False)> _
    Public Property ParentReport() As CrystalDecisions.CrystalReports.Engine.ReportDocument
        Get
            Return cr
        End Get
        Set(ByVal Value As CrystalDecisions.CrystalReports.Engine.ReportDocument)
            cr = Value
        End Set
    End Property

    Private strReportName As String = ""
    Public Property ReportName() As String
        Get
            Return strReportName
        End Get
        Set(ByVal Value As String)
            strReportName = Value
        End Set
    End Property

    Private strlHTML As String
    Public ReadOnly Property HTML() As String
        Get
            Return STRlHTML
        End Get
    End Property

    Private strEmailAddress As String = ""
    Public Property EmailAddress() As String
        Get
            Return strEmailAddress
        End Get
        Set(ByVal Value As String)
            strEmailAddress = Value
        End Set
    End Property

    Public Sub DoReport()
        If Me.Formula = "SENDTYPE" AndAlso Me.Value <> "PDF" Then
            If Me.Value = "HTML" Then
                Dim tmpFileName As String = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "HTM")
                Dim expOpt As CrystalDecisions.Shared.ExportOptions = cr.ExportOptions
                expOpt.ExportFormatType = CrystalDecisions.[Shared].ExportFormatType.HTML40
                expOpt.ExportDestinationType = CrystalDecisions.[Shared].ExportDestinationType.DiskFile
                Dim diskexport As New CrystalDecisions.Shared.DiskFileDestinationOptions()
                diskexport.DiskFileName = tmpfilename
                expOpt.DestinationOptions = diskexport
                cr.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpfilename) 'CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpFileName)
                Try  ' Protecting file load 
                    Dim iof As New IO.StreamReader(tmpfilename)
                    Dim strHTML As String = iof.ReadToEnd
                    iof.Close()
                    strlHTML = strHTML
                    If EmailAddress.Trim.Length > 0 Then
                        Medscreen.BlatEmail("Report - " & ReportName, strHTML, EmailAddress)

                    End If
                Catch ex As Exception
                    MedscreenLib.Medscreen.LogError(ex, , "TestPanelForm-myReportHandler-2280")
                Finally
                End Try
            End If
        Else
            Dim tmpFileName As String = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "PDF")
            Dim expOpt As CrystalDecisions.Shared.ExportOptions = cr.ExportOptions
            expOpt.ExportFormatType = CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat

            expOpt.ExportDestinationType = CrystalDecisions.[Shared].ExportDestinationType.DiskFile
            Dim diskexport As New CrystalDecisions.Shared.DiskFileDestinationOptions()
            diskexport.DiskFileName = tmpfilename
            expOpt.DestinationOptions = diskexport

            cr.ExportToDisk(CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, tmpfilename) 'CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpFileName)
            '(CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, tmpFileName)
            Dim Attach As String() = {" "}
            Attach.SetValue(tmpFileName, 0)
            Medscreen.BlatEmail("Report - " & ReportName, "Please find your report enclosed", EmailAddress, , , , tmpfilename)
            Threading.Thread.Sleep(1000)
            IO.File.Delete(tmpFileName)
        End If

    End Sub
    '''<summary>
    ''' Type of Parameter
    '''</summary>
    ''' <remarks></remarks>
    Public Overloads Property ParamType() As String
        Get
            Return Me.objParamType.Value
        End Get
        Set(ByVal Value As String)
            objParamType.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Type of Parameter expressed as ParameterType <see cref="ParameterTypes"/>
    '''</summary>
    ''' <remarks></remarks>
    ''' <seealso cref="ParameterTypes"/>
    Public Overloads ReadOnly Property ParamTypeConstant() As ParameterTypes
        Get
            Dim RetPtype As ParameterTypes
            If Me.ParamType.Trim.Length > 0 Then
                Dim strPType As String = "TYP" & ParamType.ToUpper
                Dim iPType As String

                For Each iPType In System.Enum.GetNames(GetType(ParameterTypes))
                    If iPType.ToUpper = strPType Then
                        RetPtype = [Enum].Parse(GetType(ParameterTypes), strPType)
                        Exit For
                    End If

                Next

            End If
            Return RetPtype

        End Get
    End Property

    Public Function ToXMLElement() As XElement
        Dim Element As New XElement("Parameter")
        Element.Add(New XElement("formulaname", Formula))
        Element.Add(New XElement("paramtype", ParamType))
        Element.Add(New XElement("fieldname", FieldName))
        Element.Add(New XElement("value", Value))
        Return Element
    End Function

    Public Function ToXMLElementFull() As XElement
        Dim Element As XElement = ToXMLElement()
        Return Element
    End Function

    '''<summary>
    ''' Convert to XML
    '''</summary>
    ''' <remarks></remarks>
    Public Function ToXML() As String
        Dim strRet As String = ToXMLElement.ToString ' "<Parameter>"
        'strRet += "<formulaname>" & Me.Formula & "</formulaname>"
        'strRet += "<paramtype>" & Me.ParamType & "</paramtype>"
        'strRet += "<fieldname>" & Me.FieldName & "</fieldname>"
        'strRet += "<value>" & Me.Value & "</value>"
        'strRet += "</Parameter>"
        Return strRet
    End Function

    '''<summary>
    ''' Return the value property
    '''</summary>
    ''' <remarks></remarks>
    Public Property Value() As String
        Get
            Return Me.objFieldValue.Value
        End Get
        Set(ByVal Value As String)
            Me.objFieldValue.Value = Value
        End Set
    End Property
    '''<summary>
    ''' Update Parameter in Database
    '''</summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks></remarks>
    Public Function Update() As Boolean
        If Me.Fields.RowID.Trim.Length = 0 Then
            Return Me.myFields.Insert(MedscreenLib.MedConnection.Connection)
        Else
            Return Me.myFields.Update(MedscreenLib.MedConnection.Connection, True, True)
        End If
    End Function

    Public Function Insert() As Boolean
        Return Me.myFields.Insert(MedscreenLib.MedConnection.Connection)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Delete item from database
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [22/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Delete() As Boolean
        Return Me.myFields.Delete(MedscreenLib.CConnection.DbConnection)
    End Function
End Class

'''<summary>
''' Row in the ICP Table
'''</summary>
''' <remarks>ICP is the credit card database and is part of the Commidea software</remarks>
Public Class ICPRow

#Region "Declarations"


    Private myFields As TableFields
    Private objTransId As IntegerField
    Private objUserName As StringField
    Private objTXNType As StringField
    Private objSchemeName As StringField
    Private objModifier As StringField
    Private objCardNumber As StringField
    Private objExpiry As StringField
    Private objValue As StringField
    Private objAuthCode As StringField
    Private objDateTime As StringField
    Private objEftSeqNum As StringField
    Private objRef As StringField
    Private objTxnResult As StringField
    Private objAuthMessage As StringField
    Private objCardHolder As StringField
    Private objInvoiceNumber As StringField = New StringField("InvoiceNumber", "", 10)
#End Region

    Private Sub SetupFields()
        myFields = New TableFields("Transactions")
        objTransId = New IntegerField("TransactionID", 0, True)
        myFields.Add(objTransId)
        objUserName = New StringField("UserName", "", 20)
        myFields.Add(objUserName)
        objTXNType = New StringField("TxnType", "", 20)
        myFields.Add(objTXNType)
        objSchemeName = New StringField("SchemeName", "", 20)
        myFields.Add(objSchemeName)
        Me.objModifier = New StringField("Modifier", "", 20)
        myFields.Add(Me.objModifier)
        Me.objCardNumber = New StringField("CardNumber", "", 30)
        myFields.Add(Me.objCardNumber)

        Me.objExpiry = New StringField("Expiry", "", 20)
        myFields.Add(Me.objExpiry)
        Me.objValue = New StringField("TxnValue", "", 20)
        myFields.Add(Me.objValue)
        Me.objAuthCode = New StringField("AuthCode", "", 20)
        myFields.Add(Me.objAuthCode)
        Me.objDateTime = New StringField("DateTime", "", 20)
        myFields.Add(Me.objDateTime)
        Me.objEftSeqNum = New StringField("EFTSeqNum", "", 20)
        myFields.Add(Me.objEftSeqNum)
        Me.objRef = New StringField("Referance", "", 100)
        myFields.Add(Me.objRef)
        Me.objTxnResult = New StringField("TxnResult", "", 100)
        myFields.Add(Me.objTxnResult)
        Me.objAuthMessage = New StringField("AuthMessage", "", 100)
        myFields.Add(Me.objAuthMessage)
        Me.objCardHolder = New StringField("CardholderName", "", 80)
        myFields.Add(Me.objCardHolder)
        myFields.Add(objInvoiceNumber)

    End Sub

    '''<summary>
    ''' Create a new ICP Row Entry
    '''</summary>
    ''' <remarks></remarks>
    Public Sub New()
        SetupFields()
    End Sub

    Friend Property Fields() As TableFields
        Get
            Return myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    '''<summary>
    ''' ID of the Transaction (Primary Key)
    '''</summary>
    ''' <remarks></remarks>
    Public Property TransactionId() As Integer
        Get
            Return Me.objTransId.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objTransId.Value = Value
        End Set
    End Property

    Public Property InvoiceNumber() As String
        Get
            Return Me.objInvoiceNumber.Value
        End Get
        Set(ByVal Value As String)
            objInvoiceNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Transaction Type
    '''</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property TxnType() As String
        Get
            Return Me.objTXNType.Value
        End Get
    End Property

    '''<summary>
    ''' SchemeName
    '''</summary>
    ''' <remarks></remarks>
    Public Property SchemeName() As String
        Get
            Return Me.objSchemeName.Value
        End Get
        Set(ByVal Value As String)
            Me.objSchemeName.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Modifier
    '''</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property Modifier() As String
        Get
            Return Me.objModifier.Value
        End Get
    End Property

    '''<summary>
    ''' Card Number (Packed)
    '''</summary>
    ''' <remarks></remarks>
    Public Property Cardnumber() As String
        Get
            Return Me.objCardNumber.Value
        End Get
        Set(ByVal Value As String)
            Me.objCardNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Card Number (Formatted)
    '''</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property FormattedCardnumber() As String
        Get
            Dim objCno As String = Me.objCardNumber.Value
            Dim intLen As Integer = Len(objCno)
            If intLen = 16 Then
                Return Mid(objCno, 1, 4) & "-" & Mid(objCno, 5, 4) & _
                    "-" & Mid(objCno, 9, 4) & "-" & Mid(objCno, 13, 4)
            ElseIf intLen = 15 Then
                Return Mid(objCno, 1, 4) & "-" & Mid(objCno, 5, 7) & _
                    "-" & Mid(objCno, 12, 4)
            ElseIf intLen = 14 Then
                Return Mid(objCno, 1, 4) & "-" & Mid(objCno, 5, 6) & _
                    "-" & Mid(objCno, 11, 4)
            ElseIf intLen = 13 Then
                Return Mid(objCno, 1, 4) & "-" & Mid(objCno, 5, 3) & _
                    "-" & Mid(objCno, 8, 3) & "-" & Mid(objCno, 11, 3)
            Else
                Return objCno
            End If
        End Get
    End Property

    '''<summary>
    ''' Expiry(4 Character)
    '''</summary>
    ''' <remarks></remarks>
    Public Property Expiry() As String
        Get
            Return Me.objExpiry.Value
        End Get

        Set(ByVal Value As String)
            Me.objExpiry.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Transaction Value
    '''</summary>
    ''' <remarks></remarks>
    Public Property TxnValue() As String
        Get
            Return Me.objValue.Value
        End Get
        Set(ByVal Value As String)
            objValue.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Authorisation Code
    '''</summary>
    ''' <remarks></remarks>
    Public Property AuthCode() As String
        Get
            Return Me.objAuthCode.Value
        End Get
        Set(ByVal Value As String)
            Me.objAuthCode.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Return Transaction Date as a date time
    '''</summary>
    ''' <remarks></remarks>
    Public Property TransactionDate() As DateTime

        Get
            If DateTime.Length < 12 Then
                Return DateSerial(0, 0, 0)
                Exit Property
            End If
            Dim sdate As String = Mid(DateTime, 1, 4)
            Dim sMonth As String = Mid(DateTime, 5, 2)
            Dim sDay As String = Mid(DateTime, 7, 2)
            Dim sHour As String = Mid(DateTime, 9, 2)
            Dim sMin As String = Mid(DateTime, 11, 2)
            Dim sSec As String = Mid(DateTime, 13, 2)

            Dim td As DateTime
            td = DateSerial(sdate, sMonth, sDay) + " " + TimeSerial(sHour, sMin, sSec)
            Return td
        End Get
        Set(ByVal Value As DateTime)
            DateTime = Value.ToString("yyyyMMddHHmmss")

        End Set
    End Property

    '''<summary>
    ''' Transaction Date as a string
    '''</summary>
    ''' <remarks></remarks>
    Public Property DateTime() As String
        Get
            Return Me.objDateTime.Value
        End Get
        Set(ByVal Value As String)
            Me.objDateTime.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Sequence Number
    '''</summary>
    ''' <remarks></remarks>
    Public Property EFTSeqNum() As String
        Get
            Return Me.objEftSeqNum.Value
        End Get
        Set(ByVal Value As String)
            Me.objEftSeqNum.Value = Value
        End Set
    End Property

    '''<summary>
    ''' User entered comments
    '''</summary>
    ''' <remarks></remarks>
    Public Property Reference() As String
        Get
            Return Me.objRef.Value
        End Get
        Set(ByVal Value As String)
            Me.objRef.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Medscreen reference form C0XXX
    '''</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property MedscreenReference() As String
        Get
            Dim intPos As Integer = InStr(Me.Reference.ToUpper, "C0")
            If intPos > 0 Then
                Return Mid(Me.Reference, intPos, 5)
            Else
                Return ""
            End If
        End Get
    End Property

    Private strCmNumber As String = ""
    '''<summary>
    ''' Medscreen reference form CM
    '''</summary>
    ''' <remarks></remarks>
    Public Property CMJobNumber() As String
        Get
            Dim intPos As Integer = InStr(Me.Reference.ToUpper, "CM")

            If intPos > 0 Then
                strCmNumber = Mid(Me.Reference, intPos + 2).Trim
                intPos = InStr(strCmNumber, " ")
                If intPos > 0 Then 'other info on line
                    strCmNumber = Mid(strCmNumber, 1, intPos - 1)
                End If
            Else
                Return strCmNumber
            End If
            Return strCmNumber
        End Get
        Set(ByVal Value As String)
            Me.strCmNumber = Value
        End Set
    End Property

    '''<summary>
    ''' Transaction Result
    '''</summary>
    ''' <remarks></remarks>
    Public ReadOnly Property TxnResult() As String
        Get
            Return Me.objTxnResult.Value
        End Get
    End Property

    '''<summary>
    ''' Authorisation message
    '''</summary>
    ''' <remarks></remarks>
    Public Property AuthMessage() As String
        Get
            Return Me.objAuthMessage.Value
        End Get
        Set(ByVal Value As String)
            Me.objAuthMessage.Value = Value
        End Set
    End Property

    '''<summary>
    ''' User Entering data
    '''</summary>
    ''' <remarks></remarks>
    Public Property UserName() As String
        Get
            Return Me.objUserName.Value
        End Get
        Set(ByVal Value As String)
            Me.objUserName.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Card Holder
    '''</summary>
    ''' <remarks></remarks>
    Public Property CardHolder() As String
        Get
            Return Me.objCardHolder.Value
        End Get
        Set(ByVal Value As String)
            Me.objCardHolder.Value = Value
        End Set
    End Property
End Class
'''<summary>
''' ICP Table
'''</summary>
''' <remarks>ICP is the credit card database and is part of the Commidea software</remarks>
Public Class ICP
    Inherits CollectionBase

#Region "Declarations"

    Private Shared SqlConnection2 As System.Data.SqlClient.SqlConnection
    Private SQLCommand As New System.Data.SqlClient.SqlCommand()

    Private myFields As TableFields
    Private objTransId As IntegerField
    Private objUserName As StringField
    Private objTXNType As StringField
    Private objSchemeName As StringField
    Private objModifier As StringField
    Private objCardNumber As StringField
    Private objExpiry As StringField
    Private objValue As StringField
    Private objAuthCode As StringField
    Private objDateTime As StringField
    Private objEftSeqNum As StringField
    Private objRef As StringField
    Private objTxnResult As StringField
    Private objAuthMessage As StringField
    Private objCardHolder As StringField

#End Region

    Private Sub SetupFields()
        myFields = New TableFields("Transactions")
        myFields.Clear()
        objTransId = New IntegerField("TransactionID", 0, True)
        myFields.Add(objTransId)
        objUserName = New StringField("UserName", "", 20)
        myFields.Add(objUserName)
        objTXNType = New StringField("TxnType", "", 20)
        myFields.Add(objTXNType)
        objSchemeName = New StringField("SchemeName", "", 20)
        myFields.Add(objSchemeName)
        Me.objModifier = New StringField("Modifier", "", 20)
        myFields.Add(Me.objModifier)
        Me.objCardNumber = New StringField("CardNumber", "", 30)
        myFields.Add(Me.objCardNumber)
        Me.objExpiry = New StringField("Expiry", "", 20)
        myFields.Add(Me.objExpiry)
        Me.objValue = New StringField("TxnValue", "", 20)
        myFields.Add(Me.objValue)
        Me.objAuthCode = New StringField("AuthCode", "", 20)
        myFields.Add(Me.objAuthCode)
        Me.objDateTime = New StringField("DateTime", "", 20)
        myFields.Add(Me.objDateTime)
        Me.objEftSeqNum = New StringField("EFTSeqNum", "", 20)
        myFields.Add(Me.objEftSeqNum)
        Me.objRef = New StringField("Referance", "", 100)
        myFields.Add(Me.objRef)
        Me.objTxnResult = New StringField("TxnResult", "", 100)
        myFields.Add(Me.objTxnResult)
        Me.objAuthMessage = New StringField("AuthMessage", "", 100)
        myFields.Add(Me.objAuthMessage)
        Me.objCardHolder = New StringField("CardholderName", "", 80)
        myFields.Add(Me.objCardHolder)
    End Sub

    '''<summary>
    ''' Create new ICP table
    '''</summary>
    Public Sub New()
        MyBase.New()
        Load()
    End Sub

    '''<summary>
    ''' create new ICPTable
    '''</summary>
    ''' <param name='CmNumber'>Collection Manager Number</param>
    ''' <remarks>Find all of the rows that refer to this CM number</remarks>
    Public Sub New(ByVal CmNumber As String)
        MyBase.New()
        Load(CmNumber)
    End Sub

    Public Shared Function GetNextSeq() As Integer
        Dim ocmd As New SqlClient.SqlCommand("Select max(EFTSeqNum) as  maxeft from transactions", SqlConnection2)
        Dim objResult As Object
        If SqlConnection2.State = ConnectionState.Closed Then SqlConnection2.Open()
        objResult = ocmd.ExecuteScalar
        If Not objResult Is System.DBNull.Value Then
            Return CInt(objResult) + 1
        Else
            Return -1
        End If

        Try

        Catch ex As Exception
            Medscreen.LogError(ex, , "Get transaction id")

        Finally
            SqlConnection2.Close()

        End Try
    End Function

    Public Shared Function GetTransactionID(ByVal EFTNo As Integer) As Integer
        Dim ocmd As New SqlClient.SqlCommand("Select TransactionID from transactions where EFTSeqNum = @EFTSeqNum", SqlConnection2)
        Dim objResult As Object
        If SqlConnection2.State = ConnectionState.Closed Then SqlConnection2.Open()
        ocmd.Parameters.AddWithValue("@EFTSeqNum", EFTNo)
        objResult = ocmd.ExecuteScalar
        If Not objResult Is System.DBNull.Value Then
            Return CInt(objResult)
        Else
            Return -1
        End If

        Try

        Catch ex As Exception
            Medscreen.LogError(ex, , "Get transaction id")
        Finally
            SqlConnection2.Close()

        End Try
    End Function

    '''<summary>
    ''' Load ICPTable
    '''</summary>
    ''' <param name='Where'>Where Clause</param>
    ''' <returns>TRUE if rows found</returns>
    ''' <remarks>Find all of the rows that refer to this CM number</remarks>
    Public Function Load(Optional ByVal Where As String = "") As Boolean
        Dim oread As SqlClient.SqlDataReader
        Dim strTransId As String
        Dim objICP As ICPRow
        Dim ipos1 As Integer
        Dim ipos2 As Integer

        SetupFields()
        'Dim ICPConnectString As String = "Persist Security Info=False;Initial Catalog=ICP;Data Source=andrew\netsdk;Packet Size=4096;Integrated Security=SSPI;workstation id=TS02DOCK;"
        Dim ICPConnectString As String = MedscreenlibConfig.Connections.Item("ICP")
        ICP.SqlConnection2 = New System.Data.SqlClient.SqlConnection()
        ICP.SqlConnection2.ConnectionString = ICPConnectString
        Try
            SQLCommand.Connection = ICP.SqlConnection2

            SQLCommand.CommandText = myFields.SelectString

            'SQLCommand.CommandText += " where DateTime > '" & Now.AddDays(-40).ToString("yyyyMMdd") & "'"
            SQLCommand.CommandText += " where referance like '%" & Mid(Where, 3) & "%' or referance like '%" & Mid(Where, 3) & "'"
            ICP.SqlConnection2.Open()
            oread = Me.SQLCommand.ExecuteReader
            While oread.Read

                If oread.IsDBNull(0) Then
                Else
                    strTransId = oread.GetValue(0)
                    Debug.WriteLine("Transaction ID: " & strTransId)
                    objICP = New ICPRow()
                    objICP.TransactionId = strTransId
                    objICP.Fields.readfields(oread)
                    If Where.Length > 0 Then
                        If InStr(objICP.Reference.ToUpper, Where.ToUpper) > 0 Then 'Has Cm code embedded
                            If objICP.AuthMessage.ToUpper = "CONFIRMED" Then 'Check it has been accepted
                                Me.Add(objICP)
                            End If
                        ElseIf InStr(objICP.Reference.ToUpper, Mid(Where.ToUpper, 3)) > 0 Then
                            ipos1 = InStr(objICP.Reference.ToUpper, Mid(Where.ToUpper, 3))
                            ipos2 = InStr(objICP.Reference.ToUpper, "CM")
                            If ipos1 * ipos2 > 0 Then
                                If ipos1 - ipos2 < 10 Then
                                    If objICP.AuthMessage.ToUpper = "CONFIRMED" Then 'Check it has been accepted
                                        Me.Add(objICP)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Me.Add(objICP)
                    End If
                End If

            End While
            oread.Close()

            Return True
        Catch ex As Exception
            Medscreen.LogError(ex)
            Return False
        Finally
            ICP.SqlConnection2.Close()
        End Try


    End Function

    Public Function insertRow(ByVal icpr As ICPRow) As Boolean
        Dim SQLInsert As String = "insert into transactions (CardNumber,Expiry,txnvalue,datetime,referance,SchemeName,AccountNumber,EFTSeqNum,CardHolderName,UserName,AuthMessage) " & _
        "values(@CardNumber,@Expiry,@txnvalue,@datetime,@referance,@SchemeName,@AccountNumber,@EFTSeqNum,@CardHolderName,@UserName,@AuthMessage)"
        Dim ocmd As New SqlClient.SqlCommand(SQLInsert, ICP.SqlConnection2)
        Dim intret As Integer
        Try
            ocmd.Parameters.AddWithValue("@CardNumber", icpr.Cardnumber)
            ocmd.Parameters.AddWithValue("@Expiry", icpr.Expiry)
            ocmd.Parameters.AddWithValue("@txnvalue", icpr.TxnValue)
            ocmd.Parameters.AddWithValue("@datetime", icpr.DateTime)
            ocmd.Parameters.AddWithValue("@referance", icpr.Reference)
            ocmd.Parameters.AddWithValue("@SchemeName", icpr.SchemeName)
            ocmd.Parameters.AddWithValue("@AccountNumber", "MDS")
            ocmd.Parameters.AddWithValue("@EFTSeqNum", icpr.EFTSeqNum)
            ocmd.Parameters.AddWithValue("@CardHolderName", icpr.CardHolder)
            ocmd.Parameters.AddWithValue("@UserName", icpr.UserName)
            ocmd.Parameters.AddWithValue("@AuthMessage", icpr.AuthMessage)
            If ICP.SqlConnection2.State = ConnectionState.Closed Then ICP.SqlConnection2.Open()

            intret = ocmd.ExecuteNonQuery
        Catch ex As SqlClient.SqlException
        Catch ex As Exception
        Finally
            ICP.SqlConnection2.Close()
        End Try
    End Function

    Public Function SetInvoiceNumber(ByVal invNumber As String, ByVal icpr As ICPRow) As Boolean
        Dim ocmd As New SqlClient.SqlCommand("UpdateInvoice", ICP.SqlConnection2)
        Dim intret As Integer
        Try
            ocmd.CommandType = CommandType.StoredProcedure
            ocmd.Parameters.AddWithValue("@TransId", icpr.TransactionId)
            ocmd.Parameters.AddWithValue("@InvNum", invNumber)
            If ICP.SqlConnection2.State = ConnectionState.Closed Then ICP.SqlConnection2.Open()

            intret = ocmd.ExecuteNonQuery

        Catch ex As Exception
        Finally
            ICP.SqlConnection2.Close()
        End Try
        Return intret = 1
    End Function

    '''<summary>
    ''' convert table to CSV
    '''</summary>
    ''' <param name='Filename'>Where the file will be stored</param>
    Public Sub ToCSV(ByVal Filename As String)
        Dim w As IO.StreamWriter
        Dim objIcp As ICPRow
        Dim i As Integer

        w = New IO.StreamWriter(Filename, False)

        Try
            w.WriteLine(Me.myFields.CSVFileHeader)
            For i = 0 To Me.Count - 1
                objIcp = Me.Item(i)
                w.WriteLine(objIcp.Fields.CSVFileRow)
            Next
            w.Flush()
            w.Close()
        Catch ex As Exception
            Medscreen.LogError(ex)
        End Try


    End Sub

    '''<summary>
    ''' return a row in the table
    '''</summary>
    ''' <param name='index'>Index of row to be returned</param>
    ''' <returns>an ICP Row <see cref="ICPRow"/></returns>
    Public Property Item(ByVal index As Integer) As ICPRow
        Get
            Return CType(MyBase.List.Item(index), ICPRow)
        End Get
        Set(ByVal Value As ICPRow)

        End Set
    End Property

    '''<summary>
    ''' add a line to the table
    '''</summary>
    ''' <param name='item'>ICP transaction to add</param>
    ''' <returns>Index of Row added</returns>
    Public Function Add(ByVal item As ICPRow) As Integer
        Return MyBase.List.Add(item)
    End Function
End Class

'''<summary>
''' One Stop Interface table
'''</summary>
Public Class IntColl
    Inherits CollectionBase

#Region "Declarations"


    Private myFields As TableFields
    Private objId As IntegerField
    Private objCentre As StringField
    Private objExpectedNumber As IntegerField
    Private objCustomerID As StringField
    Private objDateToStart As DateField
    Private objStatus As StringField
    Private objCMJobNumber As StringField
    Private objPurchOrder As StringField
    Private objPrePaid As BooleanField
    Private objInvoiceNumber As StringField
    Private objOperator As StringField
    Private objCCNumber As StringField
    Private objCCexpiry As DateField
    Private objCCAuthCode As StringField
    Private objCCPayee As StringField
    Private objNettCost As DoubleField
    Private objVatCharge As DoubleField
    Private objScheduleId As StringField
    Private objMedscreenReference As StringField
    Private obyPaymentType As StringField
    Private objOptional1 As StringField
    Private objOptional2 As StringField
    Private objOptional3 As StringField
    Private objOptional4 As StringField
    Private objOptional5 As StringField
    Private objOptional6 As StringField
    Private objDateModified As TimeStampField = New TimeStampField("DATE_MODIFIED")

    Private oCmd As New OleDb.OleDbCommand()
#End Region

    Private Sub SetupFields()
        myFields = New TableFields("OneStopInterface")

        objId = New IntegerField("OISID", 0, True)
        myFields.Add(objId)
        objCentre = New StringField("CENTRE", "", 20)
        myFields.Add(objCentre)
        objExpectedNumber = New IntegerField("EXPECTED_NUMBER", 0)
        myFields.Add(objExpectedNumber)
        objCustomerID = New StringField("CUSTOMER_ID", "", 10)
        myFields.Add(objCustomerID)
        objDateToStart = New DateField("DATE_TO_START", DateField.ZeroDate)
        myFields.Add(objDateToStart)
        objStatus = New StringField("STATUS", "", 1)
        myFields.Add(objStatus)
        objCMJobNumber = New StringField("CM_JOB_NUMBER", "", 10)
        myFields.Add(objCMJobNumber)
        objPurchOrder = New StringField("Purch_order", "", 20)
        myFields.Add(objPurchOrder)
        objPrePaid = New BooleanField("PREPAID", "F")
        myFields.Add(objPrePaid)
        objInvoiceNumber = New StringField("INVOICE_NUMBER", "", 20)
        myFields.Add(objInvoiceNumber)
        objOperator = New StringField("OPERATOR", "", 15)
        myFields.Add(objOperator)
        objCCNumber = New StringField("CCNUMBER", "", 20)
        myFields.Add(objCCNumber)
        objCCexpiry = New DateField("CCEXPIRY", DateField.ZeroDate)
        myFields.Add(objCCexpiry)
        objCCAuthCode = New StringField("CCAUTHCODE", "", 10)
        myFields.Add(objCCAuthCode)
        objCCPayee = New StringField("CCPAYEE", "", 25)
        myFields.Add(objCCPayee)
        objNettCost = New DoubleField("NETTCOST", 0.0)
        myFields.Add(objNettCost)
        objVatCharge = New DoubleField("VATCHARGE", 0.0)
        myFields.Add(objVatCharge)
        objScheduleId = New StringField("SCHEDULE_ID", "", 10)
        myFields.Add(objScheduleId)
        objMedscreenReference = New StringField("MEDSCREENREFERENCE", "", 20)
        myFields.Add(objMedscreenReference)
        obyPaymentType = New StringField("PAYMENTTYPE", "", 1)
        myFields.Add(obyPaymentType)
        objOptional1 = New StringField("OPTIONAL1", "", 20)
        myFields.Add(objOptional1)
        objOptional2 = New StringField("OPTIONAL2", "", 120)
        myFields.Add(objOptional2)
        objOptional3 = New StringField("OPTIONAL3", "", 120)
        myFields.Add(objOptional3)
        objOptional4 = New StringField("OPTIONAL4", "", 120)
        myFields.Add(objOptional4)
        objOptional5 = New StringField("OPTIONAL5", "", 120)
        myFields.Add(objOptional5)
        objOptional6 = New StringField("OPTIONAL6", "", 120)
        myFields.Add(objOptional6)
        myFields.Add(Me.objDateModified)
    End Sub


    '''<summary>
    ''' create One Stop Interface table row entity
    '''</summary>
    Public Sub New()
        MyBase.New()
        SetupFields()
    End Sub

    '''<summary>
    ''' Create a new interface using a CM number
    ''' </summary>
    ''' <returns>New or existing InterfaceClass</returns>
    Public Overloads Function NewInterface() As InterfaceClass2
        Dim dataRead As OleDb.OleDbDataReader = Nothing
        Dim oCmd As New OleDb.OleDbCommand()
        Dim intOISID As Integer
        Dim intRet As Integer
        Dim ifc As InterfaceClass2 = Nothing

        oCmd.Connection = medConnection.Connection
        oCmd.CommandText = "Select max(OISID) + 1 from onestopinterface"
        Try
            CConnection.SetConnOpen()
            dataRead = oCmd.ExecuteReader
            If dataRead.Read Then
                intOISID = dataRead.GetDecimal(0)
                dataRead.Close()
                oCmd.CommandText = "insert into onestopinterface (oisid) values(?)" ' & _
                '   intOISID & ")"
                oCmd.Parameters.Clear()
                oCmd.Parameters.Add(CConnection.IntegerParameter("OISID", intOISID))
                intRet = oCmd.ExecuteNonQuery

                ifc = New InterfaceClass2(intOISID)
            End If

        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "New Interface")
            ifc = Nothing
        Finally
            If Not dataRead Is Nothing Then dataRead.Close()
            'IntConn.Close()
        End Try
        Return ifc
    End Function

    '''<summary>
    ''' Create a new interface using a CM number
    ''' </summary>
    ''' <param name='intId'>OISID</param>
    ''' <returns>New or existing InterfaceClass</returns>
    Public Overloads Function NewInterface(ByVal IntId As Integer) As InterfaceClass2
        Dim intOISID As Integer

        Dim ifc As InterfaceClass2 = Nothing
        Try
            intOISID = IntId
            ifc = New InterfaceClass2(intOISID)
        Catch ex As Exception
        Finally
        End Try
        Return ifc
    End Function

    '''<summary>
    ''' Create a new interface using a CM number
    ''' </summary>
    ''' <param name='strId'>CM_Job_number</param>
    ''' <returns>New or existing InterfaceClass</returns>
    Public Overloads Function NewInterface(ByVal strId As String) As InterfaceClass2
        Dim dataRead As OleDb.OleDbDataReader = Nothing
        Dim oCmd As New OleDb.OleDbCommand()
        Dim intOISID As Integer
        'Dim dr As XMLInterface.OneStopInterfaceRow


        Try
            oCmd.CommandText = "select oisid from onestopinterface where cm_job_number = ? " '" & strId.Trim & "'"
            oCmd.Parameters.Add(CConnection.StringParameter("CMJOBNUMBER", strId.Trim, 10))

            oCmd.Connection = medConnection.Connection
            CConnection.SetConnOpen()
            dataRead = oCmd.ExecuteReader
            If dataRead.Read Then
                intOISID = dataRead.GetValue(0)
                dataRead.Close()
                Return New InterfaceClass2(intOISID)
            Else
                dataRead.Close()
                Dim tmpInt As InterfaceClass2 = MyClass.NewInterface
                tmpInt.CMJobNumber = strId.Trim
                Return tmpInt
            End If

        Catch ex As Exception
            Return Nothing
        Finally
            If dataRead IsNot Nothing Then dataRead.Close()
            'IntConn.Close()
        End Try

    End Function


    Public Function ReturnFromSM(ByRef DR As InterfaceClass2, _
Optional ByVal chrLookfor As Char = "J") As Boolean
        Dim oCmd As New OleDb.OleDbCommand()
        Dim strQuery As String
        'Dim intRow As Integer
        Dim datR As OleDb.OleDbDataReader = Nothing
        'Dim blnRead As Boolean

        Try

            oCmd.Connection = medConnection.Connection
            CConnection.SetConnOpen()
            strQuery = myFields.SelectString & _
                " where status = ?  and oisid = ?"

            oCmd.CommandText = strQuery

            oCmd.Parameters.Add(CConnection.StringParameter("Status", chrLookfor, 1))
            oCmd.Parameters.Add(CConnection.IntegerParameter("OISID", DR.OisID))
            datR = oCmd.ExecuteReader

            If datR.Read Then
                DR.Fields.ReadFields(datR)
                Return True
            Else
                Return False
            End If
            Debug.WriteLine("Status : " & DR.Status)

        Catch oex As OleDb.OleDbException
            MedscreenLib.Medscreen.LogError(oex, True, "ReturnFromSM")
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex, True, "ReturnFromSM")
        Finally
            If datR IsNot Nothing Then datR.Close()

            'IntConn.Close()
        End Try

    End Function


End Class

'''<summary>
''' base class representing a Sample Manager Table
'''</summary>
''' <remarks>This class is used to represent a Sample Manager Table in the database,<para/>
''' it is inherited to provide specific functiionality in various business objects.
''' </remarks>
Public MustInherit Class SMTable

#Region "Declarations"


    Private strTableName As String

    Private myFields As TableFields

    'Field to capture PseudoRow RowID
#End Region

    '''<summary>
    ''' Create a new table row entity
    '''</summary>
    ''' <param name='TableName'>Interface table in use</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal TableName As String)
        strTableName = TableName
        myFields = New TableFields(strTableName)
    End Sub

    Public Property Fields() As TableFields
        Get
            Return myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    '''<summary>
    ''' Returns / sets the name of the table
    '''</summary>
    ''' <remarks></remarks>
    Protected Property TableName() As String
        Get
            Return strTableName
        End Get
        Set(ByVal Value As String)
            strTableName = Value
        End Set
    End Property

    '''<summary>
    ''' Updates this row into the database
    '''</summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>If the row has not been loaded will do an Insert instead
    ''' </remarks>
    Public Function DoUpdate() As Boolean
        If Me.Fields.Loaded Then
            Return Me.Fields.Update(MedscreenLib.MedConnection.Connection, , True)
        Else
            Return Me.Fields.Insert(MedscreenLib.medConnection.Connection)
        End If
    End Function

    '''<summary>
    ''' Has this row been loaded from the database
    '''</summary>
    ''' <returns>TRUE if loaded from database</returns>
    ''' <remarks></remarks>
    Public Property Loaded() As Boolean
        Get
            Return Me.Fields.Loaded
        End Get
        Set(ByVal Value As Boolean)
            Me.Fields.Loaded = Value
        End Set
    End Property

    '''<summary>
    ''' Get the contents of this row from the database
    '''</summary>
    ''' <param name='conn'>OLEDB Data Connector</param>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks></remarks>
    Public Overloads Function Load(ByVal conn As OleDb.OleDbConnection) As Boolean
        Dim blnRet As Boolean = False

        Dim strQuery As String = Me.Fields.FullRowSelect & " where "
        Dim objTf As TableField
        Dim blnFirst As Boolean = True
        Dim objParam As OleDb.OleDbParameter
        Dim oCmd As New OleDb.OleDbCommand()
        For Each objTf In Me.Fields                 'Go through each field looking for ones that form the identity
            If objTf.IsIdentity Then                'Is an identity field
                If Not blnFirst Then strQuery += " and " 'If not the first field add the AND operator
                blnFirst = False
                strQuery += objTf.FieldName & " = "      'Add the field
                If TypeOf objTf Is StringField Then         'Add the value info (different for strings and dates)
                    strQuery += "? "
                    objParam = New OleDb.OleDbParameter(objTf.FieldName, objTf.Value)
                    objParam.DbType = DbType.String
                    objParam.Size = objTf.FieldLength
                    oCmd.Parameters.Add(objParam)
                    'strQuery += "'" & objTf.Value & "' "
                ElseIf TypeOf objTf Is DateField Then
                    strQuery += "? "
                    objParam = New OleDb.OleDbParameter(objTf.FieldName, objTf.Value)
                    objParam.DbType = DbType.DateTime
                    objParam.Size = objTf.FieldLength
                    oCmd.Parameters.Add(objParam)
                    '                    strQuery += "to_date('" & CType(objTf.Value, Date).ToString("yyyyMMddHHmm") & "','yyyymmddHH24mi') "
                Else
                    strQuery += "? "
                    objParam = New OleDb.OleDbParameter(objTf.FieldName, objTf.Value)
                    If TypeOf objTf Is IntegerField Then
                        objParam.DbType = DbType.Int32
                    Else
                        objParam.DbType = DbType.Decimal
                    End If
                    'objParam.Size = objTf.FieldLength
                    oCmd.Parameters.Add(objParam)

                    'strQuery += objTf.Value & " "
                End If
            End If
        Next        'Query built


        oCmd.CommandText = strQuery
        Dim oRead As OleDb.OleDbDataReader = Nothing

        Try
            oCmd.Connection = conn
            If conn.State = ConnectionState.Closed Then conn.Open()

            oRead = oCmd.ExecuteReader
            If Not oRead Is Nothing Then                            'Reader open 
                If oRead.Read Then
                    Me.Fields.readfields(oRead)
                    Me.Fields.Loaded = True                         'We had an open reader with no errors so we succesfully loaded from the database
                End If
            End If


            blnRet = True
        Catch ex As Exception
            Medscreen.LogError(ex, , "Table Load - " & Me.TableName & "-" & oCmd.CommandText)
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            conn.Close()
        End Try

        Return blnRet

    End Function

    '''<summary>
    ''' Get the contents of this row from the database
    '''</summary>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>Uses medConnection.Connection to provide the connection
    ''' </remarks>
    ''' <seealso cref="medConnection.Connection"/>
    Public Overloads Function Load() As Boolean
        Return MyClass.Load(medConnection.Connection)
    End Function


End Class

'''<summary>
''' base class for the various interface classes such as sample, accounts
'''</summary>
Public MustInherit Class SMInterface
    Inherits SMTable
#Region "Declarations"


    Private objId As IntegerField = New IntegerField("ID", 0, True)
    Private objStatus As StringField = New StringField("STATUS", "", 1)
    Private objMessage As StringField = New StringField("MESSAGE", "", 400)
    Private objRequestCode As StringField = New StringField("REQUESTCODE", "", 4)
#End Region

    '''<summary>
    ''' create a new interface table entry with common fields
    '''</summary>
    Public Sub New(ByVal TableName As String)
        MyBase.New(TableName)
        SetupFields()
        GetMaxID()
    End Sub

    Private Sub SetupFields()
        Fields.Add(objId)
        Fields.Add(objStatus)
        Fields.Add(objMessage)
        Fields.Add(objRequestCode)
    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' Alternative constructor using a sequence
    ''' </summary>
    ''' <param name="TableName"></param>
    ''' <param name="Sequence"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>01-nov-2011</created><Author>Taylor</Author></revision></revisionHistory>
    Public Sub New(ByVal TableName As String, ByVal Sequence As String)
        MyBase.new(TableName)
        SetupFields()
        Me.ID = CInt(CConnection.NextSequence(Sequence))
        Me.DoUpdate()
    End Sub

    '''<summary>
    ''' delete interface row 
    '''</summary>
    ''' <returns>TRUE If succesful</returns>
    Public Function Delete() As Boolean
        Return Me.Fields.Delete(medConnection.Connection)
    End Function


    '''<summary>
    ''' Get the current highest ID number in table
    '''</summary>
    ''' <returns>Maximum Value</returns>
    ''' <remarks>This routine gets the current maximum value 
    ''' and then inserts a new row with value max +1.  The ID property is set to this value
    ''' </remarks>
    Public Function GetMaxID() As Integer
        Dim objCmd As New OleDb.OleDbCommand()
        Dim intMax As Integer
        Try
            objCmd.Connection = medConnection.Connection
            CConnection.SetConnOpen()
            objCmd.CommandText = "Select Max(id) from " & Me.TableName
            intMax = objCmd.ExecuteScalar + 1
            objCmd.CommandText = "Insert into " & Me.TableName & " (ID) values(" & intMax & ")"
            Dim intRet As Integer = objCmd.ExecuteNonQuery()
            Me.objId.Value = intMax
            Me.objId.OldValue = intMax
            Me.Fields.RowID = Me.Fields.GetRowId
            Me.Fields.Loaded = True
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            medConnection.Connection.Close()

        End Try
        Return intMax

    End Function

    '''<summary>
    ''' The ID of the row 
    '''</summary>
    Protected Property ID() As Integer
        Get
            Return Me.objId.Value
        End Get
        Set(ByVal Value As Integer)
            objId.Value = Value
            objId.OldValue = Value
        End Set
    End Property

    '''<summary>
    ''' Request code associated with this row
    '''</summary>
    Public Property RequestCode() As String
        Get
            Return Me.objRequestCode.Value
        End Get
        Set(ByVal Value As String)
            Me.objRequestCode.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Status of this Row
    '''</summary>
    Public Property Status() As String
        Get
            Return Me.objStatus.Value
        End Get
        Set(ByVal Value As String)
            Me.objStatus.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Wait on the row getting to a particular status value
    '''</summary>
    ''' <param name='ErrorMessage'>Error message or information returned by SM process</param>
    ''' <param name='Timeout'>Delay in 200 milliseconds before returning, default = 100</param>
    ''' <returns>Status of the row</returns>
    ''' <remarks>will loop whilst the status is <see cref="Constants.GCST_IFace_StatusLocked"/> Or 
    ''' <see cref="Constants.GCST_IFace_StatusRequest"/> will return if the status is <see cref="Constants.GCST_IFace_StatusFailed"/>
    ''' or <see cref="Constants.GCST_IFace_StatusCreated"/>
    ''' </remarks>
    Public Overridable Function Wait(ByRef ErrorMessage As String, _
    Optional ByVal Timeout As Integer = 100) As Char
        Dim strCmd As String = "Select Status from " & Me.TableName & " where id = " & CStr(Me.ID)
        Dim oCmd As New OleDb.OleDbCommand()
        Dim CRet As String = ""
        Try
            oCmd.Connection = medConnection.Connection
            oCmd.CommandText = strCmd
            CConnection.SetConnOpen()
            CRet = oCmd.ExecuteScalar

            While ((CRet = Constants.GCST_IFace_StatusLocked) OrElse (CRet = Constants.GCST_IFace_StatusRequest) OrElse (CRet = Constants.GCST_IFace_StatusAvailable)) And Timeout > 0
                System.Threading.Thread.Sleep(200)
                CRet = oCmd.ExecuteScalar
                Timeout -= 1
            End While
            If CRet = Constants.GCST_IFace_StatusFailed Then
                oCmd.CommandText = "Select message from " & Me.TableName & " where id = " & CStr(Me.ID)
                ErrorMessage = oCmd.ExecuteScalar
            End If
            If CRet = Constants.GCST_IFace_StatusCreated Then
                oCmd.CommandText = "Select message from " & Me.TableName & " where id = " & CStr(Me.ID)
                ErrorMessage = oCmd.ExecuteScalar
            End If

        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            medConnection.Connection.Close()

        End Try
        Return CRet.Chars(0)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Message from code 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [27/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Message() As String
        Get
            Return Me.objMessage.Value
        End Get
        Set(ByVal Value As String)
            Me.objMessage.Value = Value
        End Set
    End Property

End Class

'''<summary>
''' specialised interface class handling the Address_interface table
'''</summary>
Public Class AddressInterface
    Inherits SMInterface

    Private objAddressID As IntegerField = New IntegerField("ADDRESS_ID", -1)
    '''<summary>
    ''' Create a new interface object
    '''</summary>
    Public Sub New()
        MyBase.New("ADDRESSINTERFACE")
        Fields.Add(objAddressID)
    End Sub

    '''<summary>
    ''' Id of the row in the interface table
    '''</summary>
    Public Property AddressID() As Integer
        Get
            Return objAddressID.Value
        End Get
        Set(ByVal Value As Integer)
            objAddressID.Value = Value
        End Set
    End Property

    '''<summary>
    '''  request a new address ID DEPRECATED
    '''</summary>
    ''' <returns>New Address Id </returns>
    ''' <remarks>This function has been deprecated with the move to using sequence numbers for the address ID</remarks>
    Public Function CreateAddressID() As Integer
        Me.RequestCode = "NEW"
        Me.Status = "R"
        Dim strError As String = ""
        Me.Fields.Update(medConnection.Connection)
        Dim chRet As Char = Me.Wait(strError)
        If chRet = Constants.GCST_IFace_StatusCreated Then 'ID is valid 
            Dim oCmd As New OleDb.OleDbCommand("Select Address_id from " & Me.TableName & " where id = " & Me.ID)
            Try
                oCmd.Connection = medConnection.Connection
                CConnection.SetConnOpen()
                Me.AddressID = oCmd.ExecuteScalar

            Catch ex As Exception
            Finally
                CConnection.SetConnClosed()
            End Try
        Else
            Medscreen.LogError(strError)
        End If
        Return Me.AddressID
    End Function
End Class

'''<summary>
''' specialised interface class handling the Accounts_interface table
'''</summary>
Public Class AccountsInterface
    Inherits SMInterface
    Private Const cst_SEQ_ACCOUNTSINTERFACE As String = "SEQ_ACCOUNTSINTERFACE"
    Private Const CST_ACCOUNTSINTERFACETable As String = "ACCOUNTSINTERFACE"

    'Private MyFields As TableFields = New TableFields("ACCOUNTSINTERFACE")

#Region "Declarations"


    Private objSendDate As DateField = New DateField("SENDDATE", DateField.ZeroDate)
    Private objUserId As StringField = New StringField("USERID", "", 10)
    Private objReference As StringField = New StringField("REFERENCE", "", 20)
    Private objSendAddr As StringField = New StringField("SENDADDRESS", "", 80)
    Private objCopyType As IntegerField = New IntegerField("COPYTYPE", 0)
    Private objValue As DoubleField = New DoubleField("VALUE", 0.0)
    Private objInvoiceType As IntegerField = New IntegerField("INVOICETYPE", 0)
#End Region

#Region "Constants"
    Public Const GCST_CancInvoice_Request As String = "INCN"
    Public Const GCST_Status_Ready As String = "R"
    Public Const GCST_Status_Complete As String = "C"
    Public Const GCST_Status_Failed As String = "F"
#End Region

#Region "Code"
#Region "Procedures"


    '''<summary>
    ''' create a new instance of the class
    '''</summary>
    Public Sub New()
        MyBase.New(CST_ACCOUNTSINTERFACETable, cst_SEQ_ACCOUNTSINTERFACE)                             'Call the base class constructor, sequence form
        SetupFields()

    End Sub

    Public Sub PrintInvoice(ByVal InvoiceID As String, ByVal TransactionType As Integer)
        Me.Status = MedscreenLib.Constants.GCST_IFace_StatusRequest
        Me.Reference = InvoiceID
        Me.UserId = MedscreenLib.Glossary.Glossary.CurrentSMUser.Identity
        Dim strOutMet As String = "PRINTER"
        Me.SendAddress = MedscreenLib.Glossary.Glossary.CurrentSMUserEmail
        Me.CopyType = 2
        Me.InvoiceType = TransactionType
        Me.RequestCode = "SEND"
        Me.Update()

    End Sub

    Private Sub SetupFields()
        Me.Fields.Add(objSendDate)
        Me.Fields.Add(objUserId)
        Me.Fields.Add(objReference)
        Me.Fields.Add(objSendAddr)
        Me.Fields.Add(Me.objCopyType)
        Me.Fields.Add(Me.objValue)
        Me.Fields.Add(Me.objInvoiceType)

    End Sub
#End Region

#Region "Functions"

#End Region

#Region "Properties"


    '''<summary>
    ''' Invoice type
    '''</summary>
    ''' <remarks>Invoice type 4 is a charge Invoice type 5 is a credit</remarks>
    Public Property InvoiceType() As Integer
        Get
            Return Me.objInvoiceType.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objInvoiceType.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Update Interface row
    '''</summary>
    ''' <returns>TRUE if successful</returns>
    Public Function Update() As Boolean
        Return Fields.Update(MedConnection.Connection)

    End Function

    '''<summary>
    ''' Value of transaction
    '''</summary>
    ''' <remarks></remarks>
    Public Property Value() As Double
        Get
            Return Me.objValue.Value
        End Get
        Set(ByVal Value As Double)
            Me.objValue.Value = Value
        End Set
    End Property

    '''<summary>
    ''' ID (Sample manger User ID personnel.Identity)
    '''</summary>
    ''' <remarks></remarks>
    Public Property UserId() As String
        Get
            Return Me.objUserId.Value
        End Get
        Set(ByVal Value As String)
            Me.objUserId.Value = Value
        End Set
    End Property

    '''<summary>
    ''' How the invoice will be treated 
    '''</summary>
    ''' <remarks></remarks>
    Public Property CopyType() As Integer
        Get
            Return Me.objCopyType.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objCopyType.Value = Value
        End Set
    End Property

    '''<summary>
    ''' The invoice number or customer
    '''</summary>
    ''' <remarks></remarks>
    Public Property Reference() As String
        Get
            Return Me.objReference.Value
        End Get
        Set(ByVal Value As String)
            Me.objReference.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Email address or fax number to which the invoice should be sent
    '''</summary>
    ''' <remarks></remarks>
    Public Property SendAddress() As String
        Get
            Return Me.objSendAddr.Value
        End Get
        Set(ByVal Value As String)
            Me.objSendAddr.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Set the send date to NULL
    '''</summary>
    ''' <remarks></remarks>
    Public Sub SetSendDateNull()
        Me.objSendDate.SetNull()
    End Sub

    '''<summary>
    ''' Date invoice Sent
    '''</summary>
    ''' <remarks></remarks>
    Public Property SendDate() As Date
        Get
            Return Me.objSendDate.Value
        End Get
        Set(ByVal Value As Date)
            Me.objSendDate.Value = Value
        End Set
    End Property

#End Region
#End Region
End Class

'''<summary>
''' interface into OneStopInterface table
'''</summary>
''' <remarks>The onestopinterface table was the first of the tables produced, 
''' it attempted to be generic, but it was soon realised that each major business entity would benefit 
''' from having its own interface table.<para/>
''' This table and interface now supports collections and jobs in the main</remarks>
Public Class InterfaceClass2
#Region "Declarations"


    Private myFields As TableFields
    Private objId As IntegerField
    Private objCentre As StringField
    Private objExpectedNumber As IntegerField
    Private objCustomerID As StringField
    Private objDateToStart As DateField
    Private objStatus As StringField
    Private objCMJobNumber As StringField
    Private objPurchOrder As StringField
    Private objPrePaid As BooleanField
    Private objInvoiceNumber As StringField
    Private objOperator As StringField
    Private objCCNumber As StringField
    Private objCCexpiry As DateField
    Private objCCAuthCode As StringField
    Private objCCPayee As StringField
    Private objNettCost As DoubleField
    Private objVatCharge As DoubleField
    Private objScheduleId As StringField
    Private objMedscreenReference As StringField
    Private obyPaymentType As StringField
    Private objOptional1 As StringField
    Private objOptional2 As StringField
    Private objOptional3 As StringField
    Private objOptional4 As StringField
    Private objOptional5 As StringField
    Private objOptional6 As StringField
    Private objDateModified As TimeStampField = New TimeStampField("DATE_MODIFIED")

    'Private oConn As New OleDb.OleDbConnection()
    Private oCmd As New OleDb.OleDbCommand()
#End Region

#Region "Code"

    Private Sub SetupFields()
        myFields = New TableFields("OneStopInterface")

        objId = New IntegerField("OISID", 0, True)
        myFields.Add(objId)
        objCentre = New StringField("CENTRE", "", 20)
        myFields.Add(objCentre)
        objExpectedNumber = New IntegerField("EXPECTED_NUMBER", 0)
        myFields.Add(objExpectedNumber)
        objCustomerID = New StringField("CUSTOMER_ID", "", 10)
        myFields.Add(objCustomerID)
        objDateToStart = New DateField("DATE_TO_START", DateField.ZeroDate)
        myFields.Add(objDateToStart)
        objStatus = New StringField("STATUS", "", 1)
        myFields.Add(objStatus)
        objCMJobNumber = New StringField("CM_JOB_NUMBER", "", 10)
        myFields.Add(objCMJobNumber)
        objPurchOrder = New StringField("Purch_order", "", 20)
        myFields.Add(objPurchOrder)
        objPrePaid = New BooleanField("PREPAID", "F")
        myFields.Add(objPrePaid)
        objInvoiceNumber = New StringField("INVOICE_NUMBER", "", 20)
        myFields.Add(objInvoiceNumber)
        objOperator = New StringField("OPERATOR", "", 15)
        myFields.Add(objOperator)
        objCCNumber = New StringField("CCNUMBER", "", 20)
        myFields.Add(objCCNumber)
        objCCexpiry = New DateField("CCEXPIRY", DateField.ZeroDate)
        myFields.Add(objCCexpiry)
        objCCAuthCode = New StringField("CCAUTHCODE", "", 10)
        myFields.Add(objCCAuthCode)
        objCCPayee = New StringField("CCPAYEE", "", 25)
        myFields.Add(objCCPayee)
        objNettCost = New DoubleField("NETTCOST", 0.0)
        objNettCost.OldValue = 0.0001
        myFields.Add(objNettCost)
        objVatCharge = New DoubleField("VATCHARGE", 0.0)
        myFields.Add(objVatCharge)
        objScheduleId = New StringField("SCHEDULE_ID", "", 10)
        myFields.Add(objScheduleId)
        objMedscreenReference = New StringField("MEDSCREENREFERENCE", "", 20)
        myFields.Add(objMedscreenReference)
        obyPaymentType = New StringField("PAYMENTTYPE", "", 1)
        myFields.Add(obyPaymentType)
        objOptional1 = New StringField("OPTIONAL1", "", 20)
        myFields.Add(objOptional1)
        objOptional2 = New StringField("OPTIONAL2", "", 120)
        myFields.Add(objOptional2)
        objOptional3 = New StringField("OPTIONAL3", "", 120)
        myFields.Add(objOptional3)
        objOptional4 = New StringField("OPTIONAL4", "", 120)
        myFields.Add(objOptional4)
        objOptional5 = New StringField("OPTIONAL5", "", 120)
        myFields.Add(objOptional5)
        objOptional6 = New StringField("OPTIONAL6", "", 120)
        myFields.Add(objOptional6)
        myFields.Add(Me.objDateModified)

    End Sub

    '''<summary>
    ''' Update this row
    ''' </summary>
    ''' <remarks></remarks>
    Public Function Update() As Boolean
        Dim iRet As Integer = -1

        Try
            'oConn.ConnectionString = Support.ConnectString(6)
            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.UpdateString
            CConnection.SetConnOpen()
            iRet = oCmd.ExecuteNonQuery()
        Catch ex As Exception
            Medscreen.LogError(ex)
            Return False
        Finally
            'oConn.Close()

        End Try
        Return (iRet > 0)
    End Function

    '''<summary>
    ''' Refresh this row
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Refresh()
        'Dim iRet As Integer
        Dim oRead As OleDb.OleDbDataReader

        Try
            'oConn.ConnectionString = Support.ConnectString(6)
            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.FullRowSelect & " where oisid = ? "
            oCmd.Parameters.Clear()
            oCmd.Parameters.Add(CConnection.IntegerParameter("OISID", Me.OisID))
            CConnection.SetConnOpen()

            oRead = oCmd.ExecuteReader
            If oRead.Read Then
                myFields.readfields(oRead)
            End If
            oRead.Close()
        Catch ex As Exception
            Medscreen.LogError(ex)
        Finally
            'oConn.Close()

        End Try

    End Sub

    '''<summary>
    ''' Load this row
    ''' </summary>
    ''' <remarks></remarks>
    Public Function load() As Boolean
        Dim oRead As OleDb.OleDbDataReader
        Dim blnret As Boolean = True
        Try
            'oConn.ConnectionString = Support.ConnectString(6)
            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.FullRowSelect & " where oisid = ?"
            oCmd.Parameters.Clear()
            oCmd.Parameters.Add(CConnection.IntegerParameter("OISID", Me.OisID))
            CConnection.SetConnOpen()
            oRead = oCmd.ExecuteReader
            If oRead.Read Then
                myFields.readfields(oRead)
            End If
            oRead.Close()

        Catch ex As Exception
            Medscreen.LogError(ex)
            blnret = False
        Finally
            'oConn.Close()
        End Try
        Return blnret
    End Function

    '''<summary>
    ''' Create a new row 
    ''' </summary>
    ''' <param name='ID'>OISID</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ID As Integer)

        SetupFields()
        OisID = ID
        load()
    End Sub

    '''<summary>
    ''' Port or site used
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Centre() As String
        Get
            Return Me.objCentre.Value
        End Get
        Set(ByVal Value As String)
            objCentre.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Test schedule in use
    ''' </summary>
    ''' <remarks></remarks>
    Public Property TestSchedule() As String
        Get
            Return Me.objScheduleId.Value
        End Get
        Set(ByVal Value As String)
            objScheduleId.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Fields in database
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Fields() As TableFields
        Get
            Return Me.myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    '''<summary>
    ''' Status of request
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Status() As Char
        Get
            Return Me.objStatus.Value
        End Get
        Set(ByVal Value As Char)
            objStatus.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter usually task parameter
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional1() As String
        Get
            Return Me.objOptional1.Value
        End Get
        Set(ByVal Value As String)
            objOptional1.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter usually primary return from task
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional2() As String
        Get
            Return objOptional2.Value
        End Get
        Set(ByVal Value As String)
            objOptional2.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional3() As String
        Get
            Return objOptional3.Value
        End Get
        Set(ByVal Value As String)
            objOptional3.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter usually date SM completed action
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional4() As String
        Get
            Return objOptional4.Value
        End Get
        Set(ByVal Value As String)
            objOptional4.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional5() As String
        Get
            Return objOptional5.Value
        End Get
        Set(ByVal Value As String)
            objOptional5.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Optional parameter
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Optional6() As String
        Get
            Return objOptional6.Value
        End Get
        Set(ByVal Value As String)
            objOptional6.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Customer ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Property Customer() As String
        Get
            Return Me.objCustomerID.Value
        End Get
        Set(ByVal Value As String)
            objCustomerID.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Customer's purchase order no
    ''' </summary>
    ''' <remarks></remarks>
    Public Property PurchaseOrder() As String
        Get
            Return Me.objPurchOrder.Value
        End Get
        Set(ByVal Value As String)
            objPurchOrder.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Invoice number given to transaction
    ''' </summary>
    ''' <remarks></remarks>
    Public Property InvoiceNumber() As String
        Get
            Return Me.objInvoiceNumber.Value
        End Get
        Set(ByVal Value As String)
            objInvoiceNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Refrence given by accounts to transaction
    ''' </summary>
    ''' <remarks></remarks>
    Public Property MedscreenReference() As String
        Get
            Return Me.objMedscreenReference.Value
        End Get
        Set(ByVal Value As String)
            objMedscreenReference.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Collection Date
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CollDate() As Date
        Get
            Return Me.objDateToStart.Value
        End Get
        Set(ByVal Value As Date)
            objDateToStart.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Expected no of donors
    ''' </summary>
    ''' <remarks></remarks>
    Public Property NoDonors() As Integer
        Get
            Return Me.objExpectedNumber.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objExpectedNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Person carrying out request
    ''' </summary>
    ''' <remarks></remarks>
    Public Property [operator]() As String
        Get
            Return Me.objOperator.Value
        End Get
        Set(ByVal Value As String)
            objOperator.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Collection Manager ID
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CMJobNumber() As String
        Get
            Return Me.objCMJobNumber.Value
        End Get
        Set(ByVal Value As String)
            objCMJobNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' how the transaction will be paid
    ''' </summary>
    ''' <remarks></remarks>
    Public Property PaymentType() As String
        Get
            Return Me.obyPaymentType.Value
        End Get
        Set(ByVal Value As String)
            obyPaymentType.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Is this collection pre paid or not 
    ''' </summary>
    ''' <remarks></remarks>
    Public Property PrePaid() As String
        Get
            Return Me.objPrePaid.Value
        End Get
        Set(ByVal Value As String)
            objPrePaid.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Nett cost of transaction (Doesn't include VAT)
    ''' </summary>
    ''' <remarks></remarks>
    Public Property NettCost() As Decimal
        Get
            Return Me.objNettCost.Value
        End Get
        Set(ByVal Value As Decimal)
            objNettCost.Value = Value
        End Set
    End Property

    '''<summary>
    ''' VAT element
    ''' </summary>
    ''' <remarks></remarks>
    Public Property VatCost() As Decimal
        Get
            Return Me.objVatCharge.Value
        End Get
        Set(ByVal Value As Decimal)
            objVatCharge.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Authorisation code from Card supplier
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CreditAuthorisation() As String
        Get
            Return Me.objCCAuthCode.Value
        End Get
        Set(ByVal Value As String)
            Me.objCCAuthCode.Value = Value
        End Set
    End Property

    '''<summary>
    ''' who is paying for the collection
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CreditPayee() As String
        Get
            Return Me.objCCPayee.Value
        End Get
        Set(ByVal Value As String)
            Me.objCCPayee.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Date of card expiry
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CreditExpiry() As Date
        Get
            Return Me.objCCexpiry.Value
        End Get
        Set(ByVal Value As Date)
            objCCexpiry.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Credit card number Packed
    ''' </summary>
    ''' <remarks></remarks>
    Public Property CreditCardNumber() As String
        Get
            Return Me.objCCNumber.Value
        End Get
        Set(ByVal Value As String)
            objCCNumber.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Primary key into table
    ''' </summary>
    ''' <remarks></remarks>
    Public Property OisID() As Integer
        Get
            Return Me.objId.Value
        End Get
        Set(ByVal Value As Integer)
            Me.objId.Value = Value
        End Set
    End Property
#End Region
End Class

#Region "Workflow"



'''<summary>
''' An individual element of a workflow display
''' </summary>
''' <remarks></remarks>
''' 
Public Class WorkFlowItem

#Region "Enumerations"



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' State of workflow element    ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action>Workflow item type for collection received added</Action></revision>
    ''' <revision><Author>[taylor]</Author><date> </date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum Progress
        '''<summary>Waiting on another element</summary>
        waiting
        '''<summary>Current active element</summary>
        Active
        '''<summary>Done but not necessarily complete</summary>
        Done
        '''<summary>Completed</summary>
        Complete
    End Enum

    '''<summary>Possible shapes to draw</summary>
    Public Enum Shape
        '''<summary>Rectangle</summary>
        Rectangle
        '''<summary>Circle</summary>
        Circle
    End Enum

    '''<summary></summary>
    Public Enum ItemTypes
        '''<summary>Not defined</summary>
        NotDefined
        '''<summary>Booking</summary>
        Booking
        '''<summary>Customer has confirmed</summary>
        CustomerConfirm
        '''<summary>Sent to collecting officer</summary>
        COSend
        '''<summary>Collecting officer has confirmed</summary>
        COConfirm
        '''<summary>Collected status D</summary>
        Collected
        '''<summary>Collected status R</summary>
        Received
        '''<summary>Collection reported</summary>
        Reported
        '''<summary>Invoiced</summary>
        Invoiced
    End Enum

#End Region

#Region "Declarations"


    Private myRect As Rectangle
    Private myBrush As Brush
    Private myText As String
    Private myFont As Font
    Private myProgress As Progress
    Private subText As String() = {"", "", ""}
    Private myRight As Single
    Private myLineText As String
    Private myShape As Shape = Shape.Rectangle
    Private MyLink As Integer = -1
    Private myParent As WorkFlowCollection
    Private isSelected As Boolean
    Private myItemType As ItemTypes = ItemTypes.NotDefined
    Private myStatusLine As String = ""
    Private myStatusShow As String = ""
    Private myHasCalendar As Boolean = False

#End Region

    Public Sub New()
        MyBase.new()
    End Sub

    '''<summary>
    ''' Create a new workflow item
    ''' </summary>
    ''' <param name='aRect'>Rectangle to use</param>
    ''' <param name='RightText'></param>
    Public Sub New(ByVal aRect As Rectangle, ByVal RightText As Single)
        myRect = aRect
        myBrush = New SolidBrush(Color.Beige)
        myFont = New Font(System.Drawing.FontFamily.GenericSansSerif, 10, FontStyle.Bold)
        myProgress = Progress.waiting
        myRight = RightText
        myLineText = ""
    End Sub

    'Indicates whether there is a calendar associated with this workflow item 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Some workflow items are time based and have a calendar associated with them
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property HasCalendar() As Boolean
        Get
            Return myHasCalendar
        End Get
        Set(ByVal Value As Boolean)
            myHasCalendar = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of workflow item  
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ItemType() As ItemTypes
        Get
            Return myItemType
        End Get
        Set(ByVal Value As ItemTypes)
            myItemType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current status of the item    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property StatusShow() As String
        Get
            Return myStatusShow
        End Get
        Set(ByVal Value As String)
            myStatusShow = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Status Line for workflow item    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property StatusLine() As String
        Get
            If myStatusLine.Length = 0 Then
                Return Me.Text & Me.Line1 & " " & Me.Line2 & " " & Me.LineText
            Else
                Return myStatusLine
            End If

        End Get
        Set(ByVal Value As String)
            myStatusLine = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is this item the currently selected one    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Selected() As Boolean
        Get
            Return isSelected
        End Get
        Set(ByVal Value As Boolean)
            isSelected = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Descriptive text    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LineText() As String
        Get
            Return myLineText
        End Get
        Set(ByVal Value As String)
            myLineText = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Workflow this item belongs to    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Parent() As WorkFlowCollection
        Get
            Return myParent
        End Get
        Set(ByVal Value As WorkFlowCollection)
            myParent = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw workflow    ''' 
    ''' </summary>
    ''' <param name="e"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Draw(ByVal e As System.Drawing.Graphics)

        Dim tBrush As Brush

        If isSelected Then
            tBrush = SystemBrushes.HighlightText
            myBrush = SystemBrushes.Highlight
        Else
            myBrush = New SolidBrush(Color.Beige)
            If myProgress = Progress.Done Then
                tBrush = SystemBrushes.ControlText
            ElseIf myProgress = Progress.Active Then
                tBrush = Brushes.DarkGreen
                myBrush = New SolidBrush(Color.BlanchedAlmond)
            ElseIf myProgress = Progress.Complete Then
                tBrush = Brushes.PaleVioletRed
                myBrush = New SolidBrush(Color.LightPink)
            Else
                tBrush = Brushes.Gray
                myBrush = New SolidBrush(Color.PaleGoldenrod)
            End If
        End If
        If myShape = Shape.Rectangle Then
            e.FillRectangle(myBrush, myRect)
            e.DrawRectangle(Drawing.Pens.Black, myRect)
        Else
            e.FillPie(myBrush, myRect, 0, 360)
            e.DrawArc(Drawing.Pens.Black, myRect, 0, 360)

        End If
        Dim textSize As SizeF = e.MeasureString(myText, myFont)
        e.DrawString(myText, myFont, tBrush, myRect.X + 10, _
            myRect.Y + 5)

        Dim subFont As New Font(System.Drawing.FontFamily.GenericSansSerif, 8, FontStyle.Regular)
        Dim i As Integer
        Dim sText As String

        For i = 0 To 2
            sText = subText.GetValue(i)
            If sText.Length > 0 Then
                e.DrawString(sText, subFont, tBrush, _
                    New RectangleF(myRect.X + 15, _
                    myRect.Y + textSize.Height + 5 + (i * 10), myRect.Width - 15, 15))
            End If
        Next
        DrawLine(e)
    End Sub


    Private Sub DrawLine(ByVal e As Drawing.Graphics)

        Dim wf As WorkFlowItem

        If Me.Link = -1 Then Exit Sub

        wf = myParent.Item(Me.Link)
        If wf Is Nothing Then Exit Sub

        'Dim X As Integer = wf.Rect.Width + wf.Rect.X
        Dim X As Integer = wf.Rect.X                'Change sides 
        Dim Y As Integer = wf.Rect.Y + (wf.Rect.Height / 2)
        Dim X2 As Integer
        Dim Y2 As Integer
        Dim ArrowArray As PointF() = {New PointF(0, 0), New PointF(0, 0), New PointF(0, 0)}

        If wf.Rect.Top <> Me.Rect.Top Then
            X = wf.Rect.Width + wf.Rect.X
            If wf.LineText.Trim.Length > 0 Then
                e.DrawString(wf.LineText, myFont, SystemBrushes.ControlText, X + 5, Y)
            End If
            X = wf.Rect.X
            e.FillPie(SystemBrushes.ControlDarkDark, X - 10, Y - 5, 10, 10, 0, 360)
            X -= 10

            e.DrawLine(SystemPens.ControlDarkDark, X, Y, X - 10, Y)
            X -= 10
            X2 = X
            Y2 = wf.Rect.Bottom + (Me.Rect.Top - wf.Rect.Bottom) / 2
            e.DrawLine(SystemPens.ControlDarkDark, X, Y, X2, Y2) 'draw line down 
            Y = Y2
            X2 = Me.Rect.Left + (Me.Rect.Width / 2)
            e.DrawLine(SystemPens.ControlDarkDark, X, Y, X2, Y2)
            X = X2
            Y2 = Me.Rect.Top
            e.DrawLine(SystemPens.ControlDarkDark, X, Y, X2, Y2)
            ArrowArray.SetValue(New PointF(X - 6, Y2 - 6), 0)
            ArrowArray.SetValue(New PointF(X + 6, Y2 - 6), 1)
            ArrowArray.SetValue(New PointF(X, Y2), 2)

        Else
            If Math.Abs(Me.Rect.Right - wf.Rect.Left) < Math.Abs(wf.Rect.Right - Me.Rect.Left) Then
                X2 = Me.Rect.Right
                X = wf.Rect.Left
            Else
                X = wf.Rect.Right
                X2 = Me.Rect.Left

            End If
            Y -= 5
            If Math.Abs(Me.Rect.Right - wf.Rect.Left) < Math.Abs(wf.Rect.Right - Me.Rect.Left) Then
                X -= 10
                e.FillPie(SystemBrushes.ControlDarkDark, X, Y - 5, 10, 10, 0, 360)
                e.DrawLine(SystemPens.ControlDarkDark, X, Y, X2, Y)
                ArrowArray.SetValue(New PointF(X2 + 6, Y - 6), 0)
                ArrowArray.SetValue(New PointF(X2 + 6, Y + 6), 1)
                ArrowArray.SetValue(New PointF(X2, Y), 2)
            Else
                e.FillPie(SystemBrushes.ControlDarkDark, X, Y - 5, 10, 10, 0, 360)
                X += 10
                e.DrawLine(SystemPens.ControlDarkDark, X, Y, X2, Y)
                ArrowArray.SetValue(New PointF(X2 - 6, Y - 6), 0)
                ArrowArray.SetValue(New PointF(X2 - 6, Y + 6), 1)
                ArrowArray.SetValue(New PointF(X2, Y), 2)
            End If

        End If

        e.FillPolygon(SystemBrushes.ControlDarkDark, ArrowArray)

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Has this item been hit by mouse click    ''' 
    ''' </summary>
    ''' <param name="x">X position of click</param>
    ''' <param name="y">Y position of click</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IsHit(ByVal x As Single, ByVal y As Single) As Boolean
        Dim blnReturn As Boolean = False

        blnReturn = (x >= myRect.Left And x <= myRect.Right And _
            y >= myRect.Top And y <= myRect.Bottom)
        Return blnReturn
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Text associated with Item    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Text() As String
        Get
            Return myText
        End Get
        Set(ByVal Value As String)
            myText = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rectangle associated with Item    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Rect() As Rectangle
        Get
            Return myRect
        End Get
        Set(ByVal Value As Rectangle)
            myRect = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sub line 1 of descriptive text    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Line1() As String
        Set(ByVal Value As String)
            subText.SetValue(Value, 0)
        End Set
        Get
            Return subText.GetValue(0)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sub line 2 of descriptive text    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Line2() As String
        Set(ByVal Value As String)
            subText.SetValue(Value, 1)
        End Set
        Get
            Return subText.GetValue(1)
        End Get
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Sub line 3 of descriptive text    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public WriteOnly Property Line3() As String
        Set(ByVal Value As String)
            subText.SetValue(Value, 2)
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Shape of object    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ObjectShape() As Shape
        Get
            Return myShape
        End Get
        Set(ByVal Value As Shape)
            myShape = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Link to item number    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Link() As Integer
        Get
            Return MyLink
        End Get
        Set(ByVal Value As Integer)
            MyLink = Value
        End Set
    End Property

    '''  -----------------------------------------------------------------------------
    ''' <summary>
    ''' Current status of item    ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Status() As Progress
        Get
            Return myProgress
        End Get
        Set(ByVal Value As Progress)
            myProgress = Value
        End Set
    End Property
End Class

'''<summary>
''' An entire workflow 
''' </summary>
''' <remarks></remarks>
Public Class WorkFlowCollection
    Inherits CollectionBase

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Add a workflow item    ''' 
    ''' </summary>
    ''' <param name="item">Item to add</param>
    ''' <returns>Position of Added Item</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Add(ByVal item As WorkFlowItem) As Integer
        item.Parent = Me
        Return MyBase.List.Add(item)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Retrieve Item from list    ''' 
    ''' </summary>
    ''' <param name="index">Position to get</param>
    ''' <returns>Item at position</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function Item(ByVal index As Integer) As WorkFlowItem
        Return CType(MyBase.List.Item(index), WorkFlowItem)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Deselect all items    ''' 
    ''' </summary>
    ''' <returns>Previous selected item</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function DeSelectAll() As WorkFlowItem
        Dim wf As WorkFlowItem = Nothing
        Dim i As Integer

        For i = 0 To Me.Count - 1
            wf = Item(i)
            If wf.Selected Then
                wf.Selected = False
                Exit For
            End If
            wf = Nothing
        Next

        Return wf
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' See if the mouse click has hit any of the elements of this work flow    ''' 
    ''' </summary>
    ''' <param name="x">X Mouse position</param>
    ''' <param name="y">Y Mouse position</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function HitTest(ByVal x As Single, ByVal y As Single) As WorkFlowItem
        Dim wf As WorkFlowItem = Nothing
        Dim i As Integer

        For i = 0 To Me.Count - 1
            wf = Item(i)
            If wf.IsHit(x, y) Then
                Exit For
            End If
            wf = Nothing
        Next
        Return wf
    End Function

End Class
#End Region

'''<summary>
''' Simplistic printing class probably better to use either RTF and Word<para/>
''' Reporter or XML/XSL and IE
''' </summary>
Public Class Printing
    Private printFont As Font
    Private streamToPrint As StreamReader
    Private filePath As String
    Private pd As New PrintDocument()

    '''<summary>Printing Font</summary>
    Public Property Font() As Font
        Get
            Return printFont
        End Get
        Set(ByVal Value As Font)
            printFont = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' 
    ''' </summary>
    ''' <param name="Document"></param>
    ''' <param name="FontFace"></param>
    ''' <param name="FontSize"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Document As String, _
              Optional ByVal FontFace As String = "Arial", Optional ByVal FontSize As Integer = 10)
        printFont = New Font(FontFace, FontSize)
        Me.filePath = Document
    End Sub

    ' The PrintPage event is raised for each page to be printed.
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="ev"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Private Sub pd_PrintPage(ByVal sender As Object, ByVal ev As PrintPageEventArgs)
        Dim linesPerPage As Single = 0
        Dim yPos As Single = 0
        Dim count As Integer = 0
        Dim leftMargin As Single = ev.MarginBounds.Left
        Dim topMargin As Single = ev.MarginBounds.Top
        Dim tab1 As Single = ev.PageBounds.Width / 12 * 5
        Dim line As String = Nothing

        ' Calculate the number of lines per page.
        linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics)

        ' Iterate over the file, printing each line.
        Dim iTab As Integer
        While count < linesPerPage
            line = streamToPrint.ReadLine()

            If line Is Nothing Then
                Exit While
            End If

            iTab = InStr(line, vbTab)           'Check for tabs

            yPos = topMargin + count * printFont.GetHeight(ev.Graphics)
            If iTab = 0 Then
                ev.Graphics.DrawString(line, printFont, Brushes.Black, leftMargin, _
                    yPos, New StringFormat())
            Else
                Dim strTemp As String
                strTemp = Mid(line, 1, iTab - 1)
                line = Mid(line, iTab + 1)
                ev.Graphics.DrawString(strTemp, printFont, Brushes.Black, leftMargin, _
                    yPos, New StringFormat())
                ev.Graphics.DrawString(line, printFont, Brushes.Black, tab1, _
                    yPos, New StringFormat())

            End If
            count += 1
        End While

        ' If more lines exist, print another page.
        If Not (line Is Nothing) Then
            ev.HasMorePages = True
        Else
            ev.HasMorePages = False
        End If
    End Sub

    ' Print the file.
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Print()
        Try
            streamToPrint = New StreamReader(filePath)
            Try

                AddHandler pd.PrintPage, AddressOf pd_PrintPage
                ' Print the document.
                pd.Print()
            Finally
                streamToPrint.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub 'Printing    

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' 
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Settings() As PrinterSettings
        Get
            Return pd.PrinterSettings
        End Get
        Set(ByVal Value As PrinterSettings)
            pd.PrinterSettings = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''     ''' 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub PrintDialog()
        Try
            streamToPrint = New StreamReader(filePath)
            Try

                Dim pd As New PrintDocument()
                AddHandler pd.PrintPage, AddressOf pd_PrintPage
                ' Print the document.
                pd.Print()
            Finally
                streamToPrint.Close()
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub 'Printing    



End Class

'''<summary>
''' A collection of ReportSchedule items
''' </summary>
Public Class ReportSchedules
    Inherits CollectionBase
    'Class to manage report shedules 
#Region "Declarations"


    Private myFields As New TableFields("REPORT_SCHEDULE")

    Private objScheduleID As StringField = New StringField("SCHEDULE_ID", "", 10, True)
    Private objReportID As StringField = New StringField("REPORT_ID", "", 10)
    Private objCustomerID As StringField = New StringField("CUSTOMER_ID", "", 10)
    Private objNextReport As DateField = New DateField("NEXT_REPORT", DateField.ZeroDate)
    Private objRepeat As BooleanField = New BooleanField("REPEAT", "F")
    Private objRepeatUnits As StringField = New StringField("REPEAT_UNITS", "", 20)
    Private objRepeatInterval As IntegerField = New IntegerField("REPEAT_INTERVAL", 0)
    Private objUntil As DateField = New DateField("UNTIL", DateField.ZeroDate)
    Private objSendMethod As StringField = New StringField("SEND_METHOD", "", 10)
    Private objRecipients As StringField = New StringField("RECIPIENTS", "", 400)
    Private objModifiedON As DateField = New DateField("MODIFIED_ON", DateField.ZeroDate)
    Private objModifiedBY As StringField = New StringField("MODIFIED_BY", "", 10)
    Private objRemoveflag As BooleanField = New BooleanField("REMOVEFLAG", "F")
    Private objBCC As StringField = New StringField("BCC", "", 400)
    Private objSendLog As BooleanField = New BooleanField("SEND_LOG", "T")
    Private objInstance As StringField = New StringField("INSTANCE", "", 100)
    Private objUserName As StringField = New StringField("USERNAME", "", 40)
    Private objPassword As StringField = New StringField("PASSWORD", "", 40)

    'New fields add 15-mar-07
    'Type of schedule used only to indicate whether it is Customer or background at the moment
    Private objScheduleType As StringField = New StringField("SCHEDULE_TYPE", "", 10)
    'Period of time and frequency of a background report
    Private objRunPeriod As StringField = New StringField("RUN_PERIOD", "", 40)

    Private myClientID As String = ""
#End Region

    Private Sub SetupFields()
        myFields.Add(objScheduleID)
        myFields.Add(objReportID)
        myFields.Add(objCustomerID)
        myFields.Add(objNextReport)
        myFields.Add(objRepeat)
        myFields.Add(objRepeatUnits)
        myFields.Add(objRepeatInterval)
        myFields.Add(objUntil)
        myFields.Add(objSendMethod)
        myFields.Add(objRecipients)
        myFields.Add(objModifiedON)
        myFields.Add(objModifiedBY)
        myFields.Add(objRemoveflag)
        myFields.Add(Me.objScheduleType)
        myFields.Add(Me.objRunPeriod)
        myFields.Add(objBCC)
        myFields.Add(objSendLog)
        myFields.Add(objInstance)
        myFields.Add(objUserName)
        myFields.Add(objPassword)


    End Sub

    '''<summary>Create a new Report Schedule collection instance</summary>
    Public Sub New()
        MyBase.new()
        SetupFields()
    End Sub


    Public Function ToHTML() As String
        Dim strXML As String = ToXMLElement.ToString
        Dim strHTML As String = Medscreen.ResolveStyleSheet(strXML, "ReportSchedules.xsl", 0)
        strHTML = Medscreen.ReplaceString(strHTML, Chr(13), "<br/>")
        Return strHTML

    End Function

    Public Function ToXMLElement() As XElement
        Dim Element As New XElement("ReportSchedules")
        For Each repSched As ReportSchedule In MyBase.InnerList
            Element.Add(repSched.ToXMLElement)
        Next
        Return Element
    End Function

    Public Function ToXMLElementFull() As XElement
        Dim Element As XElement = ToXMLElement()
        Return Element
    End Function


    '''<summary>Add a ReportSchedule to the collection</summary>
    ''' <param name='Item'>ReportSchedule to add</param>
    ''' <returns>Position of Item in list</returns>
    Public Function Add(ByVal Item As ReportSchedule) As Integer
        Return MyBase.List.Add(Item)
    End Function

    '''<summary>
    ''' Customer Profile ID, if provided then the <see cref="Load" /> method will get a collection of schedules for the client 
    ''' </summary>
    Public Property ClientID() As String
        Get
            Return Me.myClientID
        End Get
        Set(ByVal Value As String)
            myClientID = Value
        End Set
    End Property

    '''<summary>
    ''' Retrieve an <see cref="ReportSchedule" /> from list
    ''' </summary>
    ''' <param name='index'>Index by position to item</param>
    Default Public Overloads Property Item(ByVal index As Integer) As ReportSchedule
        Get
            Return CType(MyBase.List.Item(index), ReportSchedule)
        End Get
        Set(ByVal Value As ReportSchedule)
            MyBase.List.Item(index) = Value
        End Set
    End Property

    '''<summary>
    ''' Load the collection from database
    ''' </summary>
    Public Function Load() As Boolean
        Dim blnRet As Boolean = False

        Dim oRead As OleDb.OleDbDataReader = Nothing
        Try
            Dim oCmd As New OleDb.OleDbCommand()
            oCmd.Connection = MedConnection.Connection
            oCmd.CommandText = myFields.SelectString
            If Me.myClientID.Trim.Length > 0 Then
                oCmd.CommandText += " where customer_id = '" & Me.myClientID.Trim & "' and removeflag = 'F'"
            Else
                oCmd.CommandText += " where removeflag = 'F'"
            End If
            CConnection.SetConnOpen()

            Dim oRepSch As ReportSchedule

            oRead = oCmd.ExecuteReader
            While oRead.Read
                oRepSch = New ReportSchedule()
                oRepSch.Fields.readfields(oRead)
                Me.Add(oRepSch)
            End While
            blnRet = True
        Catch ex As Exception
            MedscreenLib.Medscreen.LogError(ex)
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            CConnection.SetConnClosed()
        End Try
        Return blnRet
    End Function

    Public Shared Function CreateSchedule() As ReportSchedule
        Dim SchedId As String = CConnection.NextSequence("SEQ_REPORT_SCHEDULE")
        Dim NewSched As New ReportSchedule
        NewSched.ScheduleID = SchedId
        NewSched.Fields.Insert(CConnection.DbConnection)
        Return NewSched
    End Function

End Class



''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : ReportSchedule
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class to manage a schedule for an individual report''' 
''' </summary>
''' <remarks>
''' Report schedules are used in conjunction with Report definitions <see cref="CRMenuItem"/> and their parameters 
''' <see cref="CRFormulaItem"/>.  A report schedule relates a report to a customer, basically it ties the customer 
''' and the recipients at that customer to a reporting event.  <para/>There is no reason that these 
''' data structures need to be tied to particular times, for example if all the time information is left blank (NULL), 
''' then any automated code would ignore these entries other code could however could look for unscheduled reports 
''' for a customer and use that information.<para/>
''' The report's schedule (if scheduled) is controlled by the <see cref="ReportSchedule.NextReport"/> field, on completion of 
''' sending a report this should be changed by calling the <see cref="ReportSchedule.NextReportDate"/> method, 
''' which uses the <see cref="ReportSchedule.RepeatUnits"/> to identify when the next report will be run in conjunction with the
''' <see cref="ReportSchedule.RepeatInterval"/>, which acts as a multiplier for the <see cref="ReportSchedule.RepeatUnits"/>.  Whether a report is repeated or not 
''' is controlled by the <see cref="ReportSchedule.Repeat"/> and the <see cref="ReportSchedule.UntilDate" /> parameters.  <para/>
''' How this report is sent is controlled by the <see cref="ReportSchedule.SendMethod"/>. An enumeration exists that has the possible values that this
''' can take <see cref="MedscreenLib.Constants.SendMethod"/>.  
'''  
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class ReportSchedule

#Region "Constants"
    Public Const GCST_BACKGROUND_REPORT As String = "BACKGROUND"
    Public Const GCST_CUSTOMER_REPORT As String = "CUSTOMER"
    Public Const GCST_SOAP As String = "SOAP"
#End Region

#Region "Declarations"


    Private myFields As New TableFields("REPORT_SCHEDULE")

    Private objScheduleID As StringField = New StringField("SCHEDULE_ID", "", 10, True)
    Private objReportID As StringField = New StringField("REPORT_ID", "", 10)
    Private objCustomerID As StringField = New StringField("CUSTOMER_ID", "", 10)
    Private objNextReport As DateField = New DateField("NEXT_REPORT", DateField.ZeroDate)
    Private objRepeat As BooleanField = New BooleanField("REPEAT", "F")
    Private objRepeatUnits As StringField = New StringField("REPEAT_UNITS", "", 20)
    Private objRepeatInterval As IntegerField = New IntegerField("REPEAT_INTERVAL", 0)
    Private objUntil As DateField = New DateField("UNTIL", DateField.ZeroDate)
    Private objSendMethod As StringField = New StringField("SEND_METHOD", "", 10)
    Private objRecipients As StringField = New StringField("RECIPIENTS", "", 400)
    Private objModifiedON As DateField = New DateField("MODIFIED_ON", DateField.ZeroDate)
    Private objModifiedBY As StringField = New StringField("MODIFIED_BY", "", 10)
    Private objRemoveflag As BooleanField = New BooleanField("REMOVEFLAG", "F")
    Private objBCC As StringField = New StringField("BCC", "", 400)
    Private objSendLog As BooleanField = New BooleanField("SEND_LOG", "T")
    Private objInstance As StringField = New StringField("INSTANCE", "", 100)
    Private objUserName As StringField = New StringField("USERNAME", "", 40)
    Private objPassword As StringField = New StringField("PASSWORD", "", 40)

    'New fields add 15-mar-07
    'Type of schedule used only to indicate whether it is Customer or background at the moment
    Private objScheduleType As StringField = New StringField("SCHEDULE_TYPE", "", 10)
    'Period of time and frequency of a background report
    Private objRunPeriod As StringField = New StringField("RUN_PERIOD", "", 40)

    Private myCrystalReport As CRMenuItem
#End Region

    Private Sub SetupFields()
        myFields.Add(objScheduleID)
        myFields.Add(objReportID)
        myFields.Add(objCustomerID)
        myFields.Add(objNextReport)
        myFields.Add(objRepeat)
        myFields.Add(objRepeatUnits)
        myFields.Add(objRepeatInterval)
        myFields.Add(objUntil)
        myFields.Add(objSendMethod)
        myFields.Add(objRecipients)
        myFields.Add(objModifiedON)
        myFields.Add(objModifiedBY)
        myFields.Add(objRemoveflag)
        myFields.Add(Me.objScheduleType)
        myFields.Add(Me.objRunPeriod)
        myFields.Add(objBCC)
        myFields.Add(objSendLog)
        myFields.Add(objInstance)
        myFields.Add(objUserName)
        myFields.Add(objPassword)

    End Sub

    '''<summary>Create a new ReportSchedule instance</summary>
    Public Sub New()
        SetupFields()
    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' Check report is in date
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Function IsInDate() As Boolean
        Return UntilDate = DateField.ZeroDate OrElse UntilDate > Now
    End Function

    Public Sub New(ByVal ReportID As String)
        MyClass.New()
        Dim ocmd As New OleDb.OleDbCommand("Select * from report_schedule where schedule_id = ?", MedConnection.Connection)
        ocmd.Parameters.Add(CConnection.StringParameter("ID", ReportID, 10))
        Dim oRead As OleDb.OleDbDataReader = Nothing
        Try
            CConnection.SetConnOpen()
            oRead = ocmd.ExecuteReader
            While oRead.Read
                myFields.readfields(oRead)
            End While
        Catch ex As Exception
            Medscreen.LogError(ex, , "reading reportschedule")
        Finally
            If Not oRead Is Nothing Then
                If Not oRead.IsClosed Then oRead.Close()
            End If
            CConnection.SetConnClosed()
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Period over which background process runs.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property RunPeriod() As String
        Get
            Return Me.objRunPeriod.Value
        End Get
        Set(ByVal Value As String)
            Me.objRunPeriod.Value = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Schedule Type
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ScheduleType() As String
        Get
            Return Me.objScheduleType.Value
        End Get
        Set(ByVal Value As String)
            objScheduleType.Value = Value
        End Set
    End Property
    ''' <summary>
    ''' Blind CC 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BCC() As String
        Get
            Return Me.objBCC.Value
        End Get
        Set(ByVal value As String)
            Me.objBCC.Value = value
        End Set
    End Property
    '''<summary>
    ''' A description of what the report is doing used for display purposes
    ''' </summary>
    ''' <param name='AsXML'>Produce in XML format default = TRUE</param>
    ''' <returns>A description of the schedule</returns>
    Public Function ActionString(Optional ByVal AsXML As Boolean = True) As String
        Dim strAction As String = ""
        Dim oCRF As CRFormulaItem
        Dim j As Integer

        Dim ampString As String = ""

        If AsXML Then
            ampString = "&amp;"
        Else
            ampString = "&"
        End If

        For j = 0 To myCrystalReport.Formulae.Count - 1
            oCRF = myCrystalReport.Formulae.Item(j)
            If oCRF.FieldName.Trim.Length > 0 Then

                strAction += ampString & oCRF.Formula & "="
                Select Case oCRF.FieldName
                    Case "NEXT_REPORT"
                        strAction += DateSerial(Me.NextReport.Year, Me.NextReport.Month, 1).AddDays(-1).ToString("dd-MMM-yyyy")
                    Case "RECIPIENTS"
                        strAction += Me.Recipients
                    Case "CUSTOMER_ID"
                        strAction += Me.CustomerID
                    Case "SMID"
                        strAction += oCRF.Value
                    Case "VALUE"
                        strAction += oCRF.Value
                    Case "PREV_REPORT"
                        strAction += DateSerial(Me.PrevReport.Year, Me.PrevReport.Month, Me.PrevReport.Day).ToString("dd-MMM-yyyy")

                End Select

            End If
        Next

        Return strAction

    End Function

    '''<summary>
    '''     ''' The <see cref="CRMenuItem" /> to be reported
    ''' </summary>
    Public Property CrystalReport() As CRMenuItem
        Get
            If Me.myCrystalReport Is Nothing Then
                Me.myCrystalReport = MedscreenLib.Glossary.Glossary.Menus.Item(Me.ReportID)
            End If
            Return myCrystalReport
        End Get
        Set(ByVal Value As CRMenuItem)
            myCrystalReport = Value
        End Set
    End Property

    '''<summary>Customer Profile ID report is for</summary>
    Public Property CustomerID() As String
        Get
            Return Me.objCustomerID.Value
        End Get
        Set(ByVal Value As String)
            objCustomerID.Value = Value
        End Set
    End Property

    '''<summary></summary>
    Public Property Fields() As TableFields
        Get
            Return myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    '''<summary>Stored value of next report</summary>
    Public Property NextReport() As Date
        Get
            If objNextReport.Value < DateField.ZeroDate Then
                objNextReport.Value = DateField.ZeroDate
            End If
            Return Me.objNextReport.Value
        End Get
        Set(ByVal Value As Date)
            objNextReport.Value = Value
        End Set
    End Property

    Private Function DoNormalAdd(ByVal tmpDate As Date, ByRef shouldReturn As Boolean, ByVal HourArray As String(), ByVal dayarray As String()) As Date
        shouldReturn = False
        Dim strU As String = Me.RepeatUnits.Trim.ToUpper

        If strU.Length = 1 Then
            If RepeatInterval > 0 Then      ' = 0 then it is a list of hours
                If strU = "H" Or strU = "I" Then
                    tmpDate = Date.Parse(tmpDate.ToString("dd-MMM-yyyy") & " " & tmpDate.ToString("HH:mm"))
                End If
                If strU = "M" Then tmpDate = tmpDate.AddMonths(Me.RepeatInterval)
                If strU = "Q" Then tmpDate = tmpDate.AddMonths(Me.RepeatInterval * 3)
                If strU = "D" Then tmpDate = tmpDate.AddDays(Me.RepeatInterval)
                If strU = "Y" Then tmpDate = tmpDate.AddYears(Me.RepeatInterval)
                If strU = "W" Then tmpDate = tmpDate.AddDays((Me.RepeatInterval * 7))
                If strU = "H" Then tmpDate = tmpDate.AddHours(Me.RepeatInterval)
                If strU = "I" Then tmpDate = tmpDate.AddMinutes(Me.RepeatInterval)
            Else            'we have a list of hours check it is valid first
                If HourArray.Length = 0 Then
                    shouldReturn = True : Exit Function
                End If
                Dim i As Integer = 0
                Dim hours As Integer = NextHour(HourArray(i))
                If hours = -1 Then          'Error
                    shouldReturn = True : Exit Function
                End If
                Dim advanced As Boolean = False
                While i < HourArray.Length And tmpDate.Hour >= hours
                    i += 1                  'Increement to next value
                    If i < HourArray.Length Then hours = NextHour(HourArray(i))

                End While
                Dim tmpdate2 As Date = tmpDate
                tmpDate = tmpDate.AddHours(tmpDate.Hour * -1).AddHours(hours)
                If tmpDate = tmpdate2 Then
                    WeAreAfterLastTime(tmpDate, Dayarray, HourArray)
                    'tmpDate = tmpDate.AddDays(1)
                    'tmpDate = Date.Parse(tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(0))
                End If
            End If
        ElseIf strU.Trim.Length = 0 Then
            shouldReturn = True : Exit Function
        Else
            If strU.Chars(0) = "X" Then 'Complex info
                If strU.Chars(1) = "W" Then 'Weekday 
                    Dim chWeekday As System.DayOfWeek = Val(strU.Chars(3))

                    If strU.Chars(2) = " " Or strU.Chars(2) = "0" Then 'On the following weekday
                        tmpDate = tmpDate.AddDays(1)
                        While tmpDate.DayOfWeek <> chWeekday
                            tmpDate = tmpDate.AddDays(1)
                        End While
                    Else
                        Dim intMonth As Integer = tmpDate.Month
                        While tmpDate.Month = intMonth
                            tmpDate = tmpDate.AddDays(1)        ' move to next month 
                        End While
                        Dim intRepeat As Integer = Val(strU.Chars(2))
                        While intRepeat > 0         ' Should be at the start of the month 
                            While tmpDate.DayOfWeek <> chWeekday    'No need to get to the day of week
                                tmpDate = tmpDate.AddDays(1)
                            End While
                            intRepeat -= 1                          '
                            If intRepeat > 0 Then                   'Are we at the nth if no next day and continue
                                tmpDate = tmpDate.AddDays(1)
                            End If
                        End While

                    End If
                End If
            ElseIf strU.Length > 5 Then
                tmpDate = CrystalSupport.DateParameter(strU, "", tmpDate.ToString("dd-MMM-yyyy"))

            End If
        End If
        Return tmpDate
    End Function
    Private Shared Sub WeAreAfterLastTime(ByRef tmpDate As Date, ByVal DayArray As String(), ByVal HourArray As String())
        tmpDate = tmpDate.AddDays(1) 'Move on one day
        If DayArray(0).ToUpper = "WEEKDAY" Then
            If tmpDate.DayOfWeek = DayOfWeek.Saturday Then

                tmpDate = tmpDate.AddDays(2)                    'Move to Monday
            ElseIf tmpDate.DayOfWeek = DayOfWeek.Sunday Then
                tmpDate = tmpDate.AddDays(1)                    'Move to Monday probably will never occur
            End If
        ElseIf DayArray(0).ToUpper = "WORKINGDAY" Then
            While Medscreen.IsAHoliday(tmpDate) OrElse tmpDate.DayOfWeek = DayOfWeek.Saturday Or tmpDate.DayOfWeek = DayOfWeek.Sunday
                tmpDate = tmpDate.AddDays(1)                    'Move over holidays
            End While
        Else

            If Weekday(tmpDate) > DayArray(1) Then
                While Weekday(tmpDate) <> DayArray(0)
                    tmpDate = tmpDate.AddDays(1)
                End While
            End If
        End If

        Dim tmpString As String = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(0)
        tmpDate = Date.Parse(tmpString)
    End Sub

    Private Function NextHour(ByVal HourString As String) As Integer
        Dim intRet As Integer = -1
        If HourString.Length > 0 Then
            Dim tmpDate As Date = Date.Parse(Today.ToString("dd-MMM-yyyy") & " " & HourString)
            intRet = tmpDate.Hour

        End If
        Return intRet
    End Function

    Private Function DealWithBackgroundReports(ByVal tmpDate As Date, ByRef DayArray As String(), ByRef HourArray As String(), ByRef tmpDate2 As Date, ByRef shouldReturn As Boolean) As Date
        shouldReturn = False
        Dim tmpString As String
        Dim DataArray As String() = Me.RunPeriod.Split(New Char() {"="}) 'Split into days and hours
        If DataArray.Length < 2 Then shouldReturn = True : Exit Function
        Dim strDays As String = DataArray(0)            'Days of the week part
        'investigating whether weekday and workingday could be supported.
        Dim strHours As String = DataArray(1)           'Hours Part
        DayArray = strDays.Split(New Char() {"-", ","})
        HourArray = strHours.Split(New Char() {"-", ","})

        'Temporary string to hold date for parsing

        'Create a last time date
        If HourArray.Length = 1 Then                        'Run only at one time
            tmpString = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(0)
        ElseIf HourArray.Length = 2 Then                    'run between times
            tmpString = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(HourArray.Length - 1)
        Else                                                'run at multiple times
            Dim iH As Integer = 0
            While iH < HourArray.Length - 1 AndAlso Date.Parse(HourArray(iH)).Hour < tmpDate.Hour
                iH += 1
            End While
            tmpString = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(iH)

        End If
        tmpDate2 = Date.Parse(tmpString)
        Return DateTime.MinValue
    End Function
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Calculate the next report date
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' The code acts differently for customer reports and background reports
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [16/03/2007]</date><Action></Action></revision>
    ''' <revision><modified>03-Apr-2012 15:33</modified><Author>CONCATENO\taylor</Author><Action>Added a call to CrystalSupport.DateParameter on special dates</Action></revision>
    ''' <revision><modified>26-Apr-2012 07:10</modified><Author>CONCATENO\taylor</Author><Action>Added missing hour determinant</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function NextReportDate() As Date
        Dim tmpDate As Date = Me.NextReport
        Dim Originaldate As Date = tmpDate
        'If Me.ScheduleType = GCST_CUSTOMER_REPORT OrElse RunPeriod.Trim.Length = 0 Then
        If RunPeriod.Trim.Length = 0 Then
            Dim strU As String = Me.RepeatUnits.Trim.ToUpper
            If strU.Length = 1 Then
                If strU = "M" Then tmpDate = tmpDate.AddMonths(Me.RepeatInterval)
                If strU = "Q" Then tmpDate = tmpDate.AddMonths(Me.RepeatInterval * 3)
                If strU = "D" Then tmpDate = tmpDate.AddDays(Me.RepeatInterval)
                If strU = "Y" Then tmpDate = tmpDate.AddYears(Me.RepeatInterval)
                If strU = "W" Then tmpDate = tmpDate.AddDays((Me.RepeatInterval * 7))
                If strU = "I" Then tmpDate = tmpDate.AddMinutes(Me.RepeatInterval)
                If strU = "H" Then tmpDate = tmpDate.AddHours(Me.RepeatInterval)
            ElseIf strU.Trim.Length = 0 Then
                Exit Function
            Else
                If strU.Chars(0) = "X" Then 'Complex info
                    If strU.Chars(1) = "W" Then 'Weekday 
                        Dim chWeekday As System.DayOfWeek = Val(strU.Chars(3))

                        If strU.Chars(2) = " " Or strU.Chars(2) = "0" Then 'On the following weekday
                            tmpDate = tmpDate.AddDays(1)
                            While tmpDate.DayOfWeek <> chWeekday
                                tmpDate = tmpDate.AddDays(1)
                            End While
                        Else
                            Dim intMonth As Integer = tmpDate.Month
                            While tmpDate.Month = intMonth
                                tmpDate = tmpDate.AddDays(1)        ' move to next month 
                            End While
                            Dim intRepeat As Integer = Val(strU.Chars(2))
                            While intRepeat > 0         ' Should be at the start of the month 
                                While tmpDate.DayOfWeek <> chWeekday    'No need to get to the day of week
                                    tmpDate = tmpDate.AddDays(1)
                                End While
                                intRepeat -= 1                          '
                                If intRepeat > 0 Then                   'Are we at the nth if no next day and continue
                                    tmpDate = tmpDate.AddDays(1)
                                End If
                            End While

                        End If
                    End If
                ElseIf strU.Length > 5 Then
                    tmpDate = CrystalSupport.DateParameter(strU, "", tmpDate.ToString("dd-MMM-yyyy"))

                End If
            End If
        Else 'Deal with Background Reports 
            Dim lShouldReturn1 As Boolean
            Dim DayArray As String() = Nothing
            Dim HourArray As String() = Nothing
            Dim tmpDate2 As Date
            Dim lResult As Date = DealWithBackgroundReports(tmpDate, DayArray, HourArray, tmpDate2, lShouldReturn1)
            If lShouldReturn1 Then
                Exit Function
            End If
            Dim tmpString As String
            If DayArray.Length = 1 AndAlso DayArray(0).Length = 1 AndAlso Weekday(tmpDate) <> DayArray(0) Then  'Only run on a single day of week
                While Weekday(tmpDate) <> DayArray(0)
                    tmpDate = tmpDate.AddDays(1)
                End While
                tmpString = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(0)
                tmpDate = Date.Parse(tmpString)
            ElseIf DayArray.Length = 2 AndAlso DayArray(0).Length = 1 AndAlso (Weekday(tmpDate) < DayArray(0) OrElse Weekday(tmpDate) > DayArray(1)) Then  'outside bounds
                'We need to move to the start date and time.
                While Weekday(tmpDate) <> DayArray(0)
                    tmpDate = tmpDate.AddDays(1)
                End While
                tmpString = tmpDate.ToString("dd-MMM-yyyy") & " " & HourArray(0)
                tmpDate = Date.Parse(tmpString)
            ElseIf tmpDate > tmpDate2 Then      'We are after the last time 
                WeAreAfterLastTime(tmpDate, DayArray, HourArray)
            Else 'Do a normal add 
                Dim lShouldReturn As Boolean
                tmpDate = DoNormalAdd(tmpDate, lShouldReturn, HourArray, DayArray)
                If lShouldReturn Then
                    Exit Function
                End If
            End If
        End If
        If tmpDate.Year = DateField.ZeroDate.Year Then
            tmpDate = Originaldate.AddDays(1)
            Medscreen.LogAction("Date Reset")
        End If
        Return tmpDate '.AddDays(1)

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Export report into filename according to method in send method
    ''' </summary>
    ''' <param name="cr"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [17/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    <CLSCompliant(False)> _
    Public Function ExportReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument) As String
        Dim cstSendMethod As MedscreenLib.Constants.SendMethod
        If Me.SendMethod = "HTML" Then
            cstSendMethod = Constants.SendMethod.HTML
        ElseIf Me.SendMethod = "PDF" Then
            cstSendMethod = Constants.SendMethod.PDF
        ElseIf Me.SendMethod = "EXCEL" Then
            cstSendMethod = Constants.SendMethod.Excel
        Else
            cstSendMethod = Constants.SendMethod.PDF
        End If
        Return CrystalSupport.ExportReport(cr, cstSendMethod)
        'Dim tmpFileName As String
        'If Me.SendMethod = "EMAIL" Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "DOC")
        '    Medscreen.ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.WordForWindows, tmpFileName)
        'ElseIf Me.SendMethod = "PDF" Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "PDF")
        '    Medscreen.ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.PortableDocFormat, tmpFileName)
        'ElseIf Me.SendMethod = "EXCEL" Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "XLS")
        '    Medscreen.ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.ExcelRecord, tmpFileName)
        'ElseIf Me.SendMethod = "RTF" Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "RTF")
        '    Medscreen.ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.RichText, tmpFileName)
        'ElseIf Me.SendMethod = "HTML" Then
        '    tmpFileName = Medscreen.GetFileName("Report-" & Now.ToString("HHmmss") & "-", Now, "HTM")
        '    Medscreen.ExportToDisk(cr, CrystalDecisions.[Shared].ExportFormatType.HTML40, tmpFileName)
        'End If
        'Return tmpFileName


    End Function
    <CLSCompliant(False)> _
    Public Function PrintReport(ByVal cr As CrystalDecisions.CrystalReports.Engine.ReportDocument, ByVal PrinterName As String) As Boolean

        cr.PrintOptions.PrinterName = PrinterName
        cr.PrintToPrinter(1, True, 0, 0)
    End Function

    '''<summary>Date of Previous report, note: does not work for complex reports</summary>
    Public Function PrevReport() As Date
        Dim tmpDate As Date = DateSerial(Me.NextReport.Year, Me.NextReport.Month, 1)

        Dim strU As String = Me.RepeatUnits.Trim.ToUpper
        If strU.Length = 1 Then
            If strU = "M" Then tmpDate = tmpDate.AddMonths(-Me.RepeatInterval)
            If strU = "D" Then tmpDate = tmpDate.AddDays(-Me.RepeatInterval)
            If strU = "Q" Then tmpDate = tmpDate.AddMonths(-Me.RepeatInterval * 3)
            If strU = "Y" Then tmpDate = tmpDate.AddYears(-Me.RepeatInterval)
            If strU = "W" Then tmpDate = tmpDate.AddDays(-(Me.RepeatInterval * 7))
        ElseIf strU.Length > 1 AndAlso strU.Chars(0) = "X" Then
            tmpDate = tmpDate.AddMonths(-RepeatInterval)
        ElseIf Me.RepeatInterval = 1 Then
            tmpDate = tmpDate.AddYears(-1)
        End If


        Return tmpDate

    End Function

    Public Property RecipientsRaw() As String
        Get
            Return objRecipients.Value
        End Get
        Set(ByVal value As String)
            objRecipients.Value = value
        End Set
    End Property

    '''<summary>A mailing list</summary>
    Public ReadOnly Property Recipients() As String
        Get
            Dim StrReturn As String = objRecipients.Value
            Dim iPos As Integer = InStr(StrReturn, "{")
            Dim iPos2 As Integer = InStr(StrReturn, "}")
            If iPos * iPos2 > 0 Then
                Dim strContact As String = Mid(StrReturn, iPos + 1, (iPos2 - iPos) - 1)
                If ScheduleType = GCST_CUSTOMER_REPORT Then
                    'Get contact email address
                    Dim oColl As New Collection
                    oColl.Add(CustomerID)
                    oColl.Add(strContact)
                    Dim Email As String = MedscreenLib.CConnection.PackageStringList("Lib_contacts.CustomerContactsEmails", oColl)
                    StrReturn = Medscreen.ReplaceString(StrReturn, "{" & strContact & "}", Email)
                End If

            End If
            Return StrReturn
        End Get
        'Set(ByVal Value As String)
        '    If InStr(Value, "{") > 0 Then
        '        Exit Property
        '    End If
        '    objRecipients.Value = Value
        'End Set
    End Property

    Private ReadOnly Property TidyRecipients() As String

        Get
            Dim strRet As String = Recipients
            strRet = Medscreen.ReplaceString(strRet, ",", "," & vbCr)
            strRet = Medscreen.ReplaceString(strRet, ";", "," & vbCr)

            Return strRet
        End Get
    End Property

    '''<summary>Repeat or one off report</summary>
    Public Property Repeat() As Boolean
        Get
            Return Me.objRepeat.Value
        End Get
        Set(ByVal Value As Boolean)
            objRepeat.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Multiplier for <see cref="RepeatUnits" />, e.g. 2, RepeatUnits = "Q" = HalyYearly
    ''' </summary>
    Public Property RepeatInterval() As Integer
        Get
            Return Me.objRepeatInterval.Value
        End Get
        Set(ByVal Value As Integer)
            objRepeatInterval.Value = Value
        End Set
    End Property

    '''<summary>Converts Repeat Units into a Human readable form</summary>
    Public Function RepeatString() As String
        Dim strTemp As String = ""
        Dim strU As String = Me.RepeatUnits.Trim.ToUpper
        If strU.Length = 1 Then
            If strU = "M" Then strTemp = "Calendar Month"
            If strU = "D" Then strTemp = "Day"
            If strU = "Y" Then strTemp = "Year"
            If strU = "W" Then strTemp = "Week"
        ElseIf strU.Length > 0 AndAlso strU.Chars(0) = "X" Then
            If strU.Chars(1) = "W" Then 'Weekday
                If strU.Chars(2) = " " Then
                    strTemp = "Next " & [Enum].GetName(GetType(System.DayOfWeek), Val(strU.Chars(2)))
                Else
                    If strU.Chars(2) = "1" Then
                        strTemp = "First " & [Enum].GetName(GetType(System.DayOfWeek), Val(strU.Chars(2)))
                    ElseIf strU.Chars(2) = "2" Then
                        strTemp = "Second " & [Enum].GetName(GetType(System.DayOfWeek), Val(strU.Chars(2)))
                    ElseIf strU.Chars(2) = "3" Then
                        strTemp = "Third " & [Enum].GetName(GetType(System.DayOfWeek), Val(strU.Chars(2)))
                    ElseIf strU.Chars(2) = "4" Then
                        strTemp = "Fourth " & [Enum].GetName(GetType(System.DayOfWeek), Val(strU.Chars(2)))

                    End If
                End If
            End If
        End If
        Return strTemp
    End Function

    '''<summary>
    ''' The units for the repeat see remarks for details of what can be done
    ''' </summary>
    ''' <remarks>
    ''' The repeat units can be a simple type or complex <para/>
    ''' Simple Types<para/>
    ''' D - Days<para/>
    ''' W - Weeks<para/>
    ''' M - Calendar Months<para/>
    ''' Q - Quarters <para/>
    ''' Y - Years <para/><para/>
    ''' Complex Types<para/>
    ''' Complex types have the Character 'X' in the first position of the string 
    ''' At the moment the only second position defined value is 'W' indicating weekday<para/>
    ''' The fourth position indicates the numeric index into weekday enumeration e.g. 0 = Sunday<para/>
    ''' The third position indicates which week with ' ' indicating the next week day of that kind occuring.<para/>
    ''' otherwise 2 = second week etc<para/>
    ''' For Example <para/>
    ''' XW20 = 2nd Sunday 
    ''' XW 1 = Next Monday
    ''' </remarks>
    Public Property RepeatUnits() As String
        Get
            Return Me.objRepeatUnits.Value
        End Get
        Set(ByVal Value As String)
            objRepeatUnits.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Primary key of the report that is to be sent <see cref="CRMenuItem"/>
    ''' </summary>
    Public Property ReportID() As String
        Get
            Return Me.objReportID.Value

        End Get
        Set(ByVal Value As String)
            Me.objReportID.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Primary key of the schedule not usually presented to users
    ''' </summary>
    Public Property ScheduleID() As Integer
        Get
            Return Me.objScheduleID.Value
        End Get
        Set(ByVal Value As Integer)
            objScheduleID.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Method for sending report 
    ''' </summary>
    Public Property SendMethod() As String
        Get
            Return Me.objSendMethod.Value
        End Get
        Set(ByVal Value As String)
            Me.objSendMethod.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Send Method for report <see cref="MedscreenLib.Constants.SendMethod"/>
    '''</summary>
    ''' <remarks></remarks>
    ''' <seealso cref="MedscreenLib.Constants.SendMethod"/>
    Public Overloads ReadOnly Property SendTypeConstant() As MedscreenLib.Constants.SendMethod
        Get
            Dim RetPtype As MedscreenLib.Constants.SendMethod = Constants.SendMethod.NoConfirm
            If Me.SendMethod.Trim.Length > 0 Then
                Dim strPType As String = SendMethod.ToUpper
                Dim iPType As String

                For Each iPType In System.Enum.GetNames(GetType(MedscreenLib.Constants.SendMethod))
                    If iPType.ToUpper = strPType Then
                        RetPtype = [Enum].Parse(GetType(MedscreenLib.Constants.SendMethod), strPType)
                        Exit For
                    End If

                Next

            End If
            Return RetPtype

        End Get
    End Property


    Public Function ToXMLElement() As XElement
        Dim Element As New XElement("ReportSchedule")
        Element.Add(New XElement("client", CustomerID))
        Element.Add(New XElement("nextreport", NextReport.ToString("dd-MMM-yyyy HH:mm")))
        Element.Add(New XElement("recipients", TidyRecipients))
        Element.Add(New XElement("repeat", Repeat))
        Element.Add(New XElement("repeatinterval", RepeatInterval))
        If RepeatString() Is Nothing OrElse RepeatString.Trim.Length = 0 Then
            Element.Add(New XElement("repeatstring", "Hour"))
        Else
            Element.Add(New XElement("repeatstring", RepeatString))
        End If
        Element.Add(New XElement("sendmethod", SendMethod))
        If Me.UntilDate = DateField.ZeroDate Then
            Element.Add(New XElement("until", "forever"))
        Else
            Element.Add(New XElement("until", UntilDate.ToString("dd-MMM-yyyy")))
        End If

        If Not CrystalReport Is Nothing Then
            With myCrystalReport
                Element.Add(New XElement("MenuText", .MenuText))
                Element.Add(New XElement("Path", .MenuPath))
                Element.Add(New XElement("MenuType", .MenuType))
                If myCrystalReport.Formulae.Count = 0 Then
                    myCrystalReport.Formulae.Load(MedConnection.Connection)
                End If
                Element.Add(.Formulae.ToXMLElement)

                'build action string for html 
                Element.Add(New XElement("ActionString", ActionString))
                Element.Add(New XElement("Instance", .Instance))
                If .MenuType.Trim.Length > 0 Then Element.Add(New XElement("ReportType", .ReportType))


            End With
        End If

        Return Element
    End Function

    Public Function ToXMLElementFull() As XElement
        Dim Element As XElement = ToXMLElement()
        Return Element
    End Function


    '''<summary>
    ''' Convert Report schedule to XML
    ''' </summary>
    Public Function ToXML() As String
        Dim strRet As String = ToXMLElement.ToString ' "<ReportSchedule>"

        'strRet += "<client>" & Me.CustomerID & "</client>"
        'strRet += "<nextreport>" & Me.NextReport.ToString("dd-MMM-yyyy HH:mm") & "</nextreport>"
        'strRet += "<recipients>" & Me.Recipients & "</recipients>"
        'strRet += "<repeat>" & Me.Repeat & "</repeat>"
        'strRet += "<repeatinterval>" & Me.RepeatInterval & "</repeatinterval>"
        'strRet += "<repeatstring>" & Me.RepeatString & "</repeatstring>"
        'strRet += "<sendmethod>" & Me.SendMethod & "</sendmethod>"
        'If Me.UntilDate = DateField.ZeroDate Then
        '    strRet += "<until>forever</until>"
        'Else
        '    strRet += "<until>" & Me.UntilDate.ToString("dd-MMM-yyyy") & "</until>"
        'End If

        'If Not myCrystalReport Is Nothing Then
        '    strRet += "<MenuText>" & myCrystalReport.MenuText & "</MenuText>"
        '    strRet += "<Path>" & Medscreen.FixAmpersands(myCrystalReport.MenuPath) & "</Path>"
        '    strRet += "<MenuType>" & myCrystalReport.MenuType & "</MenuType>"
        '    If myCrystalReport.Formulae.Count = 0 Then
        '        myCrystalReport.Formulae.Load(MedConnection.Connection)
        '    End If
        '    strRet += "<Parameters>" & myCrystalReport.Formulae.ToXML & "</Parameters>"

        '    'build action string for html 
        '    strRet += "<ActionString>" & ActionString() & "</ActionString>"
        'End If

        'strRet += "</ReportSchedule>"
        Return strRet


    End Function

    '''<summary>
    ''' Update information in database
    ''' </summary>
    Public Function Update() As Boolean
        Return Me.myFields.Update(MedConnection.Connection)
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether report has been removed or not.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/04/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property RemoveFlag() As Boolean
        Get
            Return Me.objRemoveflag.Value
        End Get
        Set(ByVal Value As Boolean)
            Me.objRemoveflag.Value = Value
        End Set
    End Property

    '''<summary>
    ''' Date beyond which the report should not be run
    ''' </summary>
    Public Property UntilDate() As Date
        Get
            Return Me.objUntil.Value
        End Get
        Set(ByVal Value As Date)
            objUntil.Value = Value
        End Set
    End Property

    Public Sub SetUntilDateNull()
        Me.objUntil.SetNull()
    End Sub

End Class

'''<summary>
''' A collection of rows of type <see cref="SMTable"/>
''' </summary>
Public MustInherit Class SMTableCollection
    Inherits CollectionBase

    Private strTableName As String

    Private myFields As TableFields

    Public Sub New(ByVal TableName As String)
        MyBase.New()
        strTableName = TableName
        myFields = New TableFields(strTableName)
    End Sub

    Protected Friend Property Fields() As TableFields
        Get
            Return myFields
        End Get
        Set(ByVal Value As TableFields)
            myFields = Value
        End Set
    End Property

    Protected Property TableName() As String
        Get
            Return strTableName
        End Get
        Set(ByVal Value As String)
            strTableName = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Index of function 
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [08/02/2006]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function IndexOfObject(ByVal Value As Object) As Integer
        Return MyBase.List.IndexOf(Value)
    End Function


End Class

'''<summary>
''' A collection of rows of type <see cref="TimeStampedCollection"/>
''' </summary>
Public MustInherit Class TimedSmTableCollection
    Inherits SMTableCollection

    Private myTimeStamp As Date

    Public Sub New(ByVal TableName As String)
        MyBase.New(TableName)
        myTimeStamp = Now
    End Sub

    Protected Property TimeStamp() As Date
        Get
            Return myTimeStamp
        End Get
        Set(ByVal Value As Date)
            myTimeStamp = Value
        End Set
    End Property

    Public MustOverride Function RefreshChanged() As Boolean
End Class

'''<summary>
''' A collection of rows of type <see cref="TimeStampedCollection"/>
''' </summary>
Public MustInherit Class TimeStampedCollection
    Inherits CollectionBase

    Private myTimeStamp As Date


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new collection
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()
        myTimeStamp = Now
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Timestamp of last database access
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property TimeStamp() As Date
        Get
            Return myTimeStamp
        End Get
        Set(ByVal Value As Date)
            myTimeStamp = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Refresh any changed items
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [02/10/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public MustOverride Function RefreshChanged() As Boolean

    'End Function
End Class
#Region "Public_Variables"

'''<summary>
''' Public variables
''' </summary>
Public Class MedscreenVariables

    Private Shared strPrinterCCPlain As String = "LASER3"
    Private Shared strCollmanRoot As String = MedscreenLib.Constants.GCST_X_DRIVE & "\CollMan\"

    '''<summary>
    ''' create class instance
    ''' </summary>
    Shared Sub New()

    End Sub

    '''<summary>
    ''' Printer used to print collections
    ''' </summary>
    Public Shared Property CollectionsPrinter() As String
        Get
            Return strPrinterCCPlain
        End Get
        Set(ByVal Value As String)
            strPrinterCCPlain = Value
        End Set
    End Property



    '''<summary>
    ''' Name of application
    ''' </summary>
    Public Shared ReadOnly Property ApplicationName() As String
        Get
            'Return System.Reflection.Assembly.GetCallingAssembly.GetName.Name
            Return System.Reflection.Assembly.GetEntryAssembly.GetName.Name
        End Get
    End Property

    '''<summary>
    ''' location for collman log files
    ''' </summary>
    Public Shared Property CollmanLogRoot() As String
        Get
            Return strCollmanRoot
        End Get
        Set(ByVal Value As String)
            strCollmanRoot = Value
        End Set
    End Property

End Class
#End Region

Namespace MedscreenExceptions
#Region "Exceptions"

    '''<summary>
    ''' Specific exceptions that can be thrown
    ''' </summary>
    Public Class MedscreenException
        Inherits Exception

        '''<summary>
        ''' create new exception
        ''' </summary>
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class

    Public Class MedscreenLoginCheckTimeoutException
        Inherits MedscreenException
        Public Sub New()
            MyBase.New("SOAP LOGIN  service not accesible")
        End Sub

        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    Public Class MedscreenTestException
        Inherits MedscreenException
        Public Sub New()
            MyBase.New("Test Exception")
        End Sub

        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    '''<summary>
    ''' This exception is thrown when Oracle is unavailable
    ''' </summary>
    ''' <remarks>This exception is intended to be handled at the application top level 
    ''' to prevent a user getting to far into the application when Oracle is not available
    ''' </remarks>
    Public Class OracleFailure
        Inherits MedscreenException

        '''<summary>
        ''' Throw Oracle has failed exception
        ''' </summary>
        Public Sub New()
            MyBase.New("Oracle has failed")
        End Sub

        '''<summary>
        ''' Overload allowing a different message
        ''' </summary>
        ''' <param name='Message'>Message to include</param>
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : MedscreenExceptions.ScanFileException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Scanfile error exception
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[Taylor]	26/02/2009	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Class ScanFileException
        Inherits CollectionException

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create new scanfile error
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New("Error with scanfile, probably the scanfile is open")
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' New error with filename
        ''' </summary>
        ''' <param name="Filename"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <history>
        ''' 	[Taylor]	26/02/2009	Created
        ''' </history>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal Filename As String)
            MyBase.New("Error with scanfile, probably the scanfile is open : check " & Filename)

        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : MedscreenExceptions.NoIdentityException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Generic exception to be thrown if  no ID can be found
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [04/06/2010]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class NoIdentityException
        Inherits MedscreenException
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class

    '''<summary>
    ''' Generic collection exceptions 
    ''' </summary>
    Public Class CollectionException
        Inherits MedscreenException

        '''<summary>
        ''' create new Collection Exception 
        ''' </summary>
        ''' <param name='ErrorMessage'>Message to include</param>
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : MedscreenExceptions.DuplicateEntryException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Handle duplicate entries 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class DuplicateEntryException
        Inherits MedscreenException
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Throw a duplicate customer exception 
        ''' </summary>
        ''' <param name="ErrorMessage"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/03/2007]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub

    End Class


    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : MedscreenExceptions.DuplicateCustomerException
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A duplicate customer has tried to be created
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/03/2007]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class DuplicateCustomerException
        Inherits DuplicateEntryException

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create a duplicate customer exception, return SMID-Profile as part of the exception 
        ''' </summary>
        ''' <param name="ErrorMessage"></param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [29/03/2007]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class
    '
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : MedscreenExceptions.CanNotChangeCollectionStatus
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Exception to be thrown if can not change collection status 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class CanNotChangeCollectionStatus
        Inherits CollectionException

        '''<summary>
        ''' create new Collection Exception (Can not change status)
        ''' </summary>
        ''' <param name='ErrorMessage'>Message to include</param>
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class

    '''<summary>
    ''' IoException
    ''' </summary>
    Public Class MedscreenIOException
        Inherits MedscreenException

        '''<summary>
        ''' create new IoException
        ''' </summary>
        ''' <param name='ErrorMessage'>Message to include</param>
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub
    End Class

    '''<summary>
    ''' WordOutput problems
    ''' </summary>
    Public Class WordOutputException
        Inherits MedscreenIOException

        '''<summary>
        ''' create new WordOutputException
        ''' </summary>
        ''' <param name='ErrorMessage'>Message to include</param>
        Public Sub New(ByVal ErrorMessage As String)
            MyBase.New(ErrorMessage)
        End Sub

        '''<summary>
        ''' create new WordOutputException
        ''' </summary>
        Public Sub New()
            MyBase.New("Wordoutput directory not accesible")
        End Sub
    End Class
#End Region

End Namespace

Namespace ts01
    Partial Public Class Service1
        Public Sub New(ByVal myURL As String)
            MyBase.New()
            Me.Url = myURL
            If (Me.IsLocalFileSystemWebService(Me.Url) = True) Then
                Me.UseDefaultCredentials = True
                Me.useDefaultCredentialsSetExplicitly = False
            Else
                Me.useDefaultCredentialsSetExplicitly = True
            End If
        End Sub
    End Class
End Namespace

Namespace ts011
    Partial Public Class CardiffService
        Public Sub New(ByVal myURL As String)
            MyBase.New()
            Me.Url = myURL
            If (Me.IsLocalFileSystemWebService(Me.Url) = True) Then
                Me.UseDefaultCredentials = True
                Me.useDefaultCredentialsSetExplicitly = False
            Else
                Me.useDefaultCredentialsSetExplicitly = True
            End If
        End Sub
    End Class
End Namespace
