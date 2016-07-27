Option Strict On
Option Explicit On 
Namespace IniFile

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : IniFile.IniElement
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An single element from an IniFile used with IniCollection <see cref="IniFile.IniCollection"/>
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class IniElement
      
#Region "Declarations"
        Private strHeader As String
        Private strItem As String
#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"

#End Region

#Region "Properties"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Key Header 
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Header() As String
            Get
                Return Me.strHeader
            End Get
            Set(ByVal Value As String)
                Me.strHeader = Value
            End Set
        End Property

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Key Value
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property Item() As String
            Get
                Return Me.strItem
            End Get
            Set(ByVal Value As String)
                Me.strItem = Value
            End Set
        End Property
#End Region
#End Region

   
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : IniFile.IniCollection
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A collection of key value pairs used for the section load 
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class IniCollection
        Inherits CollectionBase

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Constructor creates a new collection
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.New()
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Add an IniElement to the collection
        ''' </summary>
        ''' <param name="item">IniElement to add</param>
        ''' <returns>Position of added item collection</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Add(ByVal item As IniElement) As Integer
            Return MyBase.List.Add(item)
        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Retrieve an Item from the collection
        ''' </summary>
        ''' <param name="Index">Index of Item to retrieve</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function Item(ByVal Index As Integer) As IniElement
            Return CType(MyBase.List.Item(Index), IniElement)
        End Function
    End Class

    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : IniFile.IniFiles
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' A class emulating Borland's approach to accessing INI files
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class IniFiles


#Region "External Function Declaration"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Windows built in function
        ''' </summary>
        ''' <param name="lpApplicationName">Long Pointer to Section name</param>
        ''' <param name="lpKeyName">Long Pointer to Key name</param>
        ''' <param name="lpDefault">Long Pointer to Default value</param>
        ''' <param name="lpReturnedString">Long Pointer to Returned Value</param>
        ''' <param name="nSize">Size of Returned sting</param>
        ''' <param name="lpFileName">Long Pointer to Filename</param>
        ''' <returns>0 = success</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a section from a profile file 
        ''' </summary>
        ''' <param name="lpAppName">Long Pointer to Section name</param>
        ''' <param name="lpReturnedString">Long Pointer to returned string terminated with two null strings, separated by null string</param>
        ''' <param name="nSize">Size of returned string</param>
        ''' <param name="lpFileName">Long Pointer to filename</param>
        ''' <returns>A string composed of each of the keys in the section, separated by nul strings, terminated by two null strings</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get an integer from the inifile
        ''' </summary>
        ''' <param name="lpApplicationName">Long Pointer to section</param>
        ''' <param name="lpKeyName">Long Pointer to key </param>
        ''' <param name="nDefault">Default return value</param>
        ''' <param name="lpFileName">Long Pointer to filename</param>
        ''' <returns>Value obtained or default</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer

        'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: ms-help://MS.MSDNVS/vbcon/html/vbup1016.htm
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write out a string to an inifile
        ''' </summary>
        ''' <param name="lpApplicationName">Long Pointer to section</param>
        ''' <param name="lpKeyName">Long Pointer to key</param>
        ''' <param name="lpString">Long Pointer to string to write</param>
        ''' <param name="lpFileName">Long Pointer to filename</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Declare Function WritePrivateProfileString Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
        ByVal lpKeyName As String, ByVal lpString As String, _
        ByVal lpFileName As String) As Integer


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write out a section to the inifile
        ''' </summary>
        ''' <param name="lpAppName">Long Pointer to Section</param>
        ''' <param name="lpString">Long Pointer to string contaning section</param>
        ''' <param name="lpFileName">Long Pointer to file name</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
#End Region


#Region "Public Constants"

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' This constant preserves key when system is rebooted.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action>Commented</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_OPTION_NON_VOLATILE As Short = 0


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 'This constant does NOT preserve key when system is rebooted.
        ''' Use this to write temporary values to the registry.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_OPTION_VOLATILE As Short = 1


        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' New Registry Key created.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_CREATED_NEW_KEY As Short = &H1S

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Existing Key opened.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_OPENED_EXISTING_KEY As Short = &H2S

        '

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Free form binary.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_BINARY As Short = 3

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 32-bit number.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_DWORD As Short = 4

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 32-bit number.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_DWORD_BIG_ENDIAN As Short = 5

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 32-bit number (same as REG_DWORD).
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_DWORD_LITTLE_ENDIAN As Short = 4

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Unicode nul terminated string.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_EXPAND_SZ As Short = 2

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Symbolic Link (unicode).
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_LINK As Short = 6

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Multiple Unicode strings.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_MULTI_SZ As Short = 7



        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Resource list in the resource map.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_RESOURCE_LIST As Short = 8

        '
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Unicode nul terminated string.
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const REG_SZ As Short = 1

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Error return Success
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const ERROR_SUCCESS As Short = 0

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Error Access Denied 
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Const ERROR_ACCESS_DENIED As Short = 5

#End Region

        Private Sub RaiseError(ByVal intNumber As Integer, ByVal strBit As String)

        End Sub

#Region "Declarations"
        Dim NL2 As String = ControlChars.NewLine & ControlChars.NewLine
        Private mvarIniFileName As String
        Const cstModuleName As String = "ModRegistry"


#End Region

#Region "Public Instance"

#Region "Functions"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get a string from the a profile 
        ''' </summary>
        ''' <param name="strSectionName">Section in the Ini file</param>
        ''' <param name="strKey">Key to access</param>
        ''' <param name="strFilename">Name of Inifile</param>
        ''' <returns>A string containing the information - Unknown if it fails</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function GetProfileString(ByRef strSectionName As String, ByRef strKey As String, ByRef strFilename As String) As String

            Dim lRet As Integer
            Dim nSize As Integer
            Dim retString As String

            nSize = 300
            retString = Space(nSize)

            lRet = GetPrivateProfileString(strSectionName, strKey, "Unknown", retString, nSize, strFilename)
            GetProfileString = Mid(retString, 1, lRet)

        End Function

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Get the contents of a section
        ''' </summary>
        ''' <param name="strSectionName">Name Of Section to retrieve</param>
        ''' <param name="strFilename">Name of Inifile</param>
        ''' <returns>Section as a string</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function GetProfileSection(ByRef strSectionName As String, ByRef strFilename As String) As String

            Dim lRet As Integer
            Dim strReturnString As String
            Dim lngBuffer As Integer

            lngBuffer = 2048
            strReturnString = Space(lngBuffer)
            lRet = GetPrivateProfileSection(strSectionName, strReturnString, lngBuffer, strFilename)

            GetProfileSection = Mid(strReturnString, 1, lRet)

        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read a string from the class handled inifile
        ''' </summary>
        ''' <param name="strSection">Section to read from </param>
        ''' <param name="strItem">Key to read</param>
        ''' <param name="strDefault">Default value to use</param>
        ''' <returns>The string associate dwith key or default value</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:355]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ReadString(ByRef strSection As String, ByRef strItem As String, ByRef strDefault As String) As String
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Integer
            Dim strRet As String = ""
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'

            Try

                If Len(mvarIniFileName) = 0 Then
                    Return strRet
                    Exit Function
                End If

                Buffer = Space(1024)
                ilen = GetPrivateProfileString(strSection, strItem, strDefault, Buffer, 1024, mvarIniFileName)
                If ilen > 0 Then
                    strRet = Left(Buffer, ilen)
                Else
                    strRet = strDefault
                End If
            Catch ex As Exception
                Dim strErr As String = ex.Message
            End Try
            Return strRet
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read a section from class inifile return as a VB6 collection
        ''' </summary>
        ''' <param name="strSection">Section to retrieve</param>
        ''' <returns>Section as a collection</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:30]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function ReadSection(ByRef strSection As String) As Collection
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Integer
            Dim mCol As Collection
            Dim strItemName As String
            Dim iPos As Integer
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            On Error GoTo ErrorManager

            mCol = New Collection()
            If Len(mvarIniFileName) = 0 Then
                Return mCol
                Exit Function
            End If



            Buffer = Space(1024)
            ilen = GetPrivateProfileSection(strSection, Buffer, 1023, mvarIniFileName)

            iPos = 1
            Do While (iPos < ilen)
                strItemName = ""
                Do While Mid(Buffer, iPos, 1) <> Chr(0)
                    strItemName = strItemName & Mid(Buffer, iPos, 1)
                    iPos = iPos + 1
                Loop
                mCol.Add(strItemName)
                iPos = iPos + 1
            Loop
            ReadSection = mCol

ExitReadSection:
            Exit Function

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - ReadSection")
            Resume ExitReadSection
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read a section filling an IniCollection <see cref="IniFile.IniCollection"/>
        ''' </summary>
        ''' <param name="strSection">Section to retrieve</param>
        ''' <returns>returns an IniCollection</returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Overloads Function ReadSectionCollection(ByRef strSection As String) As IniCollection
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Integer
            Dim mCol As IniCollection
            Dim strItemName As String
            Dim strHeader As String
            Dim objElem As IniElement
            Dim iPos As Integer
            Dim blnHeader As Boolean
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            On Error GoTo ErrorManager

            mCol = New IniCollection()
            If Len(mvarIniFileName) = 0 Then
                Return mCol
                Exit Function
            End If


            Buffer = Space(1024)
            ilen = GetPrivateProfileSection(strSection, Buffer, 1023, mvarIniFileName)

            iPos = 1
            Do While (iPos < ilen)
                strItemName = ""
                strHeader = ""
                blnHeader = True
                Do While Mid(Buffer, iPos, 1) <> Chr(0)
                    If blnHeader Then
                        If Mid(Buffer, iPos, 1) = "=" Then
                            blnHeader = False
                            iPos += 1
                        Else
                            strHeader += Mid(Buffer, iPos, 1)
                        End If
                    End If
                    If Not blnHeader Then
                        strItemName += Mid(Buffer, iPos, 1)
                    End If
                    iPos = iPos + 1
                Loop
                objElem = New IniElement()
                objElem.Header = strHeader
                objElem.Item = strItemName
                mCol.Add(objElem)
                iPos = iPos + 1
            Loop
            Return mCol

ExitReadSection:
            Exit Function

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - ReadSection")
            Resume ExitReadSection
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read an integer from current Class file
        ''' </summary>
        ''' <param name="strSection">Section to read from</param>
        ''' <param name="strItem">Key to read</param>
        ''' <param name="intDefault">Default value</param>
        ''' <returns>Default or value read </returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:28]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ReadInteger(ByRef strSection As String, ByRef strItem As String, ByRef intDefault As Integer) As Integer
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim ilen As Integer
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            On Error GoTo ErrorManager

            If Len(mvarIniFileName) = 0 Then Exit Function


            ilen = GetPrivateProfileInt(strSection, strItem, intDefault, mvarIniFileName)

            ReadInteger = ilen

ExitReadInteger:
            Exit Function

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - ReadInteger")
            Resume ExitReadInteger
        End Function


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Read a boolean value from Inifile
        ''' </summary>
        ''' <param name="strSection">Section to read from</param>
        ''' <param name="strItem">Key to read</param>
        ''' <param name="blnDefault">Default value</param>
        ''' <returns>TRUE if value is a true string</returns>
        ''' <remarks>
        ''' the values accepted as TRUE are Y, YES, T, TRUE, irrespective of case
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:22]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function ReadBoolean(ByRef strSection As String, ByRef strItem As String, ByRef blnDefault As Boolean) As Boolean
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Short
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            ReadBoolean = False
            On Error GoTo ErrorManager


            If Len(mvarIniFileName) = 0 Then Exit Function

            Buffer = "N"
            If blnDefault Then Buffer = "Y"

            Buffer = UCase(ReadString(strSection, strItem, Buffer))

            ReadBoolean = ((Buffer = "Y") Or (Buffer = "YES") Or _
                (Buffer = "T") Or (Buffer = "TRUE"))
ExitReadBoolean:
            Exit Function

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - ReadBoolean")
            Resume ExitReadBoolean
        End Function

#End Region

#Region "Procedures"

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write a string back to the Ini file
        ''' </summary>
        ''' <param name="strSectionName">Name of the section</param>
        ''' <param name="strKey">Key within the section</param>
        ''' <param name="strSetting">Value to write</param>
        ''' <param name="strFilename">Filename to write to</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub WriteProfileString(ByRef strSectionName As String, _
            ByRef strKey As String, _
            ByRef strSetting As String, ByRef strFilename As String)
            Dim lRet As Integer

            lRet = WritePrivateProfileString(strSectionName, strKey, strSetting, strFilename)

        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write the string to the inifile being handled by this class
        ''' </summary>
        ''' <param name="strSection">Section to Write to</param>
        ''' <param name="strItem">Key within section</param>
        ''' <param name="strValue">Text to write out</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:32]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub WriteString(ByRef strSection As String, ByRef strItem As String, ByRef strValue As String)
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim ilen As Integer
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'

            If Len(mvarIniFileName) = 0 Then Exit Sub
            Try
                ilen = WritePrivateProfileString(strSection, strItem, strValue, mvarIniFileName)
            Catch ex As Exception
            End Try



        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write an integer to a file
        ''' </summary>
        ''' <param name="strSection">Section to write to</param>
        ''' <param name="strItem">Key to write to </param>
        ''' <param name="intValue">Value to write</param>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:25]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub WriteInteger(ByRef strSection As String, ByRef strItem As String, ByRef intValue As Integer)
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Short
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            On Error GoTo ErrorManager

            If Len(mvarIniFileName) = 0 Then Exit Sub

            Buffer = Str(intValue)

            WriteString(strSection, strItem, Buffer)

ExitWriteInteger:
            Exit Sub

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - WriteInteger")
            Resume ExitWriteInteger
        End Sub


        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Write a boolean to IniFile
        ''' </summary>
        ''' <param name="strSection">Section to write to</param>
        ''' <param name="strItem">Key to write to</param>
        ''' <param name="blnValue">Value to be written</param>
        ''' <remarks>
        ''' Writes Y for TRUE and N for FALSE
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 08:33:16]</date><Action>Created</Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub WriteBool(ByRef strSection As String, ByRef strItem As String, ByRef blnValue As Boolean)
            '*****************************************************'
            '*               Declaration                         *'
            '*****************************************************'
            Dim Buffer As String
            Dim ilen As Short
            '*****************************************************'
            '                Implementation                      *'
            '*****************************************************'
            On Error GoTo ErrorManager

            If Len(mvarIniFileName) = 0 Then Exit Sub

            Buffer = "N"
            If blnValue Then Buffer = "Y"

            WriteString(strSection, strItem, Buffer)
ExitWriteBool:
            Exit Sub

ErrorManager:
            Call RaiseError(Err.Number, cstModuleName & " - WriteBool")
            Resume ExitWriteBool
        End Sub


#End Region

#Region "Properties"

        ' ******************************************************************************
        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' The File name of the Inifile
        ''' </summary>
        ''' <value></value>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [05/05/2000 07:59:24]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Property FileName() As String
            Get
                FileName = mvarIniFileName

            End Get
            Set(ByVal Value As String)
                Dim p As System.Security.Principal.WindowsPrincipal
                Dim s As New System.Security.Permissions.FileIOPermission(Security.Permissions.PermissionState.Unrestricted)
                'Dim dr As System.IO.DirectoryInfo
                'Dim scf As System.IO.DirectoryInfo
                'Dim wic As System.Security.Principal.WindowsImpersonationContext
                Dim strDirectory As String
                Dim iPos As Integer


                mvarIniFileName = Value

                p = New System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent)
                'wic = MyIdentity.Impersonate

                iPos = Value.Length - 1
                While iPos >= 0 And Value.Chars(iPos) <> "\"
                    iPos = iPos - 1
                End While

                strDirectory = Mid(Value, 1, iPos + 1)


                s.AllLocalFiles = Security.Permissions.FileIOPermissionAccess.AllAccess
                s.AllFiles = Security.Permissions.FileIOPermissionAccess.AllAccess
                s.AddPathList(Security.Permissions.FileIOPermissionAccess.AllAccess, strDirectory)
                s.Demand()
                s.Assert()

            End Set
        End Property 'ZFileName

#End Region
#End Region



    End Class

End Namespace

Namespace XmlError
    ''' -----------------------------------------------------------------------------
    ''' Project	 : MedscreenLib
    ''' Class	 : XmlError.XMLApplicationError
    ''' 
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Write an error message as XML
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Class XMLApplicationErrorX
        Private strFilename As String
        Private objFile As System.IO.FileInfo
        Private SW As System.IO.StreamWriter

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' Create Error message object and Open File
        ''' </summary>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Sub New()
            MyBase.new()
            strFilename = "\\corp.concateno.com\medscreen\common\AppError\" & _
                System.Windows.Forms.Application.ProductName & _
                Now.ToString("yyyyMMdd") & ".xml"
            Try
                objFile = New System.IO.FileInfo(strFilename)
                If Not objFile.Exists Then
                    objFile.Create()
                    SW = objFile.AppendText
                    SW.WriteLine("<" & System.Windows.Forms.Application.ProductName & ">")
                Else
                    SW = objFile.AppendText
                End If
            Catch ex As Exception
            End Try
        End Sub

        ''' -----------------------------------------------------------------------------
        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="strItem"></param>
        ''' <param name="strType"></param>
        ''' <param name="strInfo"></param>
        ''' <param name="strBarcode"></param>
        ''' <param name="intOpt"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        ''' <revisionHistory>
        ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
        ''' </revisionHistory>
        ''' -----------------------------------------------------------------------------
        Public Function AddErrorXML(ByVal strItem As String, _
        ByVal strType As String, _
        ByVal strInfo As String, _
        ByVal strBarcode As String, _
        ByVal intOpt As Integer) As Boolean

            Dim strItemVal As String = ""

            If intOpt = 1 Then
                strItemVal = "CUSTOMER"
            End If

            SW.WriteLine("<" & strItemVal & ">" & strItem & "</" & strItemVal & ">" & _
            "<TYPE>" & strType & "</TYPE>" & _
            "<BARCODE>" & strBarcode & "</BARCODE>" & _
            "<INFO>" & strInfo & "</INFO>")
            SW.Flush()
        End Function
    End Class
End Namespace

Namespace Support
    Module modSupport
        Private MyLiveConn As OleDb.OleDbConnection
        Private Const STATUS_LIVE As String = "Live"
        Private Const STATUS_TEST As String = "Testing"

        Private strProvider As String = "OraOLEDB.Oracle.1;Persist Security Info=False"
        Private strUserName As String = "onestop"
        Private strPassword As String = "onestop"
        Private strDataBase As String = "john-orcl"
        Private strStatus As String = STATUS_LIVE

      


    End Module
End Namespace