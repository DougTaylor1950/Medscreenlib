' Original code and credits
' Mike Woodring
' http://staff.develop.com/woodring
'
' VB.NET port by Darrell Norton
' http://dotnetjunkies.com/weblog/darrell.norton/
' 

Option Strict On
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.Xml
Imports System.Configuration

' AssemblySettings usage:
'
' If you know the keys you're after, the following is probably
' the most convenient:
'
'   C# code
'      AssemblySettings settings = new AssemblySettings();
'      string someSetting1 = settings["someKey1"];
'      string someSetting2 = settings["someKey2"];
'   VB.NET code
'       Dim settings As AssemblySettings = New AssemblySettings()
'       Dim someSetting1 As String = settings("someKey1")
'       Dim someSetting2 As String = settings("someKey2")
'
'
' If you want to enumerate over the settings (or just as an
' alternative approach), you can do this too:
'
'   C# code
'      IDictionary settings = AssemblySettings.GetConfig();
'      foreach( DictionaryEntry entry in settings )
'      {
'          // Use entry.Key or entry.Value as desired...
'      }
'   VB.NET Code
'       Dim settings As IDictionary = AssemblySettings.GetConfig()
'       Dim entry As DictionaryEntry
'       For Each entry In  settings
'           ' Use entry.Key or entry.Value as desired...
'       Next entry 
'
'
' In either of the above two scenarios, the calling assembly
' (the one that called the constructor or GetConfig) is used
' to determine what file to parse and what the name of the
' settings collection element is.  For example, if the calling
' assembly is c:\foo\bar\TestLib.dll, then the configuration file
' that's parsed is c:\foo\bar\TestLib.dll.config, and the
' configuration section that's parsed must be named <assemblySettings>.
'
' To retrieve the configuration information for an arbitrary assembly,
' use the overloaded constructor or GetConfig method that takes an
' Assembly reference as input.
'
' If your assembly is being automatically downloaded from a web
' site by an "href-exe" (an application that's run directly from a link
' on a web page), then the enclosed web.config shows the mechanism
' for allowing the AssemblySettings library to download the
' configuration files you're using for your assemblies (while not
' allowing web.config itself to be downloaded).
'
' If the assembly you are trying to use this with is installed in, and loaded
' from, the GAC then you'll need to place the config file in the GAC directory where
' the assembly is installed.  On the first release of the CLR, this directory is
' <windir>\assembly\gac\libName\verNum__pubKeyToken]]>.  For example,
' the assembly "SomeLib, Version=1.2.3.4, Culture=neutral, PublicKeyToken=abcd1234"
' would be installed to the c:\winnt\assembly\gac\SomeLib\1.2.3.4__abcd1234 diretory
' (assuming the OS is installed in c:\winnt).  For future versions of the CLR, this
' directory scheme may change, so you'll need to check the <code>CodeBase</code> property
' of a GAC-loaded assembly in the debugger to determine the correct directory location.

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : AssemblySettings
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Class to identify settings for a single assembly 
''' </summary>
''' <remarks>
''' based on code found at http://www.bearcanyon.com/dotnet/#AssemblySettings translation from C#
''' </remarks>
''' <history>
''' 	[taylor]	17/03/2008	Created
''' </history>
''' -----------------------------------------------------------------------------
Public Class AssemblySettings

#Region "Shared"
    Protected Const defaultSection As String = "assemblySettings"

    Public Overloads Shared Function GetConfig() As IDictionary
        Return GetConfig([Assembly].GetCallingAssembly())
    End Function

    Public Overloads Shared Function GetConfig(ByVal sectionName As String) As IDictionary
        Return GetConfig([Assembly].GetCallingAssembly(), sectionName)
    End Function

    Public Overloads Shared Function GetConfig(ByVal asm As [Assembly]) As IDictionary
        Return GetConfig(asm, defaultSection)
    End Function

    Public Overloads Shared Function GetConfig(ByVal asm As [Assembly], ByVal sectionName As String) As IDictionary
        ' Open and parse configuration file for specified
        ' assembly, returning collection to caller for future
        ' use outside of this class.
        '
        Try
            Dim cfgFile As String = asm.CodeBase + ".config"

            Dim doc As New XmlDocument()
            doc.Load(New XmlTextReader(cfgFile))

            Dim nodes As XmlNodeList = doc.GetElementsByTagName(sectionName)

            Dim node As XmlNode
            For Each node In nodes
                If node.LocalName = sectionName Then
                    Dim handler As New DictionarySectionHandler()
                    Return CType(handler.Create(Nothing, Nothing, node), IDictionary)
                End If
            Next node
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return Nothing

    End Function
#End Region

    Public Function Count() As Integer
        Return settings.Count
    End Function

    Private settings As IDictionary

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Default constructor 
    ''' </summary>
    ''' <param name="asm"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <history>
    ''' 	[taylor]	17/03/2008	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyClass.New([Assembly].GetCallingAssembly())
    End Sub

    Public Sub New(ByVal sectionName As String)
        MyClass.New([Assembly].GetCallingAssembly(), sectionName)
    End Sub

    Public Sub New(ByVal asm As [Assembly])
        settings = GetConfig(asm)
    End Sub

    Public Sub New(ByVal asm As [Assembly], ByVal sectionName As String)
        settings = GetConfig(asm, sectionName)
    End Sub

    Default Public ReadOnly Property Item(ByVal key As String) As String
        Get
            Dim settingValue As String = String.Empty

            If Not settings Is Nothing Then
                settingValue = DirectCast(settings.Item(key), String)
                If settingValue Is Nothing Then
                    settingValue = String.Empty
                End If
            End If

            Return settingValue

        End Get
    End Property



End Class
