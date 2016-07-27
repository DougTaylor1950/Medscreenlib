Imports System.Diagnostics
Imports System.IO
Imports System
Imports System.Xml.Linq
'''<summary>
''' a file writer utility class that uses the System.IO  library 
''' to write to the filesystem It also implements the ILog interface.
'''  Clients can then use either logger class  interchangeably.
''' </summary>
<CLSCompliant(True)> _
Public Class FileLogger
  
    Implements MedscreenLib.ILog

    'Constants
    Const DEFAULT_FILE_NAME As String = "Debug.txt"
    Const SEVERITY_HEADER As String = " SEVERITY: "
    Const MESSAGE_HEADER As String = " MESSAGE: "

    'Class level variables
    Dim mFileName As String = DEFAULT_FILE_NAME
    Dim mPath As String = Environment.CurrentDirectory


    '********************************************************
    '* Constructors
    '********************************************************

    '''<summary>
    ''' Default constructor
    ''' </summary>
    Public Sub New()

    End Sub

    '''<summary>
    '''Constructor to specify filename
    '''  </summary>
    ''' <param name='FileName'>File to use in the filesystem</param>
    Public Sub New(ByVal FileName As String)

        mFileName = FileName
        If InStr(FileName, "\") > 0 Then
            mPath = Path.GetDirectoryName(mFileName)
            mFileName = Path.GetFileName(mFileName)
        End If

    End Sub

    '''<summary>
    ''' Constructor to specify filename and path
    ''' </summary>
    ''' <param name='FileName'>File to use in the filesystem</param>
    ''' <param name='Path'>Path to File</param>
    Public Sub New(ByVal FileName As String, ByVal Path As String)

        mFileName = FileName
        mPath = Path
    End Sub

    '********************************************************
    '* Methods
    '********************************************************

    ''' <developer>CONCATENO\taylor</developer>
    ''' <summary>
    ''' Write an XML Element
    ''' </summary>
    ''' <param name="OutXML"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory><revision><created>26-Aug-2015 08:58</created><Author>CONCATENO\taylor</Author></revision></revisionHistory>
    Public Overloads Sub WriteXML(ByVal OutXML As XElement)
        Dim OutputStream As StreamWriter = Nothing
        Try
            'create the StreamWriter class which is part of the System.IO namespace
            OutputStream = New StreamWriter(mPath + Path.DirectorySeparatorChar + mFileName, True, System.Text.Encoding.Default)

            Dim tmNow As DateTime = DateTime.Now
            OutXML.Add(New XElement("TS", tmNow))

            OutputStream.WriteLine(OutXML.ToString)
        Catch ex As Exception
            'catch a general exception and pass back to caller
            MsgBox(ex.ToString)
            'Throw ex
        Finally
            'regardless of what happens we want to close the stream
            '<Added code modified 16-Apr-2007 10:11 by taylor> 
            ' user can have insufficient rights to write message so no stream
            If Not OutputStream Is Nothing Then OutputStream.Close()

            '</Added code modified 16-Apr-2007 10:11 by taylor> 
        End Try

    End Sub

    ''' <developer></developer>
    ''' <summary>
    ''' Write an XML element
    ''' </summary>
    ''' <param name="OutXML"></param>
    ''' <param name="Action"></param>
    ''' <remarks></remarks>
    ''' <revisionHistory></revisionHistory>
    Public Overloads Sub WriteXML(ByVal OutXML As String, ByVal Action As String)
        Dim OutputStream As StreamWriter = Nothing
        Try
            'create the StreamWriter class which is part of the System.IO namespace
            OutputStream = New StreamWriter(mPath + Path.DirectorySeparatorChar + mFileName, True, System.Text.Encoding.Default)
            Dim tmNow As DateTime = DateTime.Now
            Dim strTimeStamp As String = "<" & Action & "><TIME>" & tmNow.ToString("dd-MMM-yyyy HH:mm:ss.") & CStr(Date.Now.Millisecond().ToString("000")) & "</TIME>" & OutXML

            OutputStream.WriteLine(strTimeStamp & "<" & Action & ">")
        Catch ex As Exception
            'catch a general exception and pass back to caller
            MsgBox(ex.ToString)
            'Throw ex
        Finally
            'regardless of what happens we want to close the stream
            '<Added code modified 16-Apr-2007 10:11 by taylor> 
            ' user can have insufficient rights to write message so no stream
            If Not OutputStream Is Nothing Then OutputStream.Close()

            '</Added code modified 16-Apr-2007 10:11 by taylor> 
        End Try

    End Sub


    'This is the implementation of the ILog.WriteLog method
    '''<summary>
    ''' WriteLog is derived from the ILog interface
    ''' It must be defined in all classes that implement the ILog Interface
    ''' </summary>
    ''' <param name='OutStream'>Stream to write to</param>
    ''' <param name='LogType'>Event Log Entry Type</param>
    Public Sub WriteLog(ByVal OutStream As String, ByVal LogType As EventLogEntryType) _
        Implements MedscreenLib.ILog.WriteLog

        Dim OutputStream As StreamWriter = Nothing
        Try
            'create the StreamWriter class which is part of the System.IO namespace
            OutputStream = New StreamWriter(mPath + Path.DirectorySeparatorChar + mFileName, True, System.Text.Encoding.Default)
            Dim tmNow As DateTime = DateTime.Now
            Dim strTimeStamp As String = tmNow.ToString("dd-MMM-yyyy HH:mm:ss.") & CStr(Date.Now.Millisecond().ToString("000"))
            OutputStream.WriteLine(strTimeStamp & SEVERITY_HEADER & LogType.ToString() & MESSAGE_HEADER & OutStream)
        Catch ex As Exception
            'catch a general exception and pass back to caller
            MsgBox(ex.ToString)
            'Throw ex
        Finally
            'regardless of what happens we want to close the stream
            '<Added code modified 16-Apr-2007 10:11 by taylor> 
            ' user can have insufficient rights to write message so no stream
            If Not OutputStream Is Nothing Then OutputStream.Close()

            '</Added code modified 16-Apr-2007 10:11 by taylor> 
        End Try
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Set the output path for logging
    ''' </summary>
    ''' <param name="path">Path to set to</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [05/10/2005]</date><Action>Added</Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub SetPath(ByVal path As String)
        mPath = path
    End Sub
End Class
