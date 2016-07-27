'$Revision: 1.1 $
'$Author: taylor $
'$Date: 2006-04-11 10:52:53+01 $
'$Log: EventLogger.vb,v $
'Revision 1.1  2006-04-11 10:52:53+01  taylor
'<>
'
'Revision 1.1  2005-09-09 09:21:01+01  taylor
'Comments added for NDoc
'

'<>Imports System.Diagnostics
Imports System


'''<summary>
''' Event writer utility class that uses the System.diagnostics 
''' library to write to the Event Log.
''' It also implements the ILog interface.
''' Clients can then use either logger class 
''' interchangeably.</summary>
''' <remarks>
''' We use implements statement to implement the iLog interface.
''' Note the use of the fully qualified interface name &lt;namespace&gt;.&lt;interfacename&gt;
''' </remarks>
Public Class EventLogger
    
    Implements MedscreenLib.ILog

    'constants
    Const EVENT_LOG As String = "Application"
    Const DEFAULT_SOURCE As String = "Custom Application"

    'class-level variables
    Dim mSource As String = DEFAULT_SOURCE


    
    '''<summary>
    ''' default constructor
    ''' </summary>
    Public Sub New()

    End Sub


    '''<summary>
    ''' this constructor allows client to specify an event source
    ''' </summary>
    ''' <param name='Source'>Event Source</param>
    Public Sub New(ByVal Source As String)
        mSource = Source
    End Sub

    '********************************************************
    '* Methods
    '********************************************************

    '''<summary>
    ''' WriteLog is derived from the ILog interface
    ''' It must be defined in all classes that implement the ILog Interface
    ''' </summary>
    ''' <param name='Info'>Information to log</param>
    ''' <param name='LogType'>Event Log Entry Type</param>
    Public Sub WriteLog(ByVal Info As String, ByVal LogType As EventLogEntryType) _
        Implements MedscreenLib.ILog.WriteLog

        Try
            'if it doesn't already exist create a source for the event to be logged to
            If Not (EventLog.SourceExists(mSource)) Then
                EventLog.CreateEventSource(mSource, EVENT_LOG)
            End If
            'Create a new event log object
            Dim Log As EventLog = New EventLog()
            'set the source property
            Log.Source = mSource
            'log the entry
            Log.WriteEntry(Info, LogType)
        Catch ex As Exception
            'catch a general exception and pass back to caller
            Throw ex
        End Try
    End Sub


End Class
