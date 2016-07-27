Namespace Exceptions
  <CLSCompliant(True)> _
  Public MustInherit Class MedException
    Inherits Exception

    Public Shared Function ShowException(ByVal ex As MedException, ByVal title As String) As Windows.Forms.DialogResult
      If Not ex Is Nothing Then
        Dim icon As System.Windows.Forms.MessageBoxIcon = Windows.Forms.MessageBoxIcon.None
        Dim buttons As System.Windows.Forms.MessageBoxButtons = Windows.Forms.MessageBoxButtons.OK
        Select Case ex.Severity
          Case MessageSeverity.Information
            icon = Windows.Forms.MessageBoxIcon.Information
          Case MessageSeverity.Warning, MessageSeverity.Resolve
            icon = Windows.Forms.MessageBoxIcon.Exclamation
          Case MessageSeverity.Critical, MessageSeverity.Fatal
            icon = Windows.Forms.MessageBoxIcon.Stop
          Case MessageSeverity.Acknowledge
            ' Use YesNo so that Yes can be distinguished from OK for other message severities
            icon = Windows.Forms.MessageBoxIcon.Question
            buttons = Windows.Forms.MessageBoxButtons.YesNo
        End Select
        Return System.Windows.Forms.MessageBox.Show(ex.Message, title, buttons, icon)
      End If
    End Function

    Protected m_severity As MessageSeverity

    Protected Friend Sub New()
      MyBase.New()
    End Sub

    Protected Friend Sub New(ByVal message As String)
      MyBase.New(message)
    End Sub

    Protected Friend Sub New(ByVal message As String, ByVal innerException As Exception)
      MyBase.New(message, innerException)
      m_severity = MessageSeverity.Fatal
    End Sub

    Protected Friend Sub New(ByVal message As String, ByVal severity As MessageSeverity)
      MyBase.New(message)
      Me.Severity = severity
    End Sub

    Protected Friend Sub New(ByVal message As String, ByVal innerException As Exception, ByVal severity As MessageSeverity)
      MyBase.New(message, innerException)
      Me.Severity = severity
    End Sub

    Public Property Severity() As MessageSeverity
      Get
        Return m_severity
      End Get
      Protected Set(ByVal value As MessageSeverity)
        m_severity = value
      End Set
    End Property

  End Class

  Public Enum MessageSeverity
    None
    ''' <summary>
    '''   Message is for information only.
        ''' </summary>
    Information
    ''' <summary>
    '''   Possible problem, should be checked.
        ''' </summary>
    Warning
    ''' <summary>
    '''   Message must be acknowledged, or problem fixed.
        ''' </summary>
    Acknowledge
    ''' <summary>
    '''   Error must be resolved before continuing
        ''' </summary>
    Resolve
    ''' <summary>
    '''   Critical error (e.g. database update failed)
        ''' </summary>
    Critical
    ''' <summary>
    '''   Fatal error (cannot be resolved - may cause further problems)
        ''' </summary>
    Fatal
  End Enum


  ''' -----------------------------------------------------------------------------
  ''' Project	 : MedscreenLib
  ''' Class	 : ExceptionCollection
  ''' 
  ''' -----------------------------------------------------------------------------
  ''' <summary>
  '''   Class to hold a collection of Exceptions.
  ''' </summary>
  ''' <remarks>
  '''   This class is intended to work with the MedException class to provide 
  '''   severity information.  All other exception types are considered to have 
  '''   a severity of Fatal.
  ''' </remarks>
  ''' <history>
  ''' 	[Boughton]	02/10/2007	Created
  ''' </history>
  ''' -----------------------------------------------------------------------------
  <CLSCompliant(True)> _
Public Class ExceptionCollection
    Inherits CollectionBase

    Public Shared Function GetSeverity(ByVal ex As Exception) As MessageSeverity
      If ex Is Nothing Then
        Return MessageSeverity.None
      ElseIf TypeOf ex Is MedException Then
        Return DirectCast(ex, MedException).Severity
      Else
        Return MessageSeverity.Fatal
      End If
    End Function

    Sub New()
      MyBase.New()
    End Sub

    Public Sub Add(ByVal ex As Exception)
      If Not ex Is Nothing Then
        MyBase.InnerList.Add(ex)
      End If
    End Sub

    Public Sub AddRange(ByVal exceptions As ExceptionCollection)
      If exceptions IsNot Nothing Then
        MyBase.InnerList.AddRange(exceptions)
      End If
    End Sub

    Public Sub Remove(ByVal ex As Exception)
      MyBase.InnerList.Remove(ex)
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    '''   Gets the maximum (worst) severity of all exceptions in the collection
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    '''   Exceptions that do not inherit from AccessionException as considered Fatal.
    ''' </remarks>
    ''' <history>
    ''' 	[Boughton]	16/08/2007	Created
    ''' </history>
    ''' -----------------------------------------------------------------------------
    Public ReadOnly Property MaxSeverity() As MessageSeverity
      Get
        Return ExceptionCollection.GetSeverity(Me.MaxItem)
      End Get
    End Property

    Default Public ReadOnly Property Item(ByVal index As Integer) As Exception
      Get
        Return DirectCast(MyBase.InnerList.Item(index), Exception)
      End Get
    End Property

    Public ReadOnly Property MaxItem() As Exception
      Get
        Dim maxEx As Exception = Nothing
        Dim ex As Exception
        For Each ex In Me
          If ExceptionCollection.GetSeverity(ex) > ExceptionCollection.GetSeverity(maxEx) Then
            maxEx = ex
          End If
        Next
        Return maxEx
      End Get
    End Property

    Public Function FullMessageText() As String
      Dim text As New System.Text.StringBuilder()
      Dim ex As Exception
      For Each ex In Me
        text.AppendFormat("({0}) {1}{2}", MedscreenLib.EnumToString(ExceptionCollection.GetSeverity(ex)), ex.Message, ControlChars.NewLine)
      Next
      Return text.ToString
    End Function

  End Class
End Namespace
