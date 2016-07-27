
Option Strict Off
Option Explicit On 
Imports System.Windows.Forms

#Region "BallonTip"
''' <summary>
''' -----------------------------------------------------------------------------
''' Project	 : MedscreenCommonGui
''' Class	 :  CBalloonToolTip
''' -----------------------------------------------------------------------------
'''     
''' </summary>
''' <remarks>
'''     
''' </remarks>
''' <revisionHistory>
'''     <revision><Author>taylor</Author><date> 24/01/2007 08:06:40</date>
'''         <Action>$Revision: 1.0 $</Action>
'''     </revision>
''' </revisionHistory>
Public Class CBalloonToolTip

    Private myWindows As Collection

    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
    Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" ( _
        ByVal dwExStyle As Integer, _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String, _
        ByVal dwStyle As Integer, _
        ByVal X As Integer, _
        ByVal Y As Integer, _
        ByVal nWidth As Integer, _
        ByVal nHeight As Integer, _
        ByVal hWndParent As Integer, _
        ByVal hMenu As Integer, _
        ByVal hInstance As Integer, _
        ByRef lpParam As Integer) As Integer

    Private Declare Function SendMessageBT Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Integer, _
        ByVal wMsg As Integer, _
        ByVal wParam As Integer, _
        ByRef lParam As TOOLINFO) As Integer

    Private Declare Function SendMessageStr Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Integer, _
        ByVal wMsg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As String) As Integer

    Private Declare Function SendMessageLong Lib "user32.dll" Alias "SendMessageA" ( _
        ByVal hwnd As Integer, _
        ByVal wMsg As Integer, _
        ByVal wParam As Integer, _
        ByVal lParam As Integer) As Integer

    Private Declare Function DestroyWindow Lib "user32.dll" ( _
        ByVal hwnd As Integer) As Integer

    Private Declare Function ClientToScreen Lib "user32.dll" ( _
        ByVal hwnd As Integer, _
        ByRef lpPoint As POINTAPI) As Integer

    'Windows API Constants
    Private Const WM_USER As Short = &H400S
    Private Const WM_SETFONT = &H30
    Private Const CW_USEDEFAULT As Integer = &H80000000

    'Windows API Types
    Private Structure RECT
        Dim left_Renamed As Integer
        Dim top As Integer
        Dim right_Renamed As Integer
        Dim bottom As Integer
    End Structure

    Private Structure POINTAPI
        Dim X As Integer
        Dim Y As Integer
    End Structure

    'Tooltip Window Constants
    Private Const TTS_NOPREFIX As Short = &H2S
    Private Const TTF_TRANSPARENT As Short = &H100S
    Private Const TTF_CENTERTIP As Short = &H2S
    Private Const TTM_ADDTOOLA As Integer = (WM_USER + 4)
    Private Const TTM_ACTIVATE As Integer = WM_USER + 1
    Private Const TTM_UPDATETIPTEXTA As Integer = (WM_USER + 12)
    Private Const TTM_SETMAXTIPWIDTH As Integer = (WM_USER + 24)
    Private Const TTM_SETTIPBKCOLOR As Integer = (WM_USER + 19)
    Private Const TTM_SETTIPTEXTCOLOR As Integer = (WM_USER + 20)
    Private Const TTM_SETTITLE As Integer = (WM_USER + 32)
    Private Const TTS_BALLOON As Short = &H40S
    Private Const TTS_ALWAYSTIP As Short = &H1S
    Private Const TTF_SUBCLASS As Short = &H10S
    Private Const TTF_TRACK As Short = &H20S
    Private Const TTF_IDISHWND As Short = &H1S
    Private Const TTM_SETDELAYTIME As Integer = (WM_USER + 3)
    Private Const TTDT_AUTOPOP As Short = 2
    Private Const TTDT_INITIAL As Short = 3
    Private Const TTM_TRACKACTIVATE As Integer = WM_USER + 17
    Private Const TTM_TRACKPOSITION As Integer = WM_USER + 18
    Private Const WS_POPUP As Integer = &H80000000

    Private Const TOOLTIPS_CLASSA As String = "tooltips_class32"

    ''Tooltip Window Types
    Private Structure TOOLINFO
        Dim lSize As Integer
        Dim lFlags As Integer
        Dim hwnd As Integer
        Dim lId As Integer
        Dim lpRect As RECT
        Dim hInstance As Integer
        Dim lpStr As String
        Dim lParam As Integer
    End Structure

    Public Enum ttIconType
        TTNoIcon = 0
        TTIconInfo = 1
        TTIconWarning = 2
        TTIconError = 3
    End Enum

    Public Enum ttStyleEnum
        TTStandard
        TTBalloon
    End Enum

    'local variable(s) to hold property value(s)
    Private m_BackColor As Integer
    Private m_Title As String
    Private m_ForeColor As Integer
    Private m_Icon As ttIconType
    Private m_Centered As Boolean
    Private m_Style As ttStyleEnum
    Private m_TipText As String
    Private m_VisibleTime As Integer
    Private m_DelayTime As Integer
    Private m_PopupOnDemand As Boolean

    'private data
    'Private m_lTTHwnd As Integer ' hwnd of the tooltip
    Private m_TipFont As System.Drawing.Font
    'Private m_lParentHwnd As Integer ' hwnd of the window the tooltip attached to
    Private ti As TOOLINFO

    ''' <summary>
    ''' Created by taylor on ANDREW at 24/01/2007 08:07:28
    '''     Create new ballon tooltip
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>taylor</Author><date> 24/01/2007</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' 
    Public Sub New()
        MyBase.New()
        InitCommonControls()
        m_DelayTime = 500
        m_VisibleTime = 5000
        m_PopupOnDemand = False
        Me.m_ForeColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue)
        Me.m_BackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Wheat)
        myWindows = New Collection()            'Collection of Windows
    End Sub
    Protected Overrides Sub Finalize()
        Destroy()
        MyBase.Finalize()
    End Sub
    '//////////////////////////////////////////////////////
    Public Property VisibleTime(ByVal ParentHwnd As Integer) As Integer
        Get
            Return m_VisibleTime
        End Get
        Set(ByVal Value As Integer)
            m_VisibleTime = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property
            If aTooltip.ToolTipWindow <> 0 Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageLong(aTooltip.ToolTipWindow, TTM_SETDELAYTIME, TTDT_AUTOPOP, m_VisibleTime)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property DelayTime(ByVal ParentHwnd As Integer) As Integer
        Get
            Return m_DelayTime
        End Get
        Set(ByVal Value As Integer)
            m_DelayTime = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property
            If aTooltip.ToolTipWindow <> 0 Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageLong(aTooltip.ToolTipWindow, TTM_SETDELAYTIME, TTDT_INITIAL, m_DelayTime)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property TipFont(ByVal ParentHwnd As Integer) As System.Drawing.Font
        Get
            Return m_TipFont
        End Get
        Set(ByVal Value As System.Drawing.Font)
            m_TipFont = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property
            If aTooltip.ToolTipWindow <> 0 Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageLong(aTooltip.ToolTipWindow, WM_SETFONT, m_TipFont.ToHfont.ToInt32, 1)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property Icon(ByVal ParentHwnd As Integer) As ttIconType
        Get
            Return m_Icon
        End Get
        Set(ByVal Value As ttIconType)
            Dim sysNull As System.DBNull
            m_Icon = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property
            If aTooltip.ToolTipWindow <> 0 And Not (m_Title Is sysNull) And m_Icon <> ttIconType.TTNoIcon Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageStr(aTooltip.ToolTipWindow, TTM_SETTITLE, CInt(m_Icon), m_Title)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property ForeColor(ByVal ParentHwnd As Integer) As Integer
        Get
            Return m_ForeColor
        End Get
        Set(ByVal Value As Integer)
            m_ForeColor = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property
            If aTooltip.ToolTipWindow <> 0 Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageLong(aTooltip.ToolTipWindow, TTM_SETTIPTEXTCOLOR, m_ForeColor, 0)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property Title(ByVal ParentHwnd As Integer) As String
        Get
            Return ti.lpStr
        End Get
        Set(ByVal Value As String)
            m_Title = Value
            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property

            If aTooltip.ToolTipWindow <> 0 And m_Title <> "" And m_Icon <> ttIconType.TTNoIcon Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageStr(aTooltip.ToolTipWindow, TTM_SETTITLE, CInt(m_Icon), m_Title)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property TipText(ByVal ParentHwnd As Integer) As String
        Get
            Return m_TipText
        End Get
        Set(ByVal Value As String)
            m_TipText = Value
            ti.lpStr = Value
            Try

                Dim cpWind As String = "A" & CStr(ParentHwnd)
                Dim aTooltip As TipWindow = myWindows.Item(cpWind)

                If aTooltip Is Nothing Then Exit Property
                If aTooltip.ToolTipWindow <> 0 Then
                    ti.hwnd = ParentHwnd
                    ti.lId = ParentHwnd
                    Dim intRet As Integer = SendMessageBT(aTooltip.ToolTipWindow, TTM_UPDATETIPTEXTA, 0, ti)
                End If
            Catch ex As System.Exception
                ' Handle exception here
                'MedscreenLib.Medscreen.LogError(ex, , "Exception log message")
            End Try
        End Set
    End Property

    '//////////////////////////////////////////////////////
    Public Property PopupOnDemand() As Boolean
        Get
            Return m_PopupOnDemand
        End Get
        Set(ByVal Value As Boolean)
            m_PopupOnDemand = Value
            'If m_lTTHwnd <> 0 Then
            'End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property BackColor(ByVal ParentHwnd As Integer) As Integer
        Get
            Return m_BackColor
        End Get
        Set(ByVal Value As Integer)
            m_BackColor = Value

            Dim cpWind As String = "A" & CStr(ParentHwnd)
            Dim aTooltip As TipWindow = myWindows.Item(cpWind)

            If aTooltip Is Nothing Then Exit Property

            If aTooltip.ToolTipWindow <> 0 Then
                ti.hwnd = ParentHwnd
                ti.lId = ParentHwnd
                SendMessageLong(aTooltip.ToolTipWindow, TTM_SETTIPBKCOLOR, m_BackColor, 0)
            End If
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property Style() As ttStyleEnum
        Get
            Style = m_Style
        End Get
        Set(ByVal Value As ttStyleEnum)
            m_Style = Value
        End Set
    End Property
    '//////////////////////////////////////////////////////
    Public Property Centered() As Boolean
        Get
            Centered = m_Centered
        End Get
        Set(ByVal Value As Boolean)
            m_Centered = Value
        End Set
    End Property

    'X and Y are in Pixel so dont send vbTwips value
    Public Sub Show(ByVal ParentHwnd As Integer, Optional ByRef X As Integer = 0, Optional ByRef Y As Integer = 0, Optional ByRef hWndClient As Integer = 0)

        Dim pt As POINTAPI
        Dim ptTip As Integer
        Dim ret As Integer

        With pt
            .X = X
            .Y = Y
        End With

        Dim cpWind As String = "A" & CStr(ParentHwnd)
        Dim aTooltip As TipWindow = myWindows.Item(cpWind)

        If aTooltip Is Nothing Then Exit Sub

        ret = ClientToScreen(hWndClient, pt)

        ptTip = pt.Y * &H10000
        ptTip = ptTip + pt.X

        ' These two messages will set the position of the tooltip:
        ret = SendMessageLong(aTooltip.ToolTipWindow, TTM_TRACKPOSITION, 0, ptTip)
        ret = SendMessageBT(aTooltip.ToolTipWindow, TTM_TRACKACTIVATE, True, ti)
    End Sub

    Private Function GetWindowID(ByVal WinHandle As Integer) As Integer
        Dim retId As Integer = 0
        Dim aTooltip As TipWindow
        Dim i As Integer
        For i = 1 To myWindows.Count
            aTooltip = myWindows.Item(i)
            If aTooltip.ParentWindow = WinHandle Then
                retId = aTooltip.ToolTipWindow
                Exit For
            End If
        Next
        Return retId

    End Function

    Public Function CreateToolTip(ByVal ParentHwnd As Integer, ByVal TipText As String) As Boolean
        'Need a collection of windows to store the tooltips 

        Dim CPwind As String = "A" & CStr(ParentHwnd)
        Dim m_lTTHwnd As Integer = Me.GetWindowID(ParentHwnd) ' myWindows.Item(CPwind)    'See if we have a window of this ID
        m_TipText = TipText
        Dim lWinStyle As Integer
        If m_lTTHwnd <> 0 Then
            DestroyWindow(m_lTTHwnd)
            myWindows.Remove(CPwind)
        End If

        Dim aTipwindow As TipWindow

        'create baloon style if desired
        If m_Style = ttStyleEnum.TTBalloon Then lWinStyle = lWinStyle Or TTS_BALLOON

        m_lTTHwnd = CreateWindowEx(0, TOOLTIPS_CLASSA, 0, lWinStyle, 0, 0, 0, 0, ParentHwnd, 0, 0, 0)
        aTipwindow = New TipWindow(ParentHwnd, m_lTTHwnd)
        myWindows.Add(aTipwindow, CPwind)       'Add to collection



        'now set our tooltip info structure
        With ti
            'NOTE: dont incude TTF_SUBCLASS for on demand
            '      if we want it centered, then set that flag
            If m_Centered Then
                If m_PopupOnDemand = False Then
                    .lFlags = TTF_SUBCLASS Or TTF_CENTERTIP Or TTF_IDISHWND
                Else
                    .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_CENTERTIP
                End If
            Else
                If m_PopupOnDemand = False Then
                    .lFlags = TTF_SUBCLASS Or TTF_IDISHWND
                Else
                    .lFlags = TTF_IDISHWND Or TTF_TRACK Or TTF_TRANSPARENT
                End If
            End If

            'set the hwnd prop to our parent control's hwnd
            .hwnd = ParentHwnd
            .lId = ParentHwnd '0
            .hInstance = 0 'VB6.GetHInstance.ToInt32
            .lpStr = m_TipText
            '.lpRect = lpRect
            .lSize = Len(ti)
        End With

        ''add the tooltip structure
        SendMessageBT(m_lTTHwnd, TTM_ADDTOOLA, 0, ti)

        '//Set all other property of tooltip
        Title(ParentHwnd) = m_Title

        If m_BackColor <> 0 Then BackColor(ParentHwnd) = m_BackColor
        If m_ForeColor <> 0 Then ForeColor(ParentHwnd) = m_ForeColor
        If m_VisibleTime <> 0 Then VisibleTime(ParentHwnd) = m_VisibleTime
        If m_DelayTime <> 0 Then DelayTime(ParentHwnd) = m_DelayTime
        If Not (m_TipFont Is Nothing) Then TipFont(ParentHwnd) = m_TipFont
    End Function

    Public Sub Destroy()
        Dim aTipWindow As TipWindow

        Dim i As Integer
        For i = myWindows.Count To 1 Step -1
            aTipWindow = myWindows.Item(i)
            If aTipWindow.ToolTipWindow <> 0 Then
                DestroyWindow(aTipWindow.ToolTipWindow)
            End If
            myWindows.Remove(i)
        Next
    End Sub

    Private Class TipWindow

#Region "Property 'ParentWindow'"
        ''' <summary>
        '''     TODO: Describe this variable here.
        ''' </summary>
        Private m_intParentWindow As Integer = 0
        ''' <summary>
        '''     TODO: Describe this property here.
        ''' </summary>
        Public Property ParentWindow() As Integer
            Get
                Return m_intParentWindow
            End Get
            Set(ByVal Value As Integer)
                m_intParentWindow = Value
            End Set
        End Property

#End Region 'Property 'ParentWindow'

#Region "Property 'ToolTipWindow'"
        ''' <summary>
        '''     TODO: Describe this variable here.
        ''' </summary>
        Private m_intToolTipWindow As Integer = 0
        ''' <summary>
        '''     TODO: Describe this property here.
        ''' </summary>
        Public Property ToolTipWindow() As Integer
            Get
                Return m_intToolTipWindow
            End Get
            Set(ByVal Value As Integer)
                m_intToolTipWindow = Value
            End Set
        End Property

#End Region 'Property 'ToolTipWindow'

        Public Sub New(ByVal ParentWnd As Integer, ByVal TipWindow As Integer)
            Me.m_intParentWindow = ParentWnd
            Me.m_intToolTipWindow = TipWindow
        End Sub
    End Class
End Class
#End Region 'BallonTip