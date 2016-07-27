Imports System.Drawing

#Region "Graphs"


''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Charts
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Charts class, deals with drwaing a chart
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class Charts

#Region "Declarations"


    Private AxisPen As New Pen(Color.Black, 2)
    Private FontLabels As New Font("verdana", 8)

    Private blnShowValues As Boolean = False
    Private HistGap As Integer = 0
    Private sngLeftMargin As Single = 100
    Private sngRightMargin As Single = 20
    Private sngBotMargin As Single = 25
    Private sngTopMargin As Single = 25
    Private myLeftAxis As Axis
    Private myXAxis As Axis
    Private mySeries As New SeriesCollection()
    Private gr As Graphics
    Private bm As Bitmap
    Private myFileName As String
    Private ChartRect As Rectangle
    Private myRect As Rectangle
    Private myLegend As Legend
    Private myTitle As Title

#End Region

#Region "Public Enumerations"

    '''<summary>Chart Types</summary>
    Public Enum eChartType
        '''<summary>XY Chart</summary>
        XY
        '''<summary>Pie chart</summary>
        Pie
        '''<summary>Spedometer</summary>
        Speedo
        '''<summary>Slider</summary>
        Slider
    End Enum



#End Region

    Private myChartType As eChartType = eChartType.XY

#Region "Public Instance"

#Region "Functions"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a Chart, writing chart into a file for display on a web site
    ''' </summary>
    ''' <param name="Yseries">Chart Y series</param>
    ''' <param name="xSeries">Chart X series</param>
    ''' <param name="Title">Chart Title</param>
    ''' <param name="Width">Chart Width</param>
    ''' <param name="Height">Chart Height</param>
    ''' <param name="strFileName">Filename for output (including path)</param>
    ''' <param name="SeriesName">Description for first series</param>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function CreateChartFile(ByVal Yseries As ArrayList, ByVal xSeries As ArrayList, _
    ByVal Title As String, ByVal Width As Integer, ByVal Height As Integer, _
    ByVal strFileName As String, ByVal SeriesName As String) As Boolean
        bm = New Bitmap(Width, Height)

        myFileName = strFileName
        gr = Graphics.FromImage(bm)


        myRect = New Rectangle(0, 0, Width, Height)
        'Const BufferSpace As Integer = 15

        Dim BlockHeightMax As Single = Height - sngBotMargin - sngTopMargin

        ChartRect = New Rectangle(sngLeftMargin, sngTopMargin, _
        Width - sngLeftMargin - sngRightMargin, BlockHeightMax)
        myLeftAxis = New Axis(Axis.AxisType.LeftY, Me)
        myLeftAxis.Maximum = 100

        XAxis = New Axis(Axis.AxisType.XAxisDate, Me)

        'gr.DrawLine(AxisPen, sngLeftMargin, Height - sngBotMargin, Width - 20, Height - sngBotMargin)

        myLegend = New Legend(Me)
        myTitle = New Title(Me)


        Dim dataseries As New Series(Series.SeriesType.VertHistogram, Me, SeriesName)
        Me.Serieses.Add(dataseries)
        dataseries.Xvalues = xSeries
        dataseries.YValues = Yseries


    End Function

    Private Sub DrawChartLocal()
        Me.XAxis = myXAxis

        gr.FillRectangle(Brushes.White, myRect)
        gr.DrawRectangle(Pens.Black, 0, 0, myRect.Width - 2, myRect.Height - 2)

        Dim BlockHeightMax As Single = myRect.Height - sngBotMargin - sngTopMargin

        Dim ChartRectX As Single = sngLeftMargin
        Dim ChartTop As Single = sngTopMargin
        Dim ChartWidth As Single = myRect.Width
        Dim ChartHeight As Single = BlockHeightMax

        If myLegend.HasLegend Then
            Select Case myLegend.Position
                Case Legend.LegendPosition.Top
                    ChartTop = sngTopMargin + 40
                    ChartWidth = myRect.Width - sngLeftMargin - sngRightMargin
                    ChartHeight = BlockHeightMax - 80
                Case Legend.LegendPosition.Bottom
                    ChartWidth = myRect.Width - sngLeftMargin - sngRightMargin
                    ChartHeight = BlockHeightMax - Serieses.Count * 40
                Case Legend.LegendPosition.Left
                    ChartRectX = sngLeftMargin + myLegend.LegendWidth
                    ChartWidth = myRect.Width - sngLeftMargin - sngRightMargin - myLegend.LegendWidth
                Case Legend.LegendPosition.Right
                    ChartWidth = myRect.Width - sngLeftMargin - sngRightMargin - myLegend.LegendWidth
                    'ChartRect = New Rectangle(sngLeftMargin, sngTopMargin, _
                    'myRect.Width - sngLeftMargin - sngRightMargin - myLegend.LegendWidth, BlockHeightMax)

            End Select

            myLegend.Draw(gr)
            If myTitle.HasTitle Then
                ChartTop = ChartTop + 40
            End If
            ChartRect = New Rectangle(ChartRectX, ChartTop, ChartWidth, ChartHeight)
        Else
            If myTitle.HasTitle Then
                ChartTop = ChartTop + 40
            End If
            ChartRect = New Rectangle(ChartRectX, ChartTop, ChartWidth, ChartHeight)
        End If



        If myChartType = eChartType.XY Then
            myLeftAxis.DrawAxis(gr)
            XAxis.DrawAxis(gr)

        ElseIf myChartType = eChartType.Speedo Then
            myLeftAxis.DrawAxis(gr)
        ElseIf myChartType = eChartType.Slider Then
            myLeftAxis.DrawAxis(gr)
            XAxis.DrawAxis(gr)
            Dim x2 As Single = XAxis.BlockHeight(Convert.ToSingle(Gap))
            Dim BlockHeight As Single = LeftYAxis.BlockHeight(LeftYAxis.Maximum) + ChartArea.Top
            Dim BackBrush As New SolidBrush(Color.Beige)
            Dim x1 As Single = ChartArea.Left
            Dim outerwidth As Single = XAxis.BlockHeight(20)
            gr.FillRectangle(BackBrush, New Rectangle(x2 + HistGap, _
             ChartArea.Top + (ChartArea.Height - BlockHeight), outerwidth, BlockHeight))

        End If

        Dim DataSeries As Series
        For Each DataSeries In Serieses
            If DataSeries.sType = Series.SeriesType.VertHistogram Then
                DataSeries.Maximum = Me.myLeftAxis.Maximum
                DataSeries.Minimum = Me.myLeftAxis.Minimum

            End If
            DataSeries.DrawSeries(gr, FontLabels)
        Next

        If myTitle.HasTitle Then
            myTitle.Draw(gr)
        End If
    End Sub



    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw chart 
    ''' </summary>
    ''' <returns>VOID</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function DrawChart() As Boolean

        DrawChartLocal()

        Try
            bm.Save(myFileName, Imaging.ImageFormat.Jpeg)
        Catch ex As Exception
            Errors = myFileName & ex.ToString
            Return False
        End Try
        Return True

    End Function

    Public Function DrawChartStream() As Byte()
        DrawChartLocal()
        Dim bmpBytes As Byte()
        Dim ms As New IO.MemoryStream()
        bm.Save(ms, Imaging.ImageFormat.Jpeg)
        ms.Capacity = ms.Length
        bmpBytes = ms.GetBuffer()
        'bm.Dispose()
        ms.Close()
        Return bmpBytes
    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a Pie chart
    ''' </summary>
    ''' <param name="data">Pie data</param>
    ''' <param name="labels">Pie labels</param>
    ''' <param name="Title">Chart Title</param>
    ''' <param name="Width">Chart Width</param>
    ''' <returns>Chart Bitmap</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Function CreatePieChart(ByVal data As Array, ByVal labels As Array, _
    ByVal Title As String, ByVal Width As Integer) As Bitmap

        Const BufferSpace As Integer = 15

        Dim nRows As Integer = data.Length
        Dim Total As Single = 0

        Dim i As Integer

        For i = 0 To nRows - 1
            Total += Convert.ToSingle(data.GetValue(i))
        Next
        Dim FontLegend As New Font("verdana", 10)
        Dim FontTitle As New Font("Verdana", 15, FontStyle.Bold)

        Dim LegendHeight As Integer = FontLegend.Height * (nRows + 1) + BufferSpace
        Dim TitleHeight As Integer = FontTitle.Height + BufferSpace
        Dim Height As Integer = LegendHeight + TitleHeight + Width + BufferSpace
        Dim PieHeight As Integer = Width

        Dim pieRect As New Rectangle(0, TitleHeight, Width, PieHeight)

        Dim Colours As New ArrayList()
        Dim rnd As New Random()

        For i = 0 To nRows - 1
            Colours.Add(New SolidBrush(Color.FromArgb(rnd.Next(255), rnd.Next(255), rnd.Next(255))))
        Next
        Dim bm As New Bitmap(Width, Height)

        Dim gr As Graphics = Graphics.FromImage(bm)


        Dim Angle As Single = 0
        Dim oldAngle As Single = 0
        Dim sb As SolidBrush

        gr.FillRectangle(New SolidBrush(Color.White), 0, 0, Width, Height)
        For i = 0 To nRows - 1
            Angle = (data.GetValue(i) / Total * 360)
            sb = CType(Colours(i), SolidBrush)
            Debug.WriteLine(sb.Color.Name & ", " & oldAngle & ", " & Angle)
            gr.FillPie(sb, pieRect, oldAngle, Angle)
            oldAngle += Angle
        Next

        Dim StFormat As New StringFormat()
        StFormat.Alignment = StringAlignment.Center
        StFormat.LineAlignment = StringAlignment.Center

        gr.DrawString(Title, FontTitle, Drawing.Brushes.Black, _
            New RectangleF(0, 0, Width, TitleHeight), StFormat)


        For i = 0 To nRows - 1

            gr.FillRectangle(CType(Colours(i), SolidBrush), 5, _
             Height - LegendHeight + FontLegend.Height * i + 5, 10, 10)
            gr.DrawString((labels(i) + " - " & CSng(data(i))), FontLegend, Brushes.Black, _
               20, Height - LegendHeight + FontLegend.Height * i + 1)
        Next

        Return bm

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create Pie Chart as an IO stream
    ''' </summary>
    ''' <param name="data">Pie data</param>
    ''' <param name="labels">Pie labels</param>
    ''' <param name="Title">Chart Title</param>
    ''' <param name="Width">Chart Width</param>
    ''' <returns>IO stream containg chart bitmap</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Function CreatePieChartStream(ByVal data As Array, ByVal labels As Array, _
ByVal Title As String, ByVal Width As Integer) As IO.Stream

        Dim bm As New Bitmap(Width, Width)
        bm = CreatePieChart(data, labels, Title, Width)

        Dim st As New IO.MemoryStream()

        bm.Save(st, Imaging.ImageFormat.Bmp)
        st.Position = 0
        Return st

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create Pie chart save bitmap to a file
    ''' </summary>
    ''' <param name="data">Pie data</param>
    ''' <param name="labels">Pie labels</param>
    ''' <param name="Title">Chart Title</param>
    ''' <param name="Width">Chart Width</param>
    ''' <param name="StrFileName">Filename for chart</param>
    ''' <returns>TRUE if succesful</returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Function CreatePieChartFile(ByVal data As Array, ByVal labels As Array, _
ByVal Title As String, ByVal Width As Integer, ByVal StrFileName As String) As Boolean

        Dim bm As New Bitmap(Width, Width)
        bm = CreatePieChart(data, labels, Title, Width)

        Try
            bm.Save(StrFileName, Imaging.ImageFormat.Jpeg)
        Catch ex As Exception
            Return False
        End Try
        Return True

    End Function


#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Rotate text by a certain angle
    ''' </summary>
    ''' <param name="gr">Graphics object</param>
    ''' <param name="text">Text to rotate</param>
    ''' <param name="x">X position of text (rotation point)</param>
    ''' <param name="y">Y position of text (rotation point)</param>
    ''' <param name="angle">Angle to rotate (degrees)</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shared Sub RotateText(ByVal gr As Graphics, ByVal text As String, _
      ByVal x As Single, ByVal y As Single, ByVal angle As Single)
        Dim graphics_path As New  _
     Drawing2D.GraphicsPath(Drawing.Drawing2D.FillMode.Winding)
        graphics_path.AddString(text, _
            New FontFamily("Verdana"), _
            FontStyle.Bold, 10, _
            New Point(x, y), _
            StringFormat.GenericDefault)

        ' Make a rotation matrix representing 
        ' rotation around the point (150, 150).
        Dim rotation_matrix As New Drawing2D.Matrix()
        rotation_matrix.RotateAt(angle, New PointF(x, y))

        ' Transform the GraphicsPath.
        graphics_path.Transform(rotation_matrix)

        ' Draw the text.
        With gr
            .FillPath(Brushes.Black, graphics_path)
        End With

    End Sub

#End Region

#Region "Properties"

    Private myErrors As String = ""
    Public Property Errors() As String
        Get
            Return myErrors
        End Get
        Set(ByVal value As String)
            myErrors = value
        End Set
    End Property



    Public Property FileName() As String
        Get
            Return myFileName
        End Get
        Set(ByVal value As String)
            myFileName = value
        End Set
    End Property


    Public Property ChartTitle() As Title
        Get
            Return myTitle
        End Get
        Set(ByVal value As Title)
            myTitle = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of Chart
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ChartType() As eChartType
        Get
            Return myChartType
        End Get
        Set(ByVal Value As eChartType)
            myChartType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show series values
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowValues() As Boolean
        Get
            Return blnShowValues
        End Get
        Set(ByVal Value As Boolean)
            blnShowValues = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Font for Labels
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LabelFont() As Font
        Get
            Return Me.FontLabels
        End Get
        Set(ByVal Value As Font)
            Me.FontLabels = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gap between histogram blocks
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Gap() As Integer
        Get
            Return HistGap
        End Get
        Set(ByVal Value As Integer)
            HistGap = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Size of the left hand margin
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LeftMargin() As Single
        Get
            Return sngLeftMargin
        End Get
        Set(ByVal Value As Single)
            sngLeftMargin = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Size of the bottom margin
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property BottomMargin() As Single
        Get
            Return sngBotMargin
        End Get
        Set(ByVal Value As Single)
            sngBotMargin = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Axis in the X direction
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property XAxis() As Axis
        Get
            Return myXAxis
        End Get
        Set(ByVal Value As Axis)
            myXAxis = Value
            If myXAxis Is Nothing Then Exit Property

        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Axis in the Y direction
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LeftYAxis() As Axis
        Get
            Return Me.myLeftAxis
        End Get
        Set(ByVal Value As Axis)
            myLeftAxis = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Series of data on chart (can be of mixed types)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Serieses() As SeriesCollection
        Get
            Return mySeries
        End Get
        Set(ByVal Value As SeriesCollection)
            mySeries = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Bounding Rectangle for chart
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Property ChartRectanle() As Rectangle
        Get
            Return myRect
        End Get
        Set(ByVal Value As Rectangle)
            myRect = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Area of chart (area within axes)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Protected Friend Property ChartArea() As Rectangle
        Get
            Return Me.ChartRect
        End Get
        Set(ByVal Value As Rectangle)
            ChartRect = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Chart Legend
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Legend() As Legend
        Get
            Return myLegend
        End Get
        Set(ByVal Value As Legend)
            myLegend = Value
        End Set
    End Property

    Public Property Title() As Title
        Get
            Return myTitle
        End Get
        Set(ByVal Value As Title)
            myTitle = Value
        End Set
    End Property

#End Region
#End Region

End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Axis
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Chart axes manipulation routines
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class Axis

#Region "Declarations"
    Private Min As Object
    Private Max As Object
    Private myStep As Object
    Private Ticks As Integer = 5
    Private FontLabels As New Font("verdana", 8)
    Private AxisPen As New Pen(Color.Black, 2)
    Private myFormat As String
    Private blnLeft As Boolean = True
    Private myChart As Charts
    Private Scale As Single = 0

#End Region

#Region "Public Instance"

#Region "Functions"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Determine height of a block using axis scaling 
    ''' </summary>
    ''' <param name="YPos">Y position to calculate</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Function BlockHeight(ByVal YPos As Object) As Single
        If Me.MyType = AxisType.XAxis Or MyType = AxisType.XAxisDate Then
            If TypeOf YPos Is Date Then
                If Scale = 0 Then
                    Scale = myChart.ChartArea.Width / (Convert.ToDateTime(Max).Subtract(Convert.ToDateTime(Min)).TotalMinutes)
                End If
                Dim x1 As Double = Convert.ToDateTime(YPos).Subtract(Convert.ToDateTime(Min)).TotalMinutes
                Dim x2 As Double = (Convert.ToDateTime(Max).Subtract(Convert.ToDateTime(Min)).TotalMinutes)
                Return x1 * Scale + myChart.ChartArea.Left
            Else
                If Scale = 0 Then
                    Scale = myChart.ChartArea.Width / (Max - Min)
                End If
                Dim x1 As Double = YPos - Min
                Dim x2 As Double = Max - Min
                Return x1 * Scale + myChart.ChartArea.Left
            End If
        Else
            If TypeOf YPos Is Single Or TypeOf YPos Is Integer Then
                If Scale = 0 Then
                    If Scale = 0 Then
                        Scale = myChart.ChartArea.Height / (Max - Min)
                    End If

                End If
                Return myChart.ChartArea.Height / ((Max - Min) / YPos) _
                         + myChart.ChartArea.Top
            End If

        End If

    End Function

#End Region

#Region "Procedures"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw axis using supplied graphics object
    ''' </summary>
    ''' <param name="gr"></param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawAxis(ByVal gr As Graphics)

        With myChart.ChartArea
            If MyType = AxisType.LeftY Then
                If blnLeft Then
                    gr.DrawLine(AxisPen, .Left, .Top, .Left, .Bottom)
                    Dim y As Single = Convert.ToSingle(Max)
                    Dim y2 As Single = .Top
                    Dim st As String
                    Dim i As Integer

                    If Not Me.IntervalStep Is Nothing Then
                        Ticks = (Me.Maximum - Me.Minimum) / Me.IntervalStep
                    End If
                    For i = 0 To Ticks - 1
                        st = y.ToString("0.00")
                        Dim s As SizeF = gr.MeasureString(st, FontLabels)
                        gr.DrawString(st, FontLabels, Brushes.Black, .Left - s.Width - 10, y2)
                        gr.DrawLine(AxisPen, .Left - 10, y2, .Left, y2)
                        If y = 0 Then
                            gr.DrawLine(AxisPen, .Left, y2, .Right, y2)
                        End If
                        y -= (Convert.ToSingle(Max - Min) / Ticks)

                        y2 += (.Height / Ticks)

                    Next
                End If
            ElseIf MyType = AxisType.XAxis Or MyType = AxisType.XAxisDate Then
                If Me.myChart.LeftYAxis.Minimum < 0 And Me.myChart.LeftYAxis.Maximum > 0 Then
                Else
                    gr.DrawLine(AxisPen, .Left, .Bottom, .Right, .Bottom)

                End If
                If TypeOf (Minimum) Is Date Then
                    Dim i As Integer = 0
                    Dim st As String
                    Dim y2 As Single = .Left

                    Dim StDate As Date = Convert.ToDateTime(Minimum)

                    While StDate <= Convert.ToDateTime(Maximum)
                        st = StDate.ToString(myFormat)
                        Dim s As SizeF = gr.MeasureString(st, FontLabels)
                        y2 = Me.BlockHeight(StDate)
                        Charts.RotateText(gr, st, y2, .Bottom + s.width, 270)
                        'gr.DrawString(st, FontLabels, Brushes.Black, y2, myChartRect.Bottom + 8)
                        gr.DrawLine(AxisPen, y2, .Bottom, y2, .Bottom + 5)
                        'y2 += (myChartRect.Width / Ticks)
                        StDate = StDate.AddMinutes(Me.IntervalStep)
                    End While

                Else
                    Dim y As Single = Convert.ToSingle(Minimum)
                    Dim y2 As Single = .Left
                    Dim st As String
                    Dim i As Integer
                    If Not Me.IntervalStep Is Nothing Then Ticks = (Me.Maximum - Me.Minimum) / Me.IntervalStep
                    For i = 0 To Ticks - 1
                        If y < 100 Then
                            st = y.ToString("0.00")
                        Else
                            st = y.ToString("0.")
                        End If
                        Dim s As SizeF = gr.MeasureString(st, FontLabels)
                        gr.DrawString(st, FontLabels, Brushes.Black, y2, .Bottom + 8)
                        gr.DrawLine(AxisPen, y2, .Bottom, y2, .Bottom + 5)
                        y += (Convert.ToSingle(Max - Min) / Ticks)
                        y2 += (.Width / Ticks)

                    Next


                End If
            ElseIf MyType = AxisType.XAxisSpeedo Then
                Dim i As Integer = 0
                Dim myRect As Rectangle = myChart.ChartArea
                myRect.Inflate(10, 10)
                myRect.X = myRect.X - 5
                myRect.Y = myRect.Y - 5


                While i < Me.Max
                    Dim Angle As Single
                    Dim astep As Single = Me.Max / Ticks

                    Angle = ((i * astep / Me.Max * 225) + 135) Mod 360
                    'Angle = (I / Me.Maximum * 360)
                    Dim sb As SolidBrush = New SolidBrush(Color.Black)
                    gr.FillPie(sb, myRect, Angle, 0.5)
                    i += Me.IntervalStep
                End While
            ElseIf MyType = AxisType.Slider Then
                gr.DrawLine(AxisPen, .Left, .Top, .Left, .Bottom)
            End If
        End With

    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new axis
    ''' </summary>
    ''' <param name="Axis">Axis type to create</param>
    ''' <param name="Chart">Chart to create axis for</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Axis As AxisType, ByVal Chart As Charts)
        MyType = Axis
        myChart = Chart

    End Sub

#End Region

#Region "Properties"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The format used for teh text on this axis
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Format() As String
        Get
            Return myFormat
        End Get
        Set(ByVal Value As String)
            myFormat = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The type of axis using an enumeration
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property AType() As AxisType
        Get
            Return MyType
        End Get
        Set(ByVal Value As AxisType)
            MyType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The minimum for the axis.
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' Declared as being object to allow for Dates etc.
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Minimum() As Object
        Get
            Return Min
        End Get
        Set(ByVal Value As Object)
            Min = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Maximum for the axis
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Maximum() As Object
        Get
            Return Max
        End Get
        Set(ByVal Value As Object)
            Max = Value
            If TypeOf Max Is Date Then
                myFormat = "HH:mm"

            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The step at which major ticks will occure
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property IntervalStep() As Object
        Get
            Return myStep
        End Get
        Set(ByVal Value As Object)
            myStep = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is the axis a left axis or a right axis, only appropriate for Y axes
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [28/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LeftAxis() As Boolean
        Get
            Return Me.blnLeft
        End Get
        Set(ByVal Value As Boolean)
            blnLeft = Value
        End Set
    End Property

#End Region

#End Region


    '''<summary>Axes Types</summary>
    Public Enum AxisType
        '''<summary>Left hand Y Axis</summary>
        LeftY
        '''<summary>Right hand Y Axis</summary>
        RightY
        '''<summary>Xaxis</summary>
        XAxis
        '''<summary>Xaxis for dates</summary>
        XAxisDate
        '''<summary>Speedometer X Axis</summary>
        XAxisSpeedo
        '''<summary>Slider</summary>
        Slider
    End Enum

    Private MyType As AxisType

End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Series
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A series of data points used on a chart
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class Series

#Region "Declarations"
    Dim myYValues As ArrayList
    Dim myXValues As ArrayList
    Dim myLabels As ArrayList
    Dim myColours As ArrayList
    Dim mySerieType As SeriesType
    Dim blnLeft As Boolean = True
    Dim strFormat As String = ""

    Dim Max As Single
    Dim Min As Single
    Dim HistGap As Single = 1

    Dim blnShowValues As Boolean = True
    Dim blnShowPoints As Boolean = True
    Dim blnBorder As Boolean = True
    Dim myColour As Color = Color.Blue
    Dim intBorderWidth As Integer = 1
    Dim clBorderColour As Color = Color.Black
    Dim xMin As Object
    Dim xMax As Object
    Dim xStep As Object
    Dim myChart As Charts
    Dim strSeriesName As String = ""
    Dim myDashStyle As Drawing2D.DashStyle = Drawing.Drawing2D.DashStyle.Solid

#End Region

#Region "Public Instance"

#Region "Functions"

    Private Function BlockHeight(ByVal YPos As Single) As Single
        If Max = 0 Then
            Return 0
        Else
            Return myChart.ChartArea.Height * ((YPos) / (Max - Minimum)) + myChart.ChartArea.Top

        End If

    End Function

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw Series on chart
    ''' </summary>
    ''' <param name="gr">Graphics object passed in</param>
    ''' <param name="FontLabels">Font for labels</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub DrawSeries(ByVal gr As Graphics, ByVal FontLabels As Font)
        Dim i As Integer
        Dim BlockHeight As Single
        Dim x1 As Single = myChart.ChartArea.Left
        Dim St As String
        Dim bWidth As Single = (myChart.ChartArea.Width) / (Me.myYValues.Count) - HistGap * 2
        Dim x2 As Single = x1



        If Me.sType = SeriesType.VertHistogram Then
            Dim myBrush As New SolidBrush(myColour)
            Dim myPen As New Pen(Me.clBorderColour, Me.intBorderWidth)
            For i = 0 To Me.YValues.Count - 1
                'Establish the height of the block
                BlockHeight = Me.BlockHeight(Convert.ToSingle(Me.myYValues.Item(i)))
                'Fill a rectangle, remember that the y values decrease from the top of the screen/paper
                x2 = myChart.XAxis.BlockHeight(Xvalues.Item(i))
                bWidth = x2 - x1 - HistGap - HistGap
                'Debug.WriteLine(x2 - x1)
                gr.FillRectangle(myBrush, New Rectangle(x1 + HistGap, _
                 myChart.ChartArea.Top + (myChart.ChartArea.Height - BlockHeight), bWidth, BlockHeight))
                If Me.blnBorder Then
                    gr.DrawRectangle(myPen, New Rectangle(x1 + HistGap, _
                 myChart.ChartArea.Top + (myChart.ChartArea.Height - BlockHeight), bWidth, BlockHeight))
                End If

                If blnShowValues Then       '   if we are going to label the block 
                    If strFormat.Length > 0 Then
                        If TypeOf myYValues(i) Is Single Or TypeOf myYValues(i) Is Double Or TypeOf myYValues(i) Is Integer Then
                            St = CDbl(myYValues(i)).ToString(strFormat)
                        Else
                            St = myYValues(i)
                        End If
                    Else
                        St = Me.myYValues(i)    '   get value 
                    End If
                    Dim s As SizeF = gr.MeasureString(St, FontLabels)   ' Find size of drawn string 
                    gr.DrawString(St, FontLabels, Brushes.Black, x1 + bWidth / 2, _
                    (myChart.ChartArea.Top) + (myChart.ChartArea.Height - BlockHeight) - s.Height) 'Draw string in position
                End If
                x1 = x2      'Advance to next block position 
            Next
        ElseIf Me.sType = SeriesType.Slider Then
            Dim myBrush As New SolidBrush(myColour)
            Dim BackBrush As New SolidBrush(Color.AliceBlue)
            Dim outerwidth As Single = myChart.LeftYAxis.BlockHeight(6) - myChart.ChartArea.Top
            Dim innerWidth As Single = myChart.LeftYAxis.BlockHeight(4) - myChart.ChartArea.Top

            Dim x3 As Single = myChart.XAxis.BlockHeight(Convert.ToSingle(myChart.XAxis.Maximum))
            For i = 0 To Me.YValues.Count - 1
                If Not myColours Is Nothing Then
                    myBrush = New SolidBrush(myColours(i))
                End If
                x2 = myChart.XAxis.BlockHeight(Convert.ToSingle(YValues.Item(i)))
                BlockHeight = myChart.LeftYAxis.BlockHeight(Convert.ToSingle(Me.myXValues.Item(i)))


                gr.FillRectangle(BackBrush, New Rectangle(x1, BlockHeight - outerwidth / 2 _
                  , x3 - x1, outerwidth))
                gr.FillRectangle(myBrush, New Rectangle(x1, BlockHeight - innerWidth / 2 _
                  , x2 - x1, innerWidth))
                If blnShowValues Then       '   if we are going to label the block 
                    St = Me.myYValues(i)    '   get value 
                    If Not Me.Labels Is Nothing Then
                        St = St & " " & Me.Labels(i)
                    End If
                    'outerwidth = 15
                    Dim s As SizeF = gr.MeasureString(St, FontLabels)   ' Find size of drawn string 
                    gr.DrawString(St, FontLabels, Brushes.Black, x2 + 20, _
                     BlockHeight - outerwidth) 'Draw string in position
                End If
            Next
        ElseIf Me.sType = SeriesType.XYGraph Then     ' Xy graph need to add capability of drawing histograms
            Dim y2 As Single = myChart.ChartArea.Bottom
            'Dim x2 As Single = ChartRect.Left

            Dim myPen As New Pen(myColour, Me.BorderWidth)
            myPen.LineJoin = Drawing.Drawing2D.LineJoin.Round
            'mypen.Brush = New SolidBrush(myColour)
            myPen.DashStyle = Me.DashStyle

            For i = 0 To Me.YValues.Count - 1
                BlockHeight = (Me.BlockHeight((Maximum) - Convert.ToSingle(Me.myYValues.Item(i))))
                Debug.WriteLine(Me.myYValues.Item(i) & " -  " & BlockHeight & " - " & Me.myChart.LeftYAxis.BlockHeight(Convert.ToSingle(Me.myYValues.Item(i))))

                x2 = x1
                x1 = myChart.XAxis.BlockHeight(Xvalues.Item(i))
                If Me.blnShowPoints Then
                    gr.DrawArc(myPen, x1, CInt(BlockHeight), 4, 4, 1, 360)  'Draw a cicle to indicate the position of the data point 
                End If
                '                                                       'Should allow other marker styles
                gr.DrawLine(myPen, x2, y2, x1, BlockHeight)             'Draw an interconnecting line 
                '                                                       'Should allow this to be turned off or differing styles 
                y2 = BlockHeight
                If blnShowValues Then                                   'If we are going to label the point do it here 
                    If strFormat.Length > 0 Then
                        If TypeOf myYValues(i) Is Single Or TypeOf myYValues(i) Is Double Or TypeOf myYValues(i) Is Integer Then
                            St = CDbl(myYValues(i)).ToString(strFormat)
                        Else
                            St = myYValues(i)
                        End If
                    Else
                        St = Me.myYValues(i)    '   get value 
                    End If
                    Dim s As SizeF = gr.MeasureString(St, FontLabels)
                    gr.DrawString(St, FontLabels, Brushes.Black, x2 + bWidth / 2, _
                    y2)
                End If
            Next
        ElseIf Me.sType = SeriesType.Pie Then
            Dim nRows As Integer = Me.myYValues.Count
            Dim Total As Single = 0

            For i = 0 To Me.myYValues.Count - 1
                Total += Me.myYValues(i)
            Next

            Dim Colours As New ArrayList()
            Dim rnd As New Random()

            For i = 0 To nRows - 1
                Colours.Add(New SolidBrush(Color.FromArgb(rnd.Next(255), rnd.Next(255), rnd.Next(255))))
            Next

            Dim Angle As Single = 0
            Dim oldAngle As Single = 0
            Dim sb As SolidBrush

            For i = 0 To nRows - 1
                Angle = (Me.myYValues(i) / Total * 360)
                sb = CType(Colours(i), SolidBrush)
                Debug.WriteLine(sb.Color.Name & ", " & oldAngle & ", " & Angle)
                gr.FillPie(sb, myChart.ChartArea, oldAngle, Angle)
                oldAngle += Angle
            Next

        ElseIf Me.sType = SeriesType.speedo Then

            Dim nRows As Integer = Me.myYValues.Count
            Dim Total As Single = 0

            For i = 0 To Me.myYValues.Count - 1
                Total += Me.myYValues(i)
            Next

            Dim Colours As New ArrayList()
            Dim rnd As New Random()

            For i = 0 To nRows - 1
                Colours.Add(Me.myXValues(i))
            Next

            Dim Angle As Single = 135
            Dim oldAngle As Single = 135
            Dim startangle As Single = 135
            Dim sb As SolidBrush

            Angle = (1400 / Me.Max * 405) Mod 360
            sb = New SolidBrush(Color.PeachPuff)
            Debug.WriteLine(sb.Color.Name & ", " & oldAngle & ", " & Angle)
            gr.FillPie(sb, myChart.ChartArea, startangle, 225)

            Dim pn As Pen
            Dim Radius As Single = (myChart.ChartArea.Width - 40) / 2
            Dim x0 As Single = Radius + myChart.ChartArea.Left
            Dim y0 As Single = Radius + myChart.ChartArea.Top
            For i = 0 To nRows - 1
                Angle = ((Me.myYValues(i) / Me.Max * 225) + 135) Mod 360
                oldAngle = Angle - 2
                sb = New SolidBrush(Colours(i))
                pn = New Pen(sb)
                gr.FillPie(sb, myChart.ChartArea, Angle, 0.5)
                'Centre of circle 
                'Dim x As Single = System.Math.Sin(angle) * Radius
                'Dim y As Single = System.Math.Sqrt(Radius - x)
                'gr.DrawLine(pn, x, y, x0, y0)
                'Debug.WriteLine(sb.Color.Name & ", " & oldAngle & ", " & Angle & ", " & x & ", " & y)

            Next

        End If

    End Sub


#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create new series
    ''' </summary>
    ''' <param name="SerieType">Type of Series</param>
    ''' <param name="Chart">Containing Chart for Series</param>
    ''' <param name="SeriesName">Name of Series</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal SerieType As SeriesType, ByVal Chart As Charts, ByVal SeriesName As String)
        myXValues = New ArrayList()
        myYValues = New ArrayList()
        mySerieType = SerieType
        myChart = Chart
        strSeriesName = SeriesName
    End Sub

#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Format string to use on data labels
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Format() As String
        Get
            Return strFormat
        End Get
        Set(ByVal Value As String)
            strFormat = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of colours to use with data point
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Colours() As ArrayList
        Get
            Return Me.myColours
        End Get
        Set(ByVal Value As ArrayList)
            myColours = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of data labels
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Labels() As ArrayList
        Get
            Return Me.myLabels
        End Get
        Set(ByVal Value As ArrayList)
            myLabels = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Line colour
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Colour() As Color
        Get
            Return myColour
        End Get
        Set(ByVal Value As Color)
            myColour = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Colour of block border
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property BorderColour() As Color
        Get
            Return Me.clBorderColour
        End Get
        Set(ByVal Value As Color)
            clBorderColour = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Width of border
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property BorderWidth() As Integer
        Get
            Return Me.intBorderWidth
        End Get
        Set(ByVal Value As Integer)
            intBorderWidth = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is axis on the Left (normal) or right
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LeftAxis() As Boolean
        Get
            Return blnLeft
        End Get
        Set(ByVal Value As Boolean)
            blnLeft = Value
            If Not blnLeft Then
                If mySerieType = SeriesType.XYGraph And myYValues.Count > 0 Then
                    Max = myYValues(myYValues.Count - 1)
                    Min = myYValues(0)
                End If
            End If
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of X axis positions
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Xvalues() As ArrayList
        Get
            Return myXValues
        End Get
        Set(ByVal Value As ArrayList)
            myXValues = Value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Minimum in the Y direction
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Minimum() As Single
        Get
            Return Min
        End Get
        Set(ByVal Value As Single)
            Min = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Maximum in the Y direction
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Maximum() As Single
        Get
            Return Max
        End Get
        Set(ByVal Value As Single)
            Max = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Gap between blocks
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Gap() As Single
        Get
            Return HistGap
        End Get
        Set(ByVal Value As Single)
            HistGap = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' An array of points in the Y direction
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Overloads Property YValues() As ArrayList
        Get
            Return myYValues
        End Get
        Set(ByVal Value As ArrayList)
            myYValues = Value
            Dim i As Integer
            For i = 0 To myYValues.Count - 1
                If Max < Convert.ToSingle(myYValues(i)) Then _
                        Max = Convert.ToSingle(myYValues(i))
                If Min > Convert.ToSingle(myYValues(i)) Then _
                        Min = Convert.ToSingle(myYValues(i))
            Next
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Series description (for legend)
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property SeriesName() As String
        Get
            Return strSeriesName
        End Get
        Set(ByVal Value As String)
            strSeriesName = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Type of series
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property sType() As SeriesType
        Get
            Return mySerieType
        End Get
        Set(ByVal Value As SeriesType)
            mySerieType = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show Data Labels on Graph
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowLabels() As Boolean
        Get
            Return Me.blnShowValues
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowValues = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Show Points on graph
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property ShowPoints() As Boolean
        Get
            Return Me.blnShowPoints
        End Get
        Set(ByVal Value As Boolean)
            Me.blnShowPoints = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Is block bordered
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Border() As Boolean
        Get
            Return Me.blnBorder
        End Get
        Set(ByVal Value As Boolean)
            blnBorder = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Dash style for series line
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property DashStyle() As Drawing2D.DashStyle
        Get
            Return Me.myDashStyle
        End Get
        Set(ByVal Value As Drawing2D.DashStyle)
            myDashStyle = Value
        End Set
    End Property

#End Region
#End Region


#Region "Public enumerations"
    '
    '   Public enumeration of series types 
    '
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The various types of series
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum SeriesType
        ''' <summary>XY graph Real data in both axes</summary>
        XYGraph
        ''' <summary>Histogram bars running vertically</summary>
        VertHistogram
        ''' <summary>Histogram bars running horizontally</summary>
        HorizHistogram
        ''' <summary>Pie Chart</summary>
        Pie
        ''' <summary>Speedometer style chart</summary>
        speedo
        ''' <summary>Slider style chart</summary>
        Slider
    End Enum
#End Region


    '
    '   Public Properties of class 
    '


    'Subroutines and functions 



End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : SeriesCollection
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' A collection of Chart Data series
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class SeriesCollection
    Inherits ArrayList

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' create a new collection of Data Series
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New()
        MyBase.New()
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Return a Data Series by position
    ''' </summary>
    ''' <param name="index">Position of Series</param>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Shadows Property item(ByVal index As Integer) As Series
        Get
            Return CType(MyBase.Item(index), Series)
        End Get
        Set(ByVal Value As Series)
            MyBase.Item(index) = Value
        End Set
    End Property

End Class

''' -----------------------------------------------------------------------------
''' Project	 : MedscreenLib
''' Class	 : Legend
''' 
''' -----------------------------------------------------------------------------
''' <summary>
''' Deals with Chart Legends
''' </summary>
''' <remarks>
''' </remarks>
''' <revisionHistory>
''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
''' </revisionHistory>
''' -----------------------------------------------------------------------------
Public Class Legend

#Region "Public Enumerations"

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Enumeration of Legend Positions
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Enum LegendPosition
        ''' <summary>At Top of Chart</summary>
        Top
        ''' <summary>At Left Hand side of Chart</summary>
        Left
        ''' <summary>At Right Hand side of Chart</summary>
        Right
        ''' <summary>At Bottom of Chart</summary>
        Bottom
    End Enum
#End Region

#Region "Declarations"
    Private myPosition As LegendPosition = LegendPosition.Right

    Private myChart As Charts
    Private blnLegend As Boolean = True
    Private LegendRect As Rectangle
    Private intLegendWidth As Integer = 150

#End Region

#Region "Public Instance"

#Region "Functions"

#End Region

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new legend object
    ''' </summary>
    ''' <param name="Chart">Chart containing legend</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Chart As Charts)
        myChart = Chart
        Position = LegendPosition.Right
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw Chart Legend
    ''' </summary>
    ''' <param name="gr">Graphics object to draw with</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Draw(ByVal gr As Graphics)
        gr.FillRectangle(Brushes.White, LegendRect)
        gr.DrawRectangle(Pens.Black, LegendRect)

        Dim i As Integer
        Dim s As Series

        Select Case myPosition
            Case LegendPosition.Left, LegendPosition.Right
                For i = 0 To myChart.Serieses.Count - 1
                    s = myChart.Serieses(i)
                    If s.sType = Series.SeriesType.HorizHistogram Or s.sType = Series.SeriesType.VertHistogram Then
                        gr.FillRectangle(New SolidBrush(s.Colour), New Rectangle(LegendRect.Left + 2, LegendRect.Top + (i + 1) * 20, 5, 5))
                    Else
                        gr.DrawLine(New Pen(s.Colour, 3), LegendRect.Left + 2, LegendRect.Top + (i + 1) * 20, LegendRect.Left + 12, LegendRect.Top + (i + 1) * 20)
                    End If
                    gr.DrawString(s.SeriesName, myChart.LabelFont, Brushes.Black, New RectangleF(LegendRect.Left + 25, LegendRect.Top + (i + 1) * 20, LegendRect.Width - LegendRect.Left - 25, 30))
                Next
            Case LegendPosition.Bottom, LegendPosition.Top
                Dim x1 As Single = LegendRect.Left + 5
                For i = 0 To myChart.Serieses.Count - 1
                    s = myChart.Serieses(i)
                    If s.sType = Series.SeriesType.HorizHistogram Or s.sType = Series.SeriesType.VertHistogram Then
                        gr.FillRectangle(New SolidBrush(s.Colour), New Rectangle(x1 + 2, LegendRect.Top + 20, 5, 5))
                    Else
                        gr.DrawLine(New Pen(s.Colour, 3), x1 + 2, LegendRect.Top + 20, x1 + 12, LegendRect.Top + 20)
                        x1 += 25
                    End If
                    gr.DrawString(s.SeriesName, myChart.LabelFont, Brushes.Black, New RectangleF(x1, LegendRect.Top + 20, 100, 30))
                    x1 += 100
                Next

        End Select
    End Sub
#End Region

#Region "Properties"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Legend Position from enumeration
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Position() As LegendPosition
        Get
            Return myPosition
        End Get
        Set(ByVal Value As LegendPosition)
            myPosition = Value
            With myChart
                Select Case myPosition
                    Case LegendPosition.Right
                        LegendRect = New Rectangle(.ChartRectanle.Right - intLegendWidth - 5, .ChartArea.Top, intLegendWidth, .ChartArea.Height)
                    Case LegendPosition.Left
                        LegendRect = New Rectangle(myChart.ChartRectanle.Left + 2, .ChartArea.Top, intLegendWidth, .ChartArea.Height)
                    Case LegendPosition.Bottom
                        LegendRect = New Rectangle(myChart.ChartRectanle.Left + 10, .ChartArea.Bottom - 20, .ChartRectanle.Width - 40, 50)
                    Case LegendPosition.Top
                        LegendRect = New Rectangle(myChart.ChartRectanle.Left + 10, .ChartRectanle.Top + 5, .ChartRectanle.Width - 40, 50)



                End Select
            End With
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether the Legend is displayed or not
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property HasLegend() As Boolean
        Get
            Return blnLegend
        End Get
        Set(ByVal Value As Boolean)
            blnLegend = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The width of the Legend
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property LegendWidth() As Integer
        Get
            Return intLegendWidth
        End Get
        Set(ByVal Value As Integer)
            intLegendWidth = Value
        End Set
    End Property

#End Region
#End Region


End Class

Public Class Title

#Region "Declarations"
    Public Enum TitlePosition
        ''' <summary>At Top of Chart</summary>
        Top
    End Enum

    Private myFont As Font


#End Region

#Region "Declarations"
    Private myPosition As TitlePosition = TitlePosition.Top

Private myChart As Charts
    Private blnTitle As Boolean = False
    Private TitleRect As Rectangle
    Private intTitleWidth As Integer = 0

#End Region
#Region "Public Instance"

#Region "Procedures"
    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Create a new Title object
    ''' </summary>
    ''' <param name="Chart">Chart containing Title</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub New(ByVal Chart As Charts)
        myChart = Chart
        Position = TitlePosition.Top
    End Sub

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Draw Chart Title
    ''' </summary>
    ''' <param name="gr">Graphics object to draw with</param>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Sub Draw(ByVal gr As Graphics)
        gr.FillRectangle(Brushes.White, TitleRect)
        gr.DrawRectangle(Pens.Black, TitleRect)


        Select Case myPosition
            'Case TitlePosition.Left, TitlePosition.Right
            '    For i = 0 To myChart.Serieses.Count - 1
            '        s = myChart.Serieses(i)
            '        If s.sType = Series.SeriesType.HorizHistogram Or s.sType = Series.SeriesType.VertHistogram Then
            '            gr.FillRectangle(New SolidBrush(s.Colour), New Rectangle(TitleRect.Left + 2, TitleRect.Top + (i + 1) * 20, 5, 5))
            '        Else
            '            gr.DrawLine(New Pen(s.Colour, 3), TitleRect.Left + 2, TitleRect.Top + (i + 1) * 20, TitleRect.Left + 12, TitleRect.Top + (i + 1) * 20)
            '        End If
            '        gr.DrawString(s.SeriesName, myChart.LabelFont, Brushes.Black, New RectangleF(TitleRect.Left + 25, TitleRect.Top + (i + 1) * 20, TitleRect.Width - TitleRect.Left - 25, 30))
            '    Next
            Case TitlePosition.Top ', TitlePosition.Bottom
                Dim x1 As Single = TitleRect.Left + 5
                'For i = 0 To myChart.Serieses.Count - 1
                '    s = myChart.Serieses(i)
                'If s.sType = Series.SeriesType.HorizHistogram Or s.sType = Series.SeriesType.VertHistogram Then
                '    gr.FillRectangle(New SolidBrush(s.Colour), New Rectangle(x1 + 2, TitleRect.Top + 20, 5, 5))
                'Else
                '    gr.DrawLine(New Pen(s.Colour, 3), x1 + 2, TitleRect.Top + 20, x1 + 12, TitleRect.Top + 20)
                '    x1 += 25
                'End If

                ' Set format of string. 
                Dim drawFormat As New StringFormat
                drawFormat.Alignment = StringAlignment.Center

                If myFont Is Nothing Then myFont = New Font(myChart.LabelFont.FontFamily, 10, FontStyle.Bold, GraphicsUnit.Point)


                gr.DrawString(myTitleText, myFont, Brushes.Black, New RectangleF(x1, TitleRect.Top + 20, TitleWidth, 30), drawFormat)
                'x1 += 100
                'Next

        End Select
    End Sub
#End Region

#End Region
#Region "Properties"


    Public Property TitleFont() As Font
        Get
            Return myFont
        End Get
        Set(ByVal value As Font)
            myFont = value
        End Set
    End Property


    Private myTitleText As String
    Public Property TitleText() As String
        Get
            Return myTitleText
        End Get
        Set(ByVal value As String)
            myTitleText = value
        End Set
    End Property


    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Title Position from enumeration
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property Position() As TitlePosition
        Get
            Return myPosition
        End Get
        Set(ByVal Value As TitlePosition)
            myPosition = Value
            With myChart
                Select Case myPosition
                    'Case TitlePosition.Right
                    '    TitleRect = New Rectangle(.ChartRectanle.Right - intTitleWidth - 5, .ChartArea.Top, intTitleWidth, .ChartArea.Height)
                    'Case TitlePosition.Left
                    '    TitleRect = New Rectangle(myChart.ChartRectanle.Left + 2, .ChartArea.Top, intTitleWidth, .ChartArea.Height)
                    'Case TitlePosition.Bottom
                    '    TitleRect = New Rectangle(myChart.ChartRectanle.Left + 10, .ChartArea.Bottom + 10, .ChartRectanle.Width - 40, 50)
                    Case TitlePosition.Top
                        TitleRect = New Rectangle(myChart.ChartRectanle.Left + 10, .ChartRectanle.Top + 5, .ChartRectanle.Width - 40, 50)



                End Select
            End With
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' Indicates whether the Title is displayed or not
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property HasTitle() As Boolean
        Get
            Return blnTitle
        End Get
        Set(ByVal Value As Boolean)
            blnTitle = Value
        End Set
    End Property

    ''' -----------------------------------------------------------------------------
    ''' <summary>
    ''' The width of the Title
    ''' </summary>
    ''' <value></value>
    ''' <remarks>
    ''' </remarks>
    ''' <revisionHistory>
    ''' <revision><Author>[taylor]</Author><date> [29/09/2005]</date><Action></Action></revision>
    ''' </revisionHistory>
    ''' -----------------------------------------------------------------------------
    Public Property TitleWidth() As Integer
        Get
            If intTitleWidth = 0 Then intTitleWidth = myChart.ChartArea.Width

            Return intTitleWidth
        End Get
        Set(ByVal Value As Integer)
            intTitleWidth = Value
        End Set
    End Property

#End Region
End Class


#End Region