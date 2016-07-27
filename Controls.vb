Option Strict On
Option Explicit On 
Imports System.Collections
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data

Namespace cb

    Public Class CtlDTPicker
        Inherits DateTimePicker

        Public Sub New()
            MyBase.new()
        End Sub
        Public isInEditOrNavigateMode As Boolean = True

    End Class
    Public Class CtlDataGridComboBox
        Inherits ComboBox
        Public Sub New()
            MyBase.New()
        End Sub
        Public isInEditOrNavigateMode As Boolean = True
    End Class

    Public Class DataGridComboBoxColumnStyle
        Inherits DataGridColumnStyle
        '
        ' UI constants
        '
        Private xMargin As Integer = 2
        Private yMargin As Integer = 1
        Private WithEvents Combo As CtlDataGridComboBox
        Private _DisplayMember As String
        Private _ValueMember As String
        '>>>>>>>>>Two Lines Below Were Added>>>>>>>>>>>>>>>>>>>>>>
        Private _Source As CurrencyManager
        Private _RowNum As Integer
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        '
        ' Used to track editing state
        '
        Private OldVal As String = String.Empty
        Private InEdit As Boolean = False
        '
        ' Create a new column - DisplayMember, ValueMember
        ' Passed by ordinal
        '
        Public Sub New(ByRef DataSource As DataTable, _
        ByVal DisplayMember As Integer, _
        ByVal ValueMember As Integer)
            Combo = New CtlDataGridComboBox()
            _DisplayMember = DataSource.Columns.Item(index:=DisplayMember).ToString
            _ValueMember = DataSource.Columns.Item(index:=ValueMember).ToString
            With Combo
                .Visible = False
                .DataSource = DataSource
                .DisplayMember = _DisplayMember
                .ValueMember = _ValueMember
                .DropDownStyle = ComboBoxStyle.DropDownList
            End With
        End Sub
        '
        ' Create a new column - DisplayMember, ValueMember
        ' passed by string
        '
        Public Sub New(ByRef DataSource As DataTable, _
        ByVal DisplayMember As String, _
        ByVal ValueMember As String)
            Combo = New CtlDataGridComboBox()
            _DisplayMember = DisplayMember
            _ValueMember = ValueMember
            With Combo
                .Visible = False
                .DataSource = DataSource
                .DisplayMember = DisplayMember
                .ValueMember = ValueMember
                .DropDownStyle = ComboBoxStyle.DropDownList
            End With
        End Sub
        '------------------------------------------------------
        ' Methods overridden from DataGridColumnStyle
        '------------------------------------------------------
        '
        ' Abort Changes
        '
        Protected Overloads Overrides Sub Abort(ByVal RowNum As Integer)
            Debug.WriteLine("Abort()")
            RollBack()
            HideComboBox()
            EndEdit()
        End Sub
        '
        ' Commit Changes
        '
        Protected Overrides Function Commit(ByVal DataSource As CurrencyManager, _
        ByVal RowNum As Integer) As Boolean
            HideComboBox()
            If Not InEdit Then
                Return True
            End If
            Try
                Dim Value As Object = Combo.SelectedValue
                If NullText.Equals(Value) Then
                    Value = Convert.DBNull
                End If
                SetColumnValueAtRow(DataSource, RowNum, Value)
            Catch e As Exception
                RollBack()
                Return False
            End Try
            EndEdit()
            Return True
        End Function
        '
        ' Remove focus
        '
        Protected Overloads Overrides Sub ConcedeFocus()
            Combo.Visible = False
            InEdit = False
        End Sub
        '
        ' Edit Grid
        '
        Protected Overloads Overrides Sub Edit(ByVal Source As CurrencyManager, _
        ByVal Rownum As Integer, _
        ByVal Bounds As Rectangle, _
        ByVal [ReadOnly] As Boolean, _
        ByVal InstantText As String, _
        ByVal CellIsVisible As Boolean)
            '>>>>>>>>>Two Lines Below Were Added>>>>>>>>>>>>>>>>>>>>>>
            _Source = Source
            _RowNum = Rownum
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Combo.Text = String.Empty
            Dim OriginalBounds As Rectangle = Bounds
            OldVal = Combo.Text
            If CellIsVisible Then
                Bounds.Offset(xMargin, yMargin)
                Bounds.Width -= xMargin * 2
                Bounds.Height -= yMargin
                Combo.Bounds = Bounds
                Combo.Visible = True
            Else
                Combo.Bounds = OriginalBounds
                Combo.Visible = False
            End If
            Dim TextObject As Object = GetColumnValueAtRow(Source, Rownum)
            If Not (TextObject Is System.DBNull.Value) Then Combo.SelectedValue = TextObject.ToString()
            If Not InstantText Is Nothing Then
                Combo.SelectedValue = InstantText
            End If
            Combo.RightToLeft = Me.DataGridTableStyle.DataGrid.RightToLeft
            ' Combo.Focus()
            If InstantText Is Nothing Then
                Combo.SelectAll()
            Else
                Dim [End] As Integer = Combo.Text.Length
                Combo.Select([End], 0)
            End If
            If Combo.Visible Then
                DataGridTableStyle.DataGrid.Invalidate(OriginalBounds)
            End If
            InEdit = True
        End Sub
        Protected Overloads Overrides Function GetMinimumHeight() As Integer
            '
            ' Set the minimum height to the height of the combobox
            '
            Return Combo.PreferredHeight + yMargin
        End Function
        Protected Overloads Overrides Function GetPreferredHeight(ByVal g As Graphics, _
        ByVal Value As Object) As Integer
            Debug.WriteLine("GetPreferredHeight()")
            Dim NewLineIndex As Integer = 0
            Dim NewLines As Integer = 0
            Dim ValueString As String = Me.GetText(Value)
            Do
                While NewLineIndex <> -1
                    NewLineIndex = ValueString.IndexOf("r\n", NewLineIndex + 1)
                    NewLines += 1
                End While
            Loop
            Return FontHeight * NewLines + yMargin
        End Function
        Protected Overloads Overrides Function GetPreferredSize(ByVal g As Graphics, _
        ByVal Value As Object) As Size
            Dim Extents As Size = Size.Ceiling(g.MeasureString(GetText(Value), _
            Me.DataGridTableStyle.DataGrid.Font))
            Extents.Width += xMargin * 2 + DataGridTableGridLineWidth
            Extents.Height += yMargin
            Return Extents
        End Function

        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer)
            Paint(g, Bounds, Source, RowNum, False)
        End Sub
        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer, _
        ByVal AlignToRight As Boolean)
            '>>>>Changed to Show Display Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim objText As Object = GetColumnValueAtRow(Source, RowNum)
            Dim Text As String = LookupDisplayValue(objText)
            PaintText(g, Bounds, Text, AlignToRight)
        End Sub
        Protected Overloads Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer, _
        ByVal BackBrush As Brush, _
        ByVal ForeBrush As Brush, _
        ByVal AlignToRight As Boolean)
            '>>>>Changed to Show Display Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim objText As Object = GetColumnValueAtRow(Source, RowNum)
            Dim Text As String = LookupDisplayValue(objText)
            PaintText(g, Bounds, Text, AlignToRight)
        End Sub
        Protected Overloads Overrides Sub SetDataGridInColumn(ByVal Value As DataGrid)
            MyBase.SetDataGridInColumn(Value)
            If Not (Combo.Parent Is Value) Then
                If Not (Combo.Parent Is Nothing) Then
                    Combo.Parent.Controls.Remove(Combo)
                End If
            End If
            If Not (Value Is Nothing) Then Value.Controls.Add(Combo)
        End Sub
        Protected Overloads Overrides Sub UpdateUI(ByVal Source As CurrencyManager, _
    ByVal RowNum As Integer, ByVal InstantText As String)
            Combo.Text = GetText(GetColumnValueAtRow(Source, RowNum))
            If Not (InstantText Is Nothing) Then
                Combo.Text = InstantText
            End If
        End Sub
        '----------------------------------------------------------------------
        ' Helper Methods
        '----------------------------------------------------------------------
        Private ReadOnly Property DataGridTableGridLineWidth() As Integer
            Get
                If Me.DataGridTableStyle.GridLineStyle = DataGridLineStyle.Solid Then
                    Return 1
                Else
                    Return 0
                End If
            End Get
        End Property
        Private Sub EndEdit()
            InEdit = False
            Invalidate()
        End Sub
        Private Function GetText(ByVal Value As Object) As String
            If Value Is System.DBNull.Value Then Return NullText
            If Not Value Is Nothing Then
                Return Value.ToString
            Else
                Return String.Empty
            End If
        End Function
        Private Sub HideComboBox()
            If Combo.Focused Then
                Me.DataGridTableStyle.DataGrid.Focus()
            End If
            Combo.Visible = False
        End Sub
        Private Sub RollBack()
            Combo.Text = OldVal
        End Sub
        Private Sub PaintText(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Text As String, _
        ByVal AlignToRight As Boolean)
            Dim BackBrush As Brush = New SolidBrush(Me.DataGridTableStyle.BackColor)
            Dim ForeBrush As Brush = New SolidBrush(Me.DataGridTableStyle.ForeColor)
            PaintText(g, Bounds, Text, BackBrush, ForeBrush, AlignToRight)
        End Sub
        Private Sub PaintText(ByVal g As Graphics, _
        ByVal TextBounds As Rectangle, _
        ByVal Text As String, _
        ByVal BackBrush As Brush, _
        ByVal ForeBrush As Brush, _
        ByVal AlignToRight As Boolean)
            Dim Rect As Rectangle = TextBounds
            Dim RectF As RectangleF = RectangleF.op_Implicit(Rect) ' Convert to RectangleF
            Dim Format As StringFormat = New StringFormat()
            If AlignToRight Then
                Format.FormatFlags = StringFormatFlags.DirectionRightToLeft
            End If
            Select Case Me.Alignment
                Case Is = HorizontalAlignment.Left
                    Format.Alignment = StringAlignment.Near
                Case Is = HorizontalAlignment.Right
                    Format.Alignment = StringAlignment.Far
                Case Is = HorizontalAlignment.Center
                    Format.Alignment = StringAlignment.Center
            End Select
            Format.FormatFlags = Format.FormatFlags Or StringFormatFlags.NoWrap
            g.FillRectangle(Brush:=BackBrush, Rect:=Rect)
            Rect.Offset(0, yMargin)
            Rect.Height -= yMargin
            g.DrawString(Text, Me.DataGridTableStyle.DataGrid.Font, ForeBrush, RectF, Format)
            Format.Dispose()
        End Sub
        Private Sub ComboChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Combo.SelectionChangeCommitted
            ColumnStartedEditing(Combo)
            SetColumnValueAtRow(_Source, _RowNum, Combo.SelectedValue)
        End Sub
        Private Function LookupDisplayValue(ByVal LookupObject As Object) As String
            If LookupObject Is System.DBNull.Value Then Return NullText
            Dim LookupText As String = LookupObject.ToString
            Dim I As Integer
            Dim IMax As Integer = Combo.Items.Count - 1
            Dim DT As DataTable = CType(Combo.DataSource, DataTable)
            For I = 0 To IMax
                If DT.Rows(I)(_ValueMember).ToString = LookupText Then
                    Return DT.Rows(I)(_DisplayMember).ToString
                End If
            Next
            Return String.Empty
        End Function
    End Class

    Public Class DTPickerComboBoxColumnStyle
        Inherits DataGridColumnStyle
        '
        ' UI constants
        '
        Private xMargin As Integer = 2
        Private yMargin As Integer = 1
        Private WithEvents DtPick As CtlDTPicker
        Private _DisplayMember As String
        Private _ValueMember As String
        '>>>>>>>>>Two Lines Below Were Added>>>>>>>>>>>>>>>>>>>>>>
        Private _Source As CurrencyManager
        Private _RowNum As Integer
        '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        '
        ' Used to track editing state
        '
        Private OldVal As Date = Date.Now
        Private InEdit As Boolean = False
        '
        ' Create a new column - DisplayMember, ValueMember
        ' Passed by ordinal
        '
        Public Sub New()
            DtPick = New CtlDTPicker()
            With DtPick
                .Visible = False
            End With
        End Sub
        '
        ' Create a new column - DisplayMember, ValueMember
        ' passed by string
        '

        '------------------------------------------------------
        ' Methods overridden from DataGridColumnStyle
        '------------------------------------------------------
        '
        ' Abort Changes
        '
        Protected Overloads Overrides Sub Abort(ByVal RowNum As Integer)
            Debug.WriteLine("Abort()")
            RollBack()
            HidePicker()
            EndEdit()
        End Sub
        '
        ' Commit Changes
        '
        Protected Overrides Function Commit(ByVal DataSource As CurrencyManager, _
        ByVal RowNum As Integer) As Boolean
            HidePicker()
            If Not InEdit Then
                Return True
            End If
            Try
                Dim Value As Object = DtPick.Value
                If NullText.Equals(Value) Then
                    Value = Convert.DBNull
                End If
                SetColumnValueAtRow(DataSource, RowNum, Value)
            Catch e As Exception
                RollBack()
                Return False
            End Try
            EndEdit()
            Return True
        End Function
        '
        ' Remove focus
        '
        Protected Overloads Overrides Sub ConcedeFocus()
            DtPick.Visible = False
            InEdit = False
        End Sub
        '
        ' Edit Grid
        '
        Protected Overloads Overrides Sub Edit(ByVal Source As CurrencyManager, _
        ByVal Rownum As Integer, _
        ByVal Bounds As Rectangle, _
        ByVal [ReadOnly] As Boolean, _
        ByVal InstantText As String, _
        ByVal CellIsVisible As Boolean)
            '>>>>>>>>>Two Lines Below Were Added>>>>>>>>>>>>>>>>>>>>>>
            _Source = Source
            _RowNum = Rownum
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            DtPick.Text = String.Empty
            Dim OriginalBounds As Rectangle = Bounds
            OldVal = DtPick.Value
            If CellIsVisible Then
                Bounds.Offset(xMargin, yMargin)
                Bounds.Width -= xMargin * 2
                Bounds.Height -= yMargin
                DtPick.Bounds = Bounds
                DtPick.Visible = True
            Else
                DtPick.Bounds = OriginalBounds
                DtPick.Visible = False
            End If
            Dim TextObject As Object = GetColumnValueAtRow(Source, Rownum)
            If Not (TextObject Is System.DBNull.Value) Then DtPick.Value = Date.Parse(TextObject.ToString())
            If Not InstantText Is Nothing Then
                DtPick.Value = Date.Parse(InstantText)
            End If
            DtPick.RightToLeft = Me.DataGridTableStyle.DataGrid.RightToLeft
            ' Combo.Focus()
            If InstantText Is Nothing Then
                DtPick.Select()
            Else
                Dim [End] As Integer = DtPick.Text.Length
                DtPick.Select()
            End If
            If DtPick.Visible Then
                DataGridTableStyle.DataGrid.Invalidate(OriginalBounds)
            End If
            InEdit = True
        End Sub
        Protected Overloads Overrides Function GetMinimumHeight() As Integer
            '
            ' Set the minimum height to the height of the combobox
            '
            Return DtPick.PreferredHeight + yMargin
        End Function
        Protected Overloads Overrides Function GetPreferredHeight(ByVal g As Graphics, _
        ByVal Value As Object) As Integer
            Debug.WriteLine("GetPreferredHeight()")
            Dim NewLineIndex As Integer = 0
            Dim NewLines As Integer = 0
            Dim ValueString As String = Me.GetText(Value)
            Do
                While NewLineIndex <> -1
                    NewLineIndex = ValueString.IndexOf("r\n", NewLineIndex + 1)
                    NewLines += 1
                End While
            Loop
            Return FontHeight * NewLines + yMargin
        End Function
        Protected Overloads Overrides Function GetPreferredSize(ByVal g As Graphics, _
        ByVal Value As Object) As Size
            Dim Extents As Size = Size.Ceiling(g.MeasureString(GetText(Value), _
            Me.DataGridTableStyle.DataGrid.Font))
            Extents.Width += xMargin * 2 + DataGridTableGridLineWidth
            Extents.Height += yMargin
            Return Extents
        End Function

        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer)
            Paint(g, Bounds, Source, RowNum, False)
        End Sub
        Protected Overloads Overrides Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer, _
        ByVal AlignToRight As Boolean)
            '>>>>Changed to Show Display Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim objText As Object = GetColumnValueAtRow(Source, RowNum)
            Dim Text As String = objText.ToString()
            PaintText(g, Bounds, Text, AlignToRight)
        End Sub
        Protected Overloads Sub Paint(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Source As CurrencyManager, _
        ByVal RowNum As Integer, _
        ByVal BackBrush As Brush, _
        ByVal ForeBrush As Brush, _
        ByVal AlignToRight As Boolean)
            '>>>>Changed to Show Display Value>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Dim objText As Object = GetColumnValueAtRow(Source, RowNum)
            Dim Text As String = objText.ToString
            PaintText(g, Bounds, Text, AlignToRight)
        End Sub
        Protected Overloads Overrides Sub SetDataGridInColumn(ByVal Value As DataGrid)
            MyBase.SetDataGridInColumn(Value)
            If Not (DtPick.Parent Is Value) Then
                If Not (DtPick.Parent Is Nothing) Then
                    DtPick.Parent.Controls.Remove(DtPick)
                End If
            End If
            If Not (Value Is Nothing) Then Value.Controls.Add(DtPick)
        End Sub
        Protected Overloads Overrides Sub UpdateUI(ByVal Source As CurrencyManager, _
    ByVal RowNum As Integer, ByVal InstantText As String)
            DtPick.Text = GetText(GetColumnValueAtRow(Source, RowNum))
            If Not (InstantText Is Nothing) Then
                DtPick.Text = InstantText
            End If
        End Sub
        '----------------------------------------------------------------------
        ' Helper Methods
        '----------------------------------------------------------------------
        Private ReadOnly Property DataGridTableGridLineWidth() As Integer
            Get
                If Me.DataGridTableStyle.GridLineStyle = DataGridLineStyle.Solid Then
                    Return 1
                Else
                    Return 0
                End If
            End Get
        End Property
        Private Sub EndEdit()
            InEdit = False
            Invalidate()
        End Sub
        Private Function GetText(ByVal Value As Object) As String
            If Value Is System.DBNull.Value Then Return NullText
            If Not Value Is Nothing Then
                Return Value.ToString
            Else
                Return String.Empty
            End If
        End Function
        Private Sub HidePicker()
            If DtPick.Focused Then
                Me.DataGridTableStyle.DataGrid.Focus()
            End If
            DtPick.Visible = False
        End Sub
        Private Sub RollBack()
            DtPick.Value = OldVal
        End Sub
        Private Sub PaintText(ByVal g As Graphics, _
        ByVal Bounds As Rectangle, _
        ByVal Text As String, _
        ByVal AlignToRight As Boolean)
            Dim BackBrush As Brush = New SolidBrush(Me.DataGridTableStyle.BackColor)
            Dim ForeBrush As Brush = New SolidBrush(Me.DataGridTableStyle.ForeColor)
            PaintText(g, Bounds, Text, BackBrush, ForeBrush, AlignToRight)
        End Sub
        Private Sub PaintText(ByVal g As Graphics, _
        ByVal TextBounds As Rectangle, _
        ByVal Text As String, _
        ByVal BackBrush As Brush, _
        ByVal ForeBrush As Brush, _
        ByVal AlignToRight As Boolean)
            Dim Rect As Rectangle = TextBounds
            Dim RectF As RectangleF = RectangleF.op_Implicit(Rect) ' Convert to RectangleF
            Dim Format As StringFormat = New StringFormat()
            If AlignToRight Then
                Format.FormatFlags = StringFormatFlags.DirectionRightToLeft
            End If
            Select Case Me.Alignment
                Case Is = HorizontalAlignment.Left
                    Format.Alignment = StringAlignment.Near
                Case Is = HorizontalAlignment.Right
                    Format.Alignment = StringAlignment.Far
                Case Is = HorizontalAlignment.Center
                    Format.Alignment = StringAlignment.Center
            End Select
            Format.FormatFlags = Format.FormatFlags Or StringFormatFlags.NoWrap
            g.FillRectangle(Brush:=BackBrush, Rect:=Rect)
            Rect.Offset(0, yMargin)
            Rect.Height -= yMargin
            g.DrawString(Text, Me.DataGridTableStyle.DataGrid.Font, ForeBrush, RectF, Format)
            Format.Dispose()
        End Sub
        Private Sub ComboChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DtPick.ValueChanged
            ColumnStartedEditing(DtPick)
            SetColumnValueAtRow(_Source, _RowNum, DtPick.Value)
        End Sub
    End Class
End Namespace
