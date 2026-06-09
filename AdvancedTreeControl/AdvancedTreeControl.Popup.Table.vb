
Imports System.Drawing.Drawing2D

Partial Public Class AdvancedTreeControl

    ' ========================================================================
    '  DIAGNOSTIC TEMPORAR (de șters după validarea Fazei 3)
    '  Apel: AdvancedTreeControl.TT_MeasureDump("<table>...</table>")
    '  Verifică "dimensiuni sănătoase" fără hover real.
    ' ========================================================================
    Friend Shared Sub TT_MeasureDump(xml As String)
        Dim model As TooltipTableModel = Nothing
        Dim err As String = Nothing
        If Not TooltipTableParser.TryParse(xml, model, err) Then
            MessageBox.Show("TryParse FAIL: " & err, "TT_MeasureDump")
            Return
        End If

        Dim p As New TooltipPopup()
        Dim forceHandle = p.Handle      ' MeasureTable folosește Me.CreateGraphics
        MessageBox.Show(p.DebugMeasure(model), "TT_MeasureDump - dimensiuni")
        p.Dispose()
    End Sub

    Partial Private Class TooltipPopup

        ' Plafon minim pentru o coloană auto strânsă (loc de "…").
        Private Const MIN_AUTO_COL_WIDTH As Integer = 24

        ' ── Stare mod tabel ─────────────────────────────────────────────────
        Private _isTableMode As Boolean = False
        Private _tableModel As TooltipTableModel = Nothing

        ' Rezultate MeasureTable (citite în PaintTable / poziționare Faza 4)
        Private _colWidths() As Integer = Array.Empty(Of Integer)()
        Private _rowHeight As Integer = 0
        Private _tableWidth As Integer = 0      ' fără PADDING_H
        Private _tableHeight As Integer = 0     ' fără PADDING_V (natural, nelimitat)

        ' ====================================================================
        '  MĂSURARE
        ' ====================================================================
        Private Sub MeasureTable()
            Try
                If _tableModel Is Nothing Then Return
                Dim cfg As TableConfig = _tableModel.Config
                Dim nCols As Integer = _tableModel.ColCount
                If nCols < 1 Then Return

                Using tableFont As New Font(cfg.FontName, cfg.FontSize, FontStyle.Regular, GraphicsUnit.Point)

                    ' rowHeight: fix (din config) sau auto (din font + padding vertical)
                    If cfg.RowHeight > 0 Then
                        _rowHeight = cfg.RowHeight
                    Else
                        _rowHeight = tableFont.Height + cfg.CellPaddingV * 2
                    End If

                    ReDim _colWidths(nCols - 1)
                    Dim isFixed(nCols - 1) As Boolean

                    ' 1. Lățimi FIXE din celulele de header cu Width > 0
                    If _tableModel.HeaderRow IsNot Nothing Then
                        Dim hr As TtRow = _tableModel.HeaderRow
                        Dim lim As Integer = Math.Min(nCols, hr.Cells.Count) - 1
                        For i As Integer = 0 To lim
                            Dim hc As TtCell = hr.Cells(i)
                            If hc.Width > 0 Then
                                _colWidths(i) = hc.Width      ' Width = lățimea TOTALĂ a coloanei (px deja scalați DPI)
                                isFixed(i) = True
                            End If
                        Next
                    End If

                    ' 2. Lățimi AUTO = max(conținut pe toată coloana) + 2*PadH
                    Using g As Graphics = Me.CreateGraphics()
                        For i As Integer = 0 To nCols - 1
                            If Not isFixed(i) Then
                                Dim maxContent As Integer = 0
                                maxContent = Math.Max(maxContent, MeasureCellContent(g, tableFont, RowCellAt(_tableModel.HeaderRow, i)))
                                For Each r As TtRow In _tableModel.Rows
                                    maxContent = Math.Max(maxContent, MeasureCellContent(g, tableFont, RowCellAt(r, i)))
                                Next
                                maxContent = Math.Max(maxContent, MeasureCellContent(g, tableFont, RowCellAt(_tableModel.FooterRow, i)))
                                _colWidths(i) = maxContent + cfg.CellPaddingH * 2
                            End If
                        Next
                    End Using

                    ' 3. Aplică MaxWidth (S2): strânge DOAR coloanele auto, proporțional, cu plafon minim
                    Dim natural As Integer = _colWidths.Sum()
                    If cfg.MaxWidth > 0 AndAlso natural > cfg.MaxWidth Then
                        ShrinkAutoColumns(_colWidths, isFixed, cfg.MaxWidth)
                    End If

                    ' 4. Dimensiuni finale
                    _tableWidth = _colWidths.Sum()

                    Dim nRows As Integer = 0
                    If _tableModel.HeaderRow IsNot Nothing Then nRows += 1
                    nRows += _tableModel.Rows.Count
                    If _tableModel.FooterRow IsNot Nothing Then nRows += 1
                    _tableHeight = nRows * _rowHeight
                End Using

            Catch ex As Exception
                TreeLogger.Err($"Error measuring tooltip table: {ex.Message}", "TooltipPopup.MeasureTable")
            End Try
        End Sub

        ' ====================================================================
        '  DESENARE
        ' ====================================================================
        Private Sub PaintTable(g As Graphics)
            Try
                If _tableModel Is Nothing OrElse _colWidths.Length = 0 Then Return
                Dim cfg As TableConfig = _tableModel.Config

                g.SmoothingMode = SmoothingMode.None   ' grilă 1px crisp (fundalul rotunjit s-a desenat deja cu AntiAlias)

                Using tableFont As New Font(cfg.FontName, cfg.FontSize, FontStyle.Regular, GraphicsUnit.Point)
                    Using gridPen As New Pen(cfg.GridColor, 1)

                        Dim x0 As Integer = PADDING_H
                        Dim y As Integer = PADDING_V

                        ' ── Header ──────────────────────────────────────────────
                        If _tableModel.HeaderRow IsNot Nothing Then
                            Using hb As New SolidBrush(cfg.HeaderBackColor)
                                g.FillRectangle(hb, x0, y, _tableWidth, _rowHeight)
                            End Using
                            PaintRow(g, tableFont, _tableModel.HeaderRow, x0, y, cfg, isHeader:=True, isFooter:=False)
                            y += _rowHeight
                            If cfg.GridVisible Then g.DrawLine(gridPen, x0, y, x0 + _tableWidth, y)
                        End If

                        ' ── Rânduri de date ─────────────────────────────────────
                        Dim hasFooter As Boolean = (_tableModel.FooterRow IsNot Nothing)
                        Dim lastBody As Integer = _tableModel.Rows.Count - 1
                        For idx As Integer = 0 To lastBody
                            Dim r As TtRow = _tableModel.Rows(idx)
                            If Not r.BackColor.IsEmpty Then
                                Using rb As New SolidBrush(r.BackColor)
                                    g.FillRectangle(rb, x0, y, _tableWidth, _rowHeight)
                                End Using
                            End If
                            PaintRow(g, tableFont, r, x0, y, cfg, isHeader:=False, isFooter:=False)
                            y += _rowHeight
                            ' suprimă linia sub ultimul rând de date dacă urmează footer (separatorul footer-ului o desenează)
                            If cfg.GridVisible AndAlso Not (idx = lastBody AndAlso hasFooter) Then
                                g.DrawLine(gridPen, x0, y, x0 + _tableWidth, y)
                            End If
                        Next

                        ' ── Footer ───────────────────────────────────────────────
                        If _tableModel.FooterRow IsNot Nothing Then
                            ' separator discret 1px deasupra footer-ului, în GridColor
                            If cfg.FooterSeparator Then g.DrawLine(gridPen, x0, y, x0 + _tableWidth, y)
                            If Not _tableModel.FooterRow.BackColor.IsEmpty Then
                                Using fb As New SolidBrush(_tableModel.FooterRow.BackColor)
                                    g.FillRectangle(fb, x0, y, _tableWidth, _rowHeight)
                                End Using
                            End If
                            PaintRow(g, tableFont, _tableModel.FooterRow, x0, y, cfg, isHeader:=False, isFooter:=True)
                            y += _rowHeight
                        End If

                        ' ── Linii verticale de grilă (peste toată înălțimea) ──────
                        If cfg.GridVisible Then
                            Dim vx As Integer = x0
                            For i As Integer = 0 To _colWidths.Length - 2   ' nu după ultima coloană
                                vx += _colWidths(i)
                                g.DrawLine(gridPen, vx, PADDING_V, vx, PADDING_V + _tableHeight)
                            Next
                        End If

                    End Using
                End Using

            Catch ex As Exception
                TreeLogger.Err($"Error painting tooltip table: {ex.Message}", "TooltipPopup.PaintTable")
            End Try
        End Sub

        Private Sub PaintRow(g As Graphics, baseFont As Font, row As TtRow,
                             x0 As Integer, y As Integer, cfg As TableConfig,
                             isHeader As Boolean, isFooter As Boolean)

            Using fmt As New StringFormat(StringFormat.GenericTypographic)
                fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces Or StringFormatFlags.NoWrap
                fmt.LineAlignment = StringAlignment.Center
                fmt.Trimming = StringTrimming.EllipsisCharacter

                Dim x As Integer = x0
                For i As Integer = 0 To _colWidths.Length - 1
                    Dim colW As Integer = _colWidths(i)
                    Dim cell As TtCell = RowCellAt(row, i)

                    If cell IsNot Nothing Then
                        ' Fundal celulă (suprascrie fundalul rândului)
                        If Not cell.BackColor.IsEmpty Then
                            Using cb As New SolidBrush(cell.BackColor)
                                g.FillRectangle(cb, x, y, colW, _rowHeight)
                            End Using
                        End If

                        If Not String.IsNullOrEmpty(cell.Text) Then
                            ' Font derivat per celulă
                            Dim fs As FontStyle = FontStyle.Regular
                            If cell.Bold Then fs = fs Or FontStyle.Bold
                            If cell.Italic OrElse (isFooter AndAlso cfg.FooterItalic) Then fs = fs Or FontStyle.Italic

                            Using cf As New Font(baseFont, fs)
                                ' Culoare text: explicit > header-default > tooltip-default
                                Dim fore As Color
                                If Not cell.ForeColor.IsEmpty Then
                                    fore = cell.ForeColor
                                ElseIf isHeader Then
                                    fore = cfg.HeaderForeColor
                                Else
                                    fore = TT_ForeColor
                                End If

                                ' Aliniere
                                Select Case cell.Align
                                    Case HorizontalAlignment.Right : fmt.Alignment = StringAlignment.Far
                                    Case HorizontalAlignment.Center : fmt.Alignment = StringAlignment.Center
                                    Case Else : fmt.Alignment = StringAlignment.Near
                                End Select

                                Dim cellRect As New RectangleF(
                                    x + cfg.CellPaddingH,
                                    y + cfg.CellPaddingV,
                                    Math.Max(0, colW - cfg.CellPaddingH * 2),
                                    Math.Max(0, _rowHeight - cfg.CellPaddingV * 2))

                                Using tb As New SolidBrush(fore)
                                    g.DrawString(cell.Text, cf, tb, cellRect, fmt)
                                End Using
                            End Using
                        End If
                    End If

                    x += colW
                Next
            End Using
        End Sub

        ' ====================================================================
        '  HELPERS
        ' ====================================================================

        ' Lățimea conținutului unei celule (fără padding), cu fontul derivat (Bold/Italic).
        Private Function MeasureCellContent(g As Graphics, baseFont As Font, cell As TtCell) As Integer
            If cell Is Nothing OrElse String.IsNullOrEmpty(cell.Text) Then Return 0

            Dim fs As FontStyle = FontStyle.Regular
            If cell.Bold Then fs = fs Or FontStyle.Bold
            If cell.Italic Then fs = fs Or FontStyle.Italic

            Using f As New Font(baseFont, fs)
                Using fmt As New StringFormat(StringFormat.GenericTypographic)
                    fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces
                    Dim sz As SizeF = g.MeasureString(cell.Text, f, PointF.Empty, fmt)
                    Return CInt(Math.Ceiling(sz.Width))
                End Using
            End Using
        End Function

        ' Celula i dintr-un rând, sau Nothing dacă rândul lipsește / are mai puține celule (ragged).
        Private Function RowCellAt(row As TtRow, i As Integer) As TtCell
            If row Is Nothing Then Return Nothing
            If i < 0 OrElse i >= row.Cells.Count Then Return Nothing
            Return row.Cells(i)
        End Function

        ' Strânge proporțional DOAR coloanele auto ca Σ ≤ maxWidth, cu plafon minim.
        ' Coloanele fixe rămân neatinse; dacă fixele singure depășesc maxWidth,
        ' coloanele auto cad la plafon și se acceptă overflow.
        Private Sub ShrinkAutoColumns(ByRef widths() As Integer, isFixed() As Boolean, maxWidth As Integer)
            Dim fixedSum As Integer = 0
            Dim autoSum As Integer = 0
            For i As Integer = 0 To widths.Length - 1
                If isFixed(i) Then fixedSum += widths(i) Else autoSum += widths(i)
            Next
            If autoSum <= 0 Then Return   ' nimic auto de strâns

            Dim avail As Integer = maxWidth - fixedSum
            Dim scale As Double = If(avail > 0, avail / CDbl(autoSum), 0.0)

            For i As Integer = 0 To widths.Length - 1
                If Not isFixed(i) Then
                    Dim w As Integer = CInt(Math.Floor(widths(i) * scale))
                    If w < MIN_AUTO_COL_WIDTH Then w = MIN_AUTO_COL_WIDTH
                    widths(i) = w
                End If
            Next
        End Sub

        ' ── DIAGNOSTIC TEMPORAR (de șters cu TT_MeasureDump) ──
        Friend Function DebugMeasure(model As TooltipTableModel) As String
            _tableModel = model
            MeasureTable()
            Return $"colWidths = [{String.Join(", ", _colWidths)}]" & vbLf &
                   $"rowHeight = {_rowHeight}" & vbLf &
                   $"tableWidth = {_tableWidth}  (+ PADDING_H*2 = {_tableWidth + PADDING_H * 2})" & vbLf &
                   $"tableHeight = {_tableHeight}  (+ PADDING_V*2 = {_tableHeight + PADDING_V * 2})"
        End Function

    End Class
End Class