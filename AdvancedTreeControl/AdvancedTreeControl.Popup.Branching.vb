Imports System.Drawing.Drawing2D

Partial Public Class AdvancedTreeControl
    Partial Private Class TooltipPopup

        ' Fundal rotunjit + bordură. Comun ambelor moduri (tabel / RichText).
        ' (extras 1:1 din OnPaint)
        Private Sub DrawBackground(g As Graphics)
            Dim rc As New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)
            Using path As GraphicsPath = GetRoundedRect(rc, CORNER_RADIUS)
                Using bgBrush As New SolidBrush(Me.BackColor)
                    g.FillPath(bgBrush, path)
                End Using
                Dim bc As Color = TT_BackColor
                Dim borderDerived As Color = Color.FromArgb(
                    CInt(bc.R * 0.6), CInt(bc.G * 0.6), CInt(bc.B * 0.6))
                Using borderPen As New Pen(Color.FromArgb(BORDER_COLOR_ARG, borderDerived), 1)
                    g.DrawPath(borderPen, path)
                End Using
            End Using
        End Sub

        ' Desenare text RichText, linie cu linie. (extras 1:1 din OnPaint)
        Private Sub PaintRichText(g As Graphics)
            Dim fmt As StringFormat = StringFormat.GenericTypographic
            fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces

            Dim y As Single = PADDING_V
            For Each line In _lines
                Dim x As Single = PADDING_H
                For Each part In line
                    Dim sz As SizeF = g.MeasureString(If(part.Text = "", " ", part.Text), part.Font, PointF.Empty, fmt)

                    If part.HasBackColor Then
                        Using bb As New SolidBrush(part.BackColor)
                            g.FillRectangle(bb, x, y, sz.Width, _lineHeight)
                        End Using
                    End If

                    Dim textY As Single = y + (_lineHeight - part.Font.Height) / 2.0F
                    Using tb As New SolidBrush(part.ForeColor)
                        g.DrawString(part.Text, part.Font, tb, x, textY, fmt)
                    End Using

                    x += sz.Width
                Next
                y += _lineHeight
            Next
        End Sub

        ' Aplică fundalul opac (extras din ShowTooltip ca să nu se dubleze pe cele 2 ramuri).
        Private Sub ApplyOpaqueBackColor()
            Try
                Me.BackColor = Color.FromArgb(255, TT_BackColor)
            Catch ex As Exception
                TreeLogger.Err($"Error setting tooltip back color to {TT_BackColor}: {ex.Message}", "TooltipPopup.ShowTooltip")
            End Try
        End Sub

    End Class
End Class