Imports System.Drawing.Drawing2D
Imports TREEVIEW_VBA.TreeLogger

Partial Public Class AdvancedTreeControl
    Private Class TooltipPopup
        Inherits Form

        Private Const PADDING_H As Integer = 10
        Private Const PADDING_V As Integer = 7
        Private Const MAX_WIDTH As Integer = 400
        Private Const CORNER_RADIUS As Integer = 6
        Private Const BORDER_COLOR_ARG As Integer = 180

        Private _parts As List(Of AdvancedTreeControl.RichTextPart)
        Private _lines As List(Of List(Of AdvancedTreeControl.RichTextPart))
        Private _lineHeight As Integer
        Private _contentWidth As Integer
        Private _contentHeight As Integer

        Private _autoHideTimer As New Timer() With {.Interval = 5000}

        Friend Sub New()
            Me.FormBorderStyle = FormBorderStyle.None
            Me.ShowInTaskbar = False
            Me.TopMost = True
            Me.BackColor = Color.FromArgb(255, 255, 232)   ' Galben tooltip clasic
            Me.Padding = New Padding(0)
            Me.DoubleBuffered = True
            Me.StartPosition = FormStartPosition.Manual

            AddHandler _autoHideTimer.Tick, Sub(s, e)
                                                _autoHideTimer.Stop()
                                                Me.Hide()
                                            End Sub
        End Sub

        ''' <summary>
        ''' Afișează tooltip-ul cu RichText la poziția screen indicată.
        ''' text poate conține taguri: &lt;b&gt;, &lt;i&gt;, &lt;u&gt;, &lt;color=#hex&gt;, &lt;back=#hex&gt;
        ''' </summary>
        Friend Sub ShowTooltip(text As String, baseFont As Font, baseColor As Color, screenPos As Point)
            _autoHideTimer.Stop()

            ' 1. Parsăm RichText-ul (reutilizăm exact logica din Painting.vb)
            _parts = AdvancedTreeControl.ParseRichText(text, baseFont, baseColor)

            ' 2. Calculăm dimensiunile conținutului
            MeasureContent(baseFont, baseColor)

            ' 3. Dimensionăm form-ul
            Dim formW As Integer = _contentWidth + PADDING_H * 2
            Dim formH As Integer = _contentHeight + PADDING_V * 2

            ' 4. Poziționare inteligentă (să nu iasă din ecran)
            Dim screen As Screen = Screen.FromPoint(screenPos)
            Dim posX As Integer = screenPos.X + 16
            Dim posY As Integer = screenPos.Y + 20

            If posX + formW > screen.WorkingArea.Right Then
                posX = screenPos.X - formW - 4
            End If
            If posY + formH > screen.WorkingArea.Bottom Then
                posY = screenPos.Y - formH - 4
            End If

            TreeLogger.Info($"Showing tooltip at ({posX}, {posY}), size ({formW}x{formH}) with text {text}", "TooltipPopup.ShowTooltip")

            Me.Location = New Point(posX, posY)
            Me.Size = New Size(formW, formH)

            Me.Show()
            Me.Activate()
            _autoHideTimer.Start()
        End Sub

        Private Sub MeasureContent(baseFont As Font, baseColor As Color)
            ' Împărțim parts în linii (după \n)
            _lines = New List(Of List(Of AdvancedTreeControl.RichTextPart))

            Dim currentLine As New List(Of AdvancedTreeControl.RichTextPart)

            For Each part In _parts
                If part.Text.Contains(vbLf) OrElse part.Text.Contains(vbCrLf) Then
                    ' Split pe newline
                    Dim subLines() As String = part.Text.Replace(vbCrLf, vbLf).Split(vbLf)
                    For i = 0 To subLines.Length - 1
                        If subLines(i).Length > 0 Then
                            Dim sub_part = part
                            sub_part.Text = subLines(i)
                            currentLine.Add(sub_part)
                        End If
                        If i < subLines.Length - 1 Then
                            _lines.Add(currentLine)
                            currentLine = New List(Of AdvancedTreeControl.RichTextPart)
                        End If
                    Next
                Else
                    currentLine.Add(part)
                End If
            Next
            If currentLine.Count > 0 Then _lines.Add(currentLine)
            If _lines.Count = 0 Then _lines.Add(New List(Of AdvancedTreeControl.RichTextPart))

            ' Calculăm dimensiunile
            Dim fmt As StringFormat = StringFormat.GenericTypographic
            fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces

            Dim maxLineW As Single = 0
            _lineHeight = baseFont.Height + 2

            Using g As Graphics = Me.CreateGraphics()
                For Each line In _lines
                    Dim lineW As Single = 0
                    Dim lineH As Integer = baseFont.Height + 2
                    For Each part In line
                        Dim sz As SizeF = g.MeasureString(If(part.Text = "", " ", part.Text), part.Font, PointF.Empty, fmt)
                        lineW += sz.Width
                        If part.Font.Height + 2 > lineH Then lineH = part.Font.Height + 2
                    Next
                    If lineW > maxLineW Then maxLineW = lineW
                    If lineH > _lineHeight Then _lineHeight = lineH
                Next
            End Using

            _contentWidth = Math.Min(CInt(Math.Ceiling(maxLineW)), MAX_WIDTH)
            _contentHeight = _lines.Count * _lineHeight
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            Dim g As Graphics = e.Graphics
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim rc As New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)

            ' Fundal cu colțuri rotunde
            Using path As GraphicsPath = GetRoundedRect(rc, CORNER_RADIUS)
                Using bgBrush As New SolidBrush(Me.BackColor)
                    g.FillPath(bgBrush, path)
                End Using
                Using borderPen As New Pen(Color.FromArgb(BORDER_COLOR_ARG, Color.DarkGoldenrod), 1)
                    g.DrawPath(borderPen, path)
                End Using
            End Using

            ' Desenare text linie cu linie
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

        Protected Overrides Sub OnMouseLeave(e As EventArgs)
            TreeLogger.Info("Tooltip mouse leave - hiding tooltip", "TooltipPopup.OnMouseLeave")
            MyBase.OnMouseLeave(e)
            _autoHideTimer.Stop()
            Me.Hide()
        End Sub

        Protected Overrides Sub OnDeactivate(e As EventArgs)
            MyBase.OnDeactivate(e)
            _autoHideTimer.Stop()
            Me.Hide()
        End Sub
        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return True
            End Get
        End Property

        Private Function GetRoundedRect(rect As Rectangle, radius As Integer) As GraphicsPath
            Dim path As New GraphicsPath()
            Dim d As Integer = radius * 2
            If d > rect.Width Then d = rect.Width
            If d > rect.Height Then d = rect.Height
            Dim arc As New Rectangle(rect.X, rect.Y, d, d)
            path.AddArc(arc, 180, 90)
            arc.X = rect.Right - d
            path.AddArc(arc, 270, 90)
            arc.Y = rect.Bottom - d
            path.AddArc(arc, 0, 90)
            arc.X = rect.X
            path.AddArc(arc, 90, 90)
            path.CloseFigure()
            Return path
        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                _autoHideTimer.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        Protected Overrides ReadOnly Property CreateParams() As CreateParams
            Get
                Dim cp As CreateParams = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or &H80        ' WS_EX_TOOLWINDOW 
                cp.ExStyle = cp.ExStyle Or &H8000000   ' WS_EX_NOACTIVATE 
                Return cp
            End Get
        End Property
    End Class
End Class
