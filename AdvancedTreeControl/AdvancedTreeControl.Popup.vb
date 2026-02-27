Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

Partial Public Class AdvancedTreeControl
    Private Class TooltipPopup
        Inherits Form

        Private Const SW_SHOWNOACTIVATE As Integer = 4

        <DllImport("user32.dll")>
        Private Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        Private Const PADDING_H As Integer = 10
        Private Const PADDING_V As Integer = 7
        Private Const MAX_WIDTH As Integer = 400
        Private Const MAX_LINES As Integer = 10
        Private Const CORNER_RADIUS As Integer = 6
        Private Const BORDER_COLOR_ARG As Integer = 180

        ' ── Font tooltip ──────────────────────────────────────────
        Private Const TOOLTIP_FONT_NAME As String = "Segoe UI"
        Private Const TOOLTIP_FONT_SIZE As Single = 8.5F
        Private Const TOOLTIP_FORE_COLOR As String = "#333333"

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
                                                'stop timer from running if mouse is already over tooltip
                                                If Me.Bounds.Contains(Cursor.Position) Then
                                                    Return
                                                End If
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

            Dim tooltipFont As New Font(TOOLTIP_FONT_NAME, TOOLTIP_FONT_SIZE, FontStyle.Regular, GraphicsUnit.Point)

            ' 1. Parsăm RichText-ul (reutilizăm exact logica din Painting.vb)
            _parts = AdvancedTreeControl.ParseRichText(text, tooltipFont, baseColor)

            ' 2. Calculăm dimensiunile conținutului
            MeasureContent(baseFont, ColorTranslator.FromHtml(TOOLTIP_FORE_COLOR))

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

            Me.ShowWithoutActivating()
            'Me.Activate()
            _autoHideTimer.Start()
        End Sub

        Private Sub ShowWithoutActivating()
            ' Folosește ShowWindow cu SW_SHOWNOACTIVATE (&H4)
            ShowWindow(Me.Handle, SW_SHOWNOACTIVATE)
        End Sub

        Private Sub MeasureContent(baseFont As Font, baseColor As Color)
            Dim fmt As StringFormat = StringFormat.GenericTypographic
            fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces

            _lineHeight = baseFont.Height + 2

            Using g As Graphics = Me.CreateGraphics()

                ' ── Step 1: Split _parts în linii logice (după \n explicit) ──────────────
                Dim logicalLines As New List(Of List(Of AdvancedTreeControl.RichTextPart))
                Dim currentLine As New List(Of AdvancedTreeControl.RichTextPart)

                For Each part In _parts
                    If part.Text.Contains(vbLf) OrElse part.Text.Contains(vbCrLf) Then
                        Dim subLines() As String = part.Text.Replace(vbCrLf, vbLf).Split(vbLf)
                        For i = 0 To subLines.Length - 1
                            If subLines(i).Length > 0 Then
                                Dim sub_part = part
                                sub_part.Text = subLines(i)
                                currentLine.Add(sub_part)
                            End If
                            If i < subLines.Length - 1 Then
                                logicalLines.Add(currentLine)
                                currentLine = New List(Of AdvancedTreeControl.RichTextPart)
                            End If
                        Next
                    Else
                        currentLine.Add(part)
                    End If
                Next
                If currentLine.Count > 0 Then logicalLines.Add(currentLine)
                If logicalLines.Count = 0 Then logicalLines.Add(New List(Of AdvancedTreeControl.RichTextPart))

                ' ── Step 2: Word-wrap fiecare linie logică → linii vizuale ───────────────
                Dim allVisualLines As New List(Of List(Of AdvancedTreeControl.RichTextPart))
                For Each logLine In logicalLines
                    For Each vLine In WrapLogicalLine(logLine, MAX_WIDTH, g, fmt)
                        allVisualLines.Add(vLine)
                    Next
                Next

                ' ── Step 3: Aplică MAX_LINES — trunchează cu "…" dacă e nevoie ───────────
                Dim truncated As Boolean = allVisualLines.Count > MAX_LINES
                If truncated Then
                    _lines = allVisualLines.GetRange(0, MAX_LINES)
                    ' Adaugă "…" la ultimul part din ultima linie
                    Dim lastLine = _lines(MAX_LINES - 1)
                    If lastLine.Count > 0 Then
                        Dim lp = lastLine(lastLine.Count - 1)
                        lastLine(lastLine.Count - 1) = New AdvancedTreeControl.RichTextPart With {
                    .Text = lp.Text.TrimEnd() & "…",
                    .Font = lp.Font,
                    .ForeColor = lp.ForeColor,
                    .BackColor = lp.BackColor,
                    .HasBackColor = lp.HasBackColor
                }
                    Else
                        lastLine.Add(New AdvancedTreeControl.RichTextPart With {
                    .Text = "…",
                    .Font = baseFont,
                    .ForeColor = baseColor,
                    .HasBackColor = False
                })
                    End If
                Else
                    _lines = allVisualLines
                End If

                ' ── Step 4: Calculează dimensiunile finale ────────────────────────────────
                Dim maxLineW As Single = 0
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

                _contentWidth = CInt(Math.Ceiling(maxLineW))
                _contentHeight = _lines.Count * _lineHeight
            End Using
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

        Private Const WM_MOUSEACTIVATE As Integer = &H21
        Private Const WM_ACTIVATE As Integer = &H6
        Private Const WM_ACTIVATEAPP As Integer = &H1C
        Private Const WM_SETFOCUS As Integer = &H7
        Private Const MA_NOACTIVATE As Integer = 3

        Protected Overrides Sub WndProc(ByRef m As Message)
            Select Case m.Msg
                Case WM_MOUSEACTIVATE
                    m.Result = CType(MA_NOACTIVATE, IntPtr)
                    Return

                Case WM_ACTIVATE, WM_ACTIVATEAPP
                    ' Previne activarea complet
                    If m.WParam <> IntPtr.Zero Then
                        m.Result = IntPtr.Zero
                        Return
                    End If

                Case WM_SETFOCUS
                    ' Refuză focusul
                    m.Result = IntPtr.Zero
                    Return
            End Select

            MyBase.WndProc(m)
        End Sub
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
                cp.ExStyle = cp.ExStyle Or &H80       ' WS_EX_TOOLWINDOW
                cp.ExStyle = cp.ExStyle Or &H8000000  ' WS_EX_NOACTIVATE
                cp.ExStyle = cp.ExStyle Or &H20       ' WS_EX_TRANSPARENT (opțional, pentru click-through)

                ' IMPORTANT: Adaugă stilul de fereastră pentru a preveni focusul
                cp.Style = cp.Style And Not &H8000000L ' Scoate WS_BORDER dacă există

                Return cp
            End Get
        End Property

        ''' <summary>
        ''' Împarte o linie logică (List(Of RichTextPart)) în mai multe linii vizuale
        ''' care se încadrează în maxWidth. Rupe la word boundary; dacă un cuvânt
        ''' singur depășește maxWidth, îl rupe forțat caracter cu caracter.
        ''' </summary>
        Private Shared Function WrapLogicalLine(
            parts As List(Of AdvancedTreeControl.RichTextPart),
            maxWidth As Integer,
            g As Graphics,
            fmt As StringFormat) As List(Of List(Of AdvancedTreeControl.RichTextPart))

            Dim result As New List(Of List(Of AdvancedTreeControl.RichTextPart))
            Dim curLine As New List(Of AdvancedTreeControl.RichTextPart)
            Dim curW As Single = 0

            For Each part In parts
                Dim tokens As List(Of String) = SplitToWordTokens(part.Text)

                For Each token In tokens
                    ' La începutul unui rând nou eliminăm spațiile de leading
                    Dim word As String = If(curW = 0, token.TrimStart(" "c), token)
                    If word.Length = 0 Then Continue For

                    Dim wordW As Single = g.MeasureString(word, part.Font, PointF.Empty, fmt).Width

                    If wordW > maxWidth Then
                        ' ── Cuvânt prea lung: flush linie curentă, apoi rupe caracter cu caracter ──
                        If curLine.Count > 0 Then
                            result.Add(curLine)
                            curLine = New List(Of AdvancedTreeControl.RichTextPart)
                            curW = 0
                        End If
                        Dim remaining As String = word.TrimStart(" "c)
                        While remaining.Length > 0
                            Dim chunk As String = ""
                            For c As Integer = 1 To remaining.Length
                                Dim test As String = remaining.Substring(0, c)
                                If g.MeasureString(test, part.Font, PointF.Empty, fmt).Width <= maxWidth Then
                                    chunk = test
                                Else
                                    Exit For
                                End If
                            Next
                            If chunk.Length = 0 Then chunk = remaining.Substring(0, 1) ' safety

                            curLine.Add(New AdvancedTreeControl.RichTextPart With {
                                .Text = chunk,
                                .Font = part.Font,
                                .ForeColor = part.ForeColor,
                                .BackColor = part.BackColor,
                                .HasBackColor = part.HasBackColor
                            })
                            curW = g.MeasureString(chunk, part.Font, PointF.Empty, fmt).Width
                            remaining = remaining.Substring(chunk.Length)

                            If remaining.Length > 0 Then
                                result.Add(curLine)
                                curLine = New List(Of AdvancedTreeControl.RichTextPart)
                                curW = 0
                            End If
                        End While

                    ElseIf curW + wordW <= maxWidth Then
                        ' ── Încape pe linia curentă ───────────────────────────────────────────
                        curLine.Add(New AdvancedTreeControl.RichTextPart With {
                            .Text = word,
                            .Font = part.Font,
                            .ForeColor = part.ForeColor,
                            .BackColor = part.BackColor,
                            .HasBackColor = part.HasBackColor
                        })
                        curW += wordW

                    Else
                        ' ── Nu încape: rupe linia, pune cuvântul pe rândul următor ───────────
                        If curLine.Count > 0 Then
                            result.Add(curLine)
                            curLine = New List(Of AdvancedTreeControl.RichTextPart)
                            curW = 0
                        End If
                        Dim trimmed As String = token.TrimStart(" "c)
                        If trimmed.Length > 0 Then
                            curLine.Add(New AdvancedTreeControl.RichTextPart With {
                                .Text = trimmed,
                                .Font = part.Font,
                                .ForeColor = part.ForeColor,
                                .BackColor = part.BackColor,
                                .HasBackColor = part.HasBackColor
                            })
                            curW = g.MeasureString(trimmed, part.Font, PointF.Empty, fmt).Width
                        End If
                    End If
                Next
            Next

            If curLine.Count > 0 Then result.Add(curLine)
            If result.Count = 0 Then result.Add(New List(Of AdvancedTreeControl.RichTextPart))
            Return result
        End Function

        ''' <summary>
        ''' Tokenizează un string la granițe de cuvânt.
        ''' Rezultatul: fiecare token = cuvânt + spațiile trailing ale lui.
        ''' Ex: "hello world  foo" → ["hello ", "world  ", "foo"]
        ''' </summary>
        Private Shared Function SplitToWordTokens(text As String) As List(Of String)
            Dim tokens As New List(Of String)
            If String.IsNullOrEmpty(text) Then Return tokens
            Dim i As Integer = 0
            While i < text.Length
                Dim start As Integer = i
                ' Citim caractere non-spațiu
                While i < text.Length AndAlso text(i) <> " "c
                    i += 1
                End While
                ' Citim spațiile trailing
                While i < text.Length AndAlso text(i) = " "c
                    i += 1
                End While
                tokens.Add(text.Substring(start, i - start))
            End While
            Return tokens
        End Function
    End Class
End Class
