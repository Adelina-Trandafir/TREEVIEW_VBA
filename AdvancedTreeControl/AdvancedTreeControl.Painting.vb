Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

Partial Public Class AdvancedTreeControl
    Private Sub DrawItem(g As Graphics, it As TreeItem, y As Integer)
        ' Normalizare
        If it.Level = 0 AndAlso Not _RootExpander AndAlso Not it.Expanded Then it.Expanded = True
        If Me.ExpanderSize Mod 2 <> 0 Then Me.ExpanderSize -= 1
        If Me.ItemHeight Mod 2 <> 0 Then Me.ItemHeight -= 1

        ' ── Calcul layout comun ──────────────────────────────────────────────────
        Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START
        Dim expanderCenterX As Integer = gridLeft + (Indent \ 2)
        Dim midY As Integer = y + (ItemHeight \ 2)
        Dim expanderRect As New Rectangle(
        expanderCenterX - (ExpanderSize \ 2),
        midY - (ExpanderSize \ 2),
        ExpanderSize, ExpanderSize)
        Dim xBase As Integer = If(it.Level = 0 AndAlso Not _RootExpander,
                              gridLeft,
                              gridLeft + Indent + PADDING_EXPANDER_GAP)

        ' ── Pasul 1: Selecție & Hover ──────────────────────────────────────────── 
        DrawSelection(g, it, y, gridLeft, xBase)

        ' ── Pasul 2: Linii arbore ──────────────────────────────────────────────── 
        If Not (it.Level = 0 AndAlso Not _RootExpander) Then
            DrawTreeLines(g, it, y, expanderCenterX, midY, gridLeft)
        End If

        ' ── Loader (ieșire anticipată) ─────────────────────────────────────────── 
        If it.IsLoader Then
            DrawLoaderItem(g, it, y, xBase)
            Return
        End If

        ' ── Pasul 3: Checkbox / RadioButton ─────────────────────────────────────- 
        xBase = DrawCheckbox(g, it, y, xBase, midY)

        ' ── Pasul 5: Expander ──────────────────────────────────────────────────── 
        DrawExpander(g, it, expanderRect, expanderCenterX, midY)

        ' ── Pasul 6: Conținut (icon stânga + text) ─────────────────────────────── 
        DrawContent(g, it, y, xBase)

        ' ── Pasul 7: Iconiță dreapta ───────────────────────────────────────────── 
        DrawRightIcon(g, it, y)
    End Sub

    Private Sub DrawSelection(g As Graphics, it As TreeItem, y As Integer, gridLeft As Integer, xBase As Integer)
        Dim selStartX As Integer
        If it.Level = 0 AndAlso Not _RootExpander Then
            selStartX = gridLeft
        ElseIf Not _RootExpander Then
            selStartX = xBase
        Else
            selStartX = gridLeft + ExpanderSize * 2 - 3
        End If

        Dim selWidth As Integer = Math.Max(0, Me.ClientSize.Width - selStartX - PADDING_TREE_END)
        Dim fullRowRect As New Rectangle(selStartX, y, selWidth, ItemHeight)

        Dim oldSmooth = g.SmoothingMode
        g.SmoothingMode = SmoothingMode.AntiAlias

        If it Is pSelectedItem Then
            Using path As GraphicsPath = GetRoundedRect(fullRowRect, SELECTION_CORNER_RADIUS)
                Using brush As New SolidBrush(SelectedBackColor)
                    g.FillPath(brush, path)
                End Using
                Dim borderRect As New Rectangle(fullRowRect.X, fullRowRect.Y, fullRowRect.Width - 1, fullRowRect.Height - 1)
                Using borderPath As GraphicsPath = GetRoundedRect(borderRect, SELECTION_CORNER_RADIUS)
                    Using pen As New Pen(SelectedBorderColor)
                        g.DrawPath(pen, borderPath)
                    End Using
                End Using
            End Using
        ElseIf it Is pHoveredItem Then
            Using path As GraphicsPath = GetRoundedRect(fullRowRect, SELECTION_CORNER_RADIUS)
                Using brush As New SolidBrush(HoverBackColor)
                    g.FillPath(brush, path)
                End Using
            End Using
        End If

        g.SmoothingMode = oldSmooth
    End Sub

    Private Sub DrawLoaderItem(g As Graphics, it As TreeItem, y As Integer, xBase As Integer)
        Dim loaderY As Integer = y + (ItemHeight - 14) \ 2
        Dim textY As Integer = y + (ItemHeight - Me.Font.Height) \ 2 + 1

        Dim oldSmooth = g.SmoothingMode
        g.SmoothingMode = SmoothingMode.AntiAlias

        Using p As New Pen(Color.DimGray, 2)
            g.DrawArc(p, xBase, loaderY, 14, 14, loadingAngle, 300)
        End Using
        g.DrawString("Se încarcă...", Me.Font, Brushes.Gray, xBase + 20, textY)

        g.SmoothingMode = oldSmooth
    End Sub

    Private Function DrawCheckbox(g As Graphics, it As TreeItem, y As Integer, xBase As Integer, midY As Integer) As Integer
        If Not NodeHasCheckControl(it) Then Return xBase

        Dim chkSize As Integer = _checkBoxSize
        Dim chkRect As New Rectangle(xBase, midY - (chkSize \ 2), chkSize, chkSize)
        Dim accentColor As Color = Color.DodgerBlue
        Dim borderColor As Color = Color.FromArgb(180, 180, 180)

        Dim oldSmooth = g.SmoothingMode
        g.SmoothingMode = SmoothingMode.AntiAlias

        If _radioButtonLevel >= 0 AndAlso it.Level = _radioButtonLevel Then
            ' *** RADIO BUTTON ***
            If it.IsRadioSelected Then
                Using brush As New SolidBrush(accentColor) : g.FillEllipse(brush, chkRect) : End Using
                Using pen As New Pen(accentColor) : g.DrawEllipse(pen, chkRect) : End Using
                Dim dotMargin As Integer = CInt(chkSize * 0.28F)
                Dim dotRect As New Rectangle(chkRect.X + dotMargin, chkRect.Y + dotMargin,
                                         chkSize - dotMargin * 2, chkSize - dotMargin * 2)
                Using brush As New SolidBrush(Color.White) : g.FillEllipse(brush, dotRect) : End Using
            Else
                g.FillEllipse(Brushes.White, chkRect)
                Using pen As New Pen(borderColor, 1) : g.DrawEllipse(pen, chkRect) : End Using
            End If
        Else
            ' *** CHECKBOX STANDARD ***
            Using path As GraphicsPath = GetRoundedRect(chkRect, 3)
                Select Case it.CheckState
                    Case TreeCheckState.Checked
                        Using brush As New SolidBrush(accentColor) : g.FillPath(brush, path) : End Using
                        Using pen As New Pen(accentColor) : g.DrawPath(pen, path) : End Using
                        Using penTick As New Pen(Color.White, 2.0F)
                            penTick.StartCap = LineCap.Round
                            penTick.EndCap = LineCap.Round
                            penTick.LineJoin = LineJoin.Round
                            g.DrawLines(penTick, {
                            New PointF(chkRect.X + chkSize * 0.22F, chkRect.Y + chkSize * 0.52F),
                            New PointF(chkRect.X + chkSize * 0.42F, chkRect.Y + chkSize * 0.72F),
                            New PointF(chkRect.X + chkSize * 0.78F, chkRect.Y + chkSize * 0.28F)
                        })
                        End Using

                    Case TreeCheckState.Indeterminate
                        Using brush As New SolidBrush(accentColor) : g.FillPath(brush, path) : End Using
                        Using pen As New Pen(accentColor) : g.DrawPath(pen, path) : End Using
                        Using penDash As New Pen(Color.White, 2.0F)
                            penDash.StartCap = LineCap.Round
                            penDash.EndCap = LineCap.Round
                            Dim yMid As Single = chkRect.Y + (chkRect.Height / 2.0F)
                            Dim margin As Single = chkSize * 0.25F
                            g.DrawLine(penDash, chkRect.X + margin, yMid, chkRect.Right - margin, yMid)
                        End Using

                    Case Else ' Unchecked
                        g.FillPath(Brushes.White, path)
                        Using pen As New Pen(borderColor, 1) : g.DrawPath(pen, path) : End Using
                End Select
            End Using
        End If

        g.SmoothingMode = oldSmooth
        Return xBase + chkSize + PADDING_CHECKBOX_GAP
    End Function

    Private Sub DrawExpander(g As Graphics, it As TreeItem, expanderRect As Rectangle, expanderCenterX As Integer, midY As Integer)
        Dim showExpander As Boolean = (it.Children.Count > 0 OrElse it.LazyNode)
        If it.Level = 0 AndAlso Not _RootExpander Then showExpander = False
        If Not showExpander Then Return

        g.FillRectangle(Brushes.White, expanderRect)
        g.DrawRectangle(New Pen(LineColor), expanderRect)
        g.DrawLine(Pens.Black, expanderRect.Left + 2, midY, expanderRect.Right - 2, midY)
        If Not it.Expanded Then
            g.DrawLine(Pens.Black, expanderCenterX, expanderRect.Top + 2, expanderCenterX, expanderRect.Bottom - 2)
        End If
    End Sub

    Private Sub DrawContent(g As Graphics, it As TreeItem, y As Integer, xBase As Integer)
        ' ── Icon stânga ──────────────────────────────────────────────────────────
        Dim leftIconRect As New Rectangle(xBase, y + (ItemHeight - LeftIconSize.Height) \ 2,
                                      LeftIconSize.Width, LeftIconSize.Height)
        If it.TextWidth = -1 Then it.TextWidth = CInt(g.MeasureString(it.Caption, Me.Font).Width)

        If _hasNodeIcons Then
            Dim icon As Image = If(it.Expanded, it.LeftIconOpen, it.LeftIconClosed)
            If icon IsNot Nothing Then g.DrawImage(icon, leftIconRect)
        End If

        ' ── Calcul limite text ───────────────────────────────────────────────────
        Dim textX As Integer = If(it.LeftIconClosed IsNot Nothing AndAlso _hasNodeIcons,
                                 leftIconRect.Right + PADDING_ICON_GAP, xBase)
        Dim scrollW As Integer = ScrollBarWidth 'If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
        Dim maxRightX As Integer = Me.Width - scrollW - PADDING_TREE_END
        If it.RightIcon IsNot Nothing Then maxRightX -= (RightIconSize.Width + PADDING_RIGHT_ICON_GAP)
        Dim availableTextWidth As Integer = Math.Max(0, maxRightX - textX)

        ' ── Font & culoare ───────────────────────────────────────────────────────
        Dim baseTextColor As Color = If(it.NodeForeColor <> Color.Empty, it.NodeForeColor,
                                 If(Me.ForeColor <> Color.Empty, Me.ForeColor, Color.Black))

        Dim nodeStyle As FontStyle = Me.TreeFont.Style
        If it.Bold Then nodeStyle = nodeStyle Or FontStyle.Bold
        If it.Italic Then nodeStyle = nodeStyle Or FontStyle.Italic
        Dim nodeFont As Font = If(nodeStyle <> Me.Font.Style, New Font(Me.Font, nodeStyle), Me.TreeFont)

        ' ── BackColor per nod ────────────────────────────────────────────────────
        If it.NodeBackColor <> Color.Empty AndAlso it IsNot pSelectedItem Then
            Using bgBrush As New SolidBrush(it.NodeBackColor)
                g.FillRectangle(bgBrush, New Rectangle(textX, y, availableTextWidth, ItemHeight))
            End Using
        End If

        ' ── Clip + text ──────────────────────────────────────────────────────────
        Dim oldClip As Region = g.Clip.Clone()
        g.SetClip(New Rectangle(textX, y, availableTextWidth, ItemHeight))
        DrawRichText(g, it.Caption, textX, y, nodeFont, baseTextColor, availableTextWidth)
        g.Clip = oldClip

        ' ── TreeListView: deseneaza celulele coloanelor ──────────────────────────
        If _treeListViewEnabled AndAlso _treeListView AndAlso _columns.Count > 0 Then
            Try
                Dim totalColsW As Integer = 0
                For Each cd In _columns
                    totalColsW += cd.Width
                Next
                Dim colStartX As Integer = Me.Width - ScrollBarWidth - PADDING_TREE_END - totalColsW
                _captionColumnEndX = colStartX

                Dim cx As Integer = colStartX
                Dim colFmt As New StringFormat()
                colFmt.LineAlignment = StringAlignment.Center
                colFmt.Trimming = StringTrimming.EllipsisCharacter

                Try
                    Using sepPen As New Pen(Color.FromArgb(COLUMN_SEPARATOR_COLOR_ALPHA, LineColor), 1)
                        g.DrawLine(sepPen, cx, y, cx, y + ItemHeight)
                    End Using

                    For i As Integer = 0 To _columns.Count - 1
                        Try
                            Dim cd = _columns(i)
                            Dim cellRect As New Rectangle(cx, y, cd.Width, ItemHeight)

                            Dim cellData As TreeItem.CellData = Nothing
                            it.Cells.TryGetValue(cd.Name, cellData)

                            If cellData IsNot Nothing AndAlso cellData.BackColor <> Color.Empty AndAlso it IsNot pSelectedItem Then
                                Using bgBrush As New SolidBrush(cellData.BackColor)
                                    g.FillRectangle(bgBrush, cellRect)
                                End Using
                            End If

                            Using sepPen As New Pen(Color.FromArgb(COLUMN_SEPARATOR_COLOR_ALPHA, LineColor), 1)
                                g.DrawLine(sepPen, cx + cd.Width - 1, y, cx + cd.Width - 1, y + ItemHeight)
                            End Using

                            Dim cellVal As String = If(cellData IsNot Nothing, cellData.Value, "")
                            Dim cellFore As Color = If(cellData IsNot Nothing AndAlso cellData.ForeColor <> Color.Empty,
                                                       cellData.ForeColor, baseTextColor)

                            colFmt.Alignment = ColAlignToStringAlign(cd.Align)

                            Dim textPaddedRect As New Rectangle(cellRect.X + 4, cellRect.Y, cellRect.Width - 8, cellRect.Height)
                            Using fgBrush As New SolidBrush(cellFore)
                                g.DrawString(cellVal, nodeFont, fgBrush, textPaddedRect, colFmt)
                            End Using

                            cx += cd.Width
                        Catch
                            cx += If(i < _columns.Count, _columns(i).Width, 0)
                        End Try
                    Next
                Finally
                    colFmt.Dispose()
                End Try
            Catch
            End Try
        End If
        ' ── SFARSIT TreeListView cells ───────────────────────────────────────────
    End Sub

    Private Sub DrawRightIcon(g As Graphics, it As TreeItem, y As Integer)
        If it.RightIcon Is Nothing Then Return

        ' Hover-only: activat global SAU per nod
        ' Spațiul din dreapta e rezervat întotdeauna (DrawContent verifică it.RightIcon IsNot Nothing),
        ' deci textul NU sare la hover — Varianta A garantată fără modificări în DrawContent.
        Dim hoverOnly As Boolean = _showRightIconOnHover OrElse it.ShowRightIconOnHover
        If hoverOnly AndAlso it IsNot pHoveredItem Then Return

        Dim scrollW As Integer = ScrollBarWidth 'If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
        Dim rx As Integer = Me.Width - RightIconSize.Width - PADDING_RIGHT_ICON_GAP - PADDING_TREE_END - scrollW
        Dim ry As Integer = y + (ItemHeight - RightIconSize.Height) \ 2
        g.DrawImage(it.RightIcon, rx, ry, RightIconSize.Width, RightIconSize.Height)
    End Sub

    Private Sub DrawTreeLines(g As Graphics, it As TreeItem, y As Integer, expCenterX As Integer, midY As Integer, currentGridLeft As Integer)

        Using p As New Pen(LineColor)
            p.DashStyle = DashStyle.Dot

            ' ------------------------------------------------------------------
            ' PASUL 0. TRUNCHI JOS (din expander spre primul copil)
            '    Desenat pe COLOANA NODULUI CURENT (expCenterX).
            '    Acoperă golul din rândul părintelui: de la baza expanderului
            '    până la marginea de jos a rândului curent.
            '    Condiție: nodul are copii vizibili (expanded).
            '    Expanderul (white fill) se desenează DUPĂ în Pasul 5 → îl acoperă.
            ' ------------------------------------------------------------------
            If (it.Children.Count > 0 OrElse it.LazyNode) AndAlso it.Expanded Then
                Dim trunkStartY As Integer = midY + (ExpanderSize \ 2) + 1  ' imediat sub expander
                Dim trunkEndY As Integer = y + ItemHeight                  ' baza rândului curent
                If trunkStartY < trunkEndY Then
                    g.DrawLine(p, expCenterX, trunkStartY, expCenterX, trunkEndY)
                End If
            End If

            ' Nodurile root fără _RootExpander nu au trunchi ascendent → ieșim
            If it.Level = 0 Then Return

            ' X-ul trunchiului vertical = coloana expanderului PĂRINTELUI
            Dim parentColX As Integer = ((it.Level - 1) * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START + (Indent \ 2)

            ' Capătul drept al liniei orizontale = imediat înainte de conținut
            Dim hLineEnd As Integer = currentGridLeft + Indent + PADDING_EXPANDER_GAP - TREE_LINE_H_MARGIN

            ' ------------------------------------------------------------------
            ' 1. LINIA ORIZONTALĂ — de la trunchiul părintelui → înainte de conținut
            ' ------------------------------------------------------------------
            g.DrawLine(p, parentColX, midY, hLineEnd, midY)

            ' ------------------------------------------------------------------
            ' 2. LINIA VERTICALĂ SUS — jumătatea superioară a rândului curent
            ' ------------------------------------------------------------------
            If it.Parent IsNot Nothing Then
                g.DrawLine(p, parentColX, y, parentColX, midY)
            End If

            ' ------------------------------------------------------------------
            ' 3. LINIA VERTICALĂ JOS — jumătatea inferioară (dacă urmează un frate)
            ' ------------------------------------------------------------------
            If it.Parent IsNot Nothing AndAlso Not it.IsLastSibling Then
                g.DrawLine(p, parentColX, midY, parentColX, y + ItemHeight)
            End If

            ' ------------------------------------------------------------------
            ' 4. LINIILE VERTICALE ALE STRĂMOȘILOR (continuare trunchi prin rând)
            ' ------------------------------------------------------------------
            Dim ancestor As TreeItem = it.Parent
            While ancestor IsNot Nothing AndAlso ancestor.Parent IsNot Nothing
                If Not ancestor.IsLastSibling Then
                    Dim ancParentColX As Integer = ((ancestor.Level - 1) * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START + (Indent \ 2)
                    g.DrawLine(p, ancParentColX, y, ancParentColX, y + ItemHeight)
                End If
                ancestor = ancestor.Parent
            End While

        End Using
    End Sub

    ' =================================================================================
    ' SUPORT RICH TEXT (BOLD, ITALIC, COLOR)
    ' =================================================================================
    ' AdvancedTreeControl.Painting.vb

    Private Sub DrawRichText(g As Graphics, text As String, x As Integer, y As Integer,
                          defaultFont As Font, defaultColor As Color, availableWidth As Integer)
        ' ── 1. Split separator ──────────────────────────────────────────────────
        Dim leftText As String = text
        Dim rightText As String = ""
        Dim hasSplit As Boolean = False

        If text.Contains("~~~") Then
            Dim partsStr = text.Split({"~~~"}, StringSplitOptions.None)
            leftText = partsStr(0)
            If partsStr.Length > 1 Then rightText = partsStr(1)
            hasSplit = True
        End If

        Dim fmt As StringFormat = StringFormat.GenericTypographic
        fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces

        Dim rightEdgeX As Single = CSng(x) + CSng(availableWidth)
        Dim hasLeftProp As Boolean = (m_LeftTextWidth > 0)
        Dim hasRightProp As Boolean = (m_RightTextWidth > 0)

        ' ── 2. Calcul zone stânga/dreapta ────────────────────────────────────────
        Dim leftBudget As Single = availableWidth
        Dim rightZoneStart As Single = rightEdgeX
        Dim rightBudget As Single = 0
        Dim caseA As Boolean = False

        If hasSplit Then
            If Not hasLeftProp AndAlso Not hasRightProp Then
                ' Case A: nicio constrângere — stânga liberă, dreapta show/hide
                caseA = True
                leftBudget = availableWidth

            ElseIf hasLeftProp AndAlso hasRightProp Then
                ' Case B: ambele setate — stânga prioritate
                leftBudget = Math.Min(CSng(m_LeftTextWidth), CSng(availableWidth))
                Dim naturalRS = rightEdgeX - CSng(m_RightTextWidth)
                Dim forcedRS = CSng(x) + leftBudget + PADDING_SEPARATOR_GAP
                rightZoneStart = Math.Max(naturalRS, forcedRS)
                rightBudget = Math.Max(0, rightEdgeX - rightZoneStart)

            ElseIf hasRightProp Then
                ' Case C: doar dreapta setată — rezervare fixă din dreapta
                rightBudget = CSng(m_RightTextWidth)
                rightZoneStart = rightEdgeX - rightBudget
                leftBudget = Math.Max(0, rightZoneStart - CSng(x) - PADDING_SEPARATOR_GAP)

            Else
                ' Case D: doar stânga setată — stânga are budget fix
                leftBudget = Math.Min(CSng(m_LeftTextWidth), CSng(availableWidth))
                rightZoneStart = CSng(x) + leftBudget + PADDING_SEPARATOR_GAP
                rightBudget = Math.Max(0, rightEdgeX - rightZoneStart)
            End If
        End If

        ' ── 3. Desenare stânga ───────────────────────────────────────────────────
        Dim leftParts As List(Of RichTextPart) = ParseRichText(leftText, defaultFont, defaultColor)
        Dim currentX As Single

        If Not hasSplit OrElse caseA Then
            ' Case A / fără separator: desenare liberă (clipping extern gestionează limita)
            currentX = CSng(x)
            For Each part In leftParts
                Dim sz As SizeF = g.MeasureString(part.Text, part.Font, PointF.Empty, fmt)
                If part.HasBackColor Then
                    Using b As New SolidBrush(part.BackColor)
                        g.FillRectangle(b, currentX, y, sz.Width, ItemHeight)
                    End Using
                End If
                Using b As New SolidBrush(part.ForeColor)
                    g.DrawString(part.Text, part.Font, b, currentX, y + (ItemHeight - part.Font.Height) / 2.0F, fmt)
                End Using
                currentX += sz.Width
            Next
        Else
            ' Cases B/C/D: desenare cu budget explicit și trunchiere posibilă
            currentX = DrawRichPartsInZone(g, leftParts, y, CSng(x), leftBudget, fmt)
        End If

        If Not hasSplit OrElse String.IsNullOrEmpty(rightText) Then Return

        ' ── 4. Desenare dreapta ──────────────────────────────────────────────────
        Dim rightParts As List(Of RichTextPart) = ParseRichText(rightText, defaultFont, defaultColor)
        Dim rightTotal As Single = 0
        For Each part In rightParts
            rightTotal += g.MeasureString(part.Text, part.Font, PointF.Empty, fmt).Width
        Next

        If caseA Then
            ' Case A: afișăm dreapta DOAR dacă încape complet după stânga
            Dim dynamicStart As Single = currentX + PADDING_SEPARATOR_GAP
            If dynamicStart + rightTotal > rightEdgeX Then Return ' nu încape — skip total
            ' Right-aligned în spațiul rămas
            DrawRichPartsSimple(g, rightParts, y, rightEdgeX - rightTotal, fmt)
        Else
            ' Cases B/C/D: right-aligned în zona rezervată sau trunchiat cu "..."
            If rightBudget <= 0 Then Return
            If rightTotal <= rightBudget Then
                ' Încape: right-aligned (nu mai la stânga de rightZoneStart)
                Dim raStart As Single = Math.Max(rightZoneStart, rightEdgeX - rightTotal)
                DrawRichPartsSimple(g, rightParts, y, raStart, fmt)
            Else
                ' Nu încape: trunchiăm de la rightZoneStart cu "..." la rightEdgeX
                DrawRichPartsInZone(g, rightParts, y, rightZoneStart, rightBudget, fmt)
            End If
        End If
    End Sub

    ' Parser simplu bazat pe Regex
    Friend Shared Function ParseRichText(rawText As String, baseFont As Font, baseColor As Color) As List(Of RichTextPart)
        Dim list As New List(Of RichTextPart)

        ' Regex care prinde tag-urile: <tag> sau </tag>
        ' Pattern explicat: < (/?)(b|i|u|color|back) (=([^>]+))? >
        Dim pattern As String = "<(/?)(b|i|u|color|back)(?:=([^>]+))?>"
        Dim matches As MatchCollection = Regex.Matches(rawText, pattern, RegexOptions.IgnoreCase)

        Dim lastIndex As Integer = 0

        ' Starea curentă
        Dim currentStyle As FontStyle = baseFont.Style
        Dim currentColor As Color = baseColor
        Dim currentBack As Color = Color.Transparent
        Dim hasBack As Boolean = False

        ' Stive pentru a reveni la starea anterioară (pentru nesting corect)
        Dim colorStack As New Stack(Of Color)
        Dim backStack As New Stack(Of Color)

        For Each m As Match In matches
            ' 1. Adăugăm textul dintre tag-uri (dacă există)
            If m.Index > lastIndex Then
                Dim txt As String = rawText.Substring(lastIndex, m.Index - lastIndex)
                list.Add(New RichTextPart With {
                    .Text = txt,
                    .Font = New Font(baseFont, currentStyle),
                    .ForeColor = currentColor,
                    .BackColor = currentBack,
                    .HasBackColor = hasBack
                })
            End If

            ' 2. Procesăm Tag-ul
            Dim isClosing As Boolean = (m.Groups(1).Value = "/")
            Dim tagName As String = m.Groups(2).Value.ToLower()
            Dim param As String = m.Groups(3).Value ' Valoarea de după = (ex: Red)

            If isClosing Then
                ' --- TAG DE ÎNCHIDERE ---
                Select Case tagName
                    Case "b" : currentStyle = currentStyle And Not FontStyle.Bold
                    Case "i" : currentStyle = currentStyle And Not FontStyle.Italic
                    Case "u" : currentStyle = currentStyle And Not FontStyle.Underline
                    Case "color"
                        If colorStack.Count > 0 Then currentColor = colorStack.Pop() Else currentColor = baseColor
                    Case "back"
                        If backStack.Count > 0 Then
                            currentBack = backStack.Pop()
                            hasBack = True
                        Else
                            currentBack = Color.Transparent
                            hasBack = False
                        End If
                End Select
            Else
                ' --- TAG DE DESCHIDERE ---
                Select Case tagName
                    Case "b" : currentStyle = currentStyle Or FontStyle.Bold
                    Case "i" : currentStyle = currentStyle Or FontStyle.Italic
                    Case "u" : currentStyle = currentStyle Or FontStyle.Underline
                    Case "color"
                        colorStack.Push(currentColor)
                        currentColor = ParseColor(param, baseColor)
                    Case "back"
                        If hasBack Then backStack.Push(currentBack)
                        currentBack = ParseColor(param, Color.Transparent)
                        hasBack = True
                End Select
            End If

            lastIndex = m.Index + m.Length
        Next

        ' 3. Adăugăm restul textului de după ultimul tag
        If lastIndex < rawText.Length Then
            list.Add(New RichTextPart With {
                .Text = rawText.Substring(lastIndex),
                .Font = New Font(baseFont, currentStyle),
                .ForeColor = currentColor,
                .BackColor = currentBack,
                .HasBackColor = hasBack
            })
        End If

        Return list
    End Function

    ' Desenează RichTextParts în zona [startX, startX+budget].
    ' Dacă textul nu încape → trunchiază cu "..." fix la capătul drept al zonei.
    ' Returnează X-ul după ultimul caracter desenat.
    Private Function DrawRichPartsInZone(g As Graphics, parts As List(Of RichTextPart),
                                      y As Integer, startX As Single, budget As Single,
                                      fmt As StringFormat) As Single
        If budget <= 0 OrElse parts.Count = 0 Then Return startX

        ' Calculăm lățimea totală
        Dim totalWidth As Single = 0
        For Each part In parts
            totalWidth += g.MeasureString(part.Text, part.Font, PointF.Empty, fmt).Width
        Next

        If totalWidth <= budget Then
            ' Încape integral — desenare simplă
            DrawRichPartsSimple(g, parts, y, startX, fmt)
            Return startX + totalWidth
        End If

        ' Nu încape — trunchiăm cu "..."
        Dim lastPart As RichTextPart = parts(parts.Count - 1)
        Dim ellipsisFont As Font = lastPart.Font
        Dim ellipsisColor As Color = lastPart.ForeColor
        Dim ellipsisWidth As Single = g.MeasureString("...", ellipsisFont, PointF.Empty, fmt).Width
        Dim spaceForText As Single = Math.Max(0, budget - ellipsisWidth)

        Dim rx As Single = startX
        Dim drawnSoFar As Single = 0
        Dim truncated As Boolean = False

        For Each part In parts
            If truncated Then Exit For

            Dim partWidth As Single = g.MeasureString(part.Text, part.Font, PointF.Empty, fmt).Width

            If drawnSoFar + partWidth <= spaceForText Then
                ' Întregul part încape
                If part.HasBackColor Then
                    Using b As New SolidBrush(part.BackColor)
                        g.FillRectangle(b, rx, y, partWidth, ItemHeight)
                    End Using
                End If
                Using b As New SolidBrush(part.ForeColor)
                    g.DrawString(part.Text, part.Font, b, rx, y + (ItemHeight - part.Font.Height) / 2.0F, fmt)
                End Using
                rx += partWidth
                drawnSoFar += partWidth
            Else
                ' Câte caractere mai încap
                Dim spaceLeft As Single = spaceForText - drawnSoFar
                Dim fitted As Integer = 0
                For c As Integer = 1 To part.Text.Length
                    If g.MeasureString(part.Text.AsSpan(0, c), part.Font, PointF.Empty, fmt).Width <= spaceLeft Then
                        fitted = c
                    Else
                        Exit For
                    End If
                Next
                If fitted > 0 Then
                    Dim part2 As String = part.Text.Substring(0, fitted)
                    Dim partialW As Single = g.MeasureString(part2, part.Font, PointF.Empty, fmt).Width
                    If part.HasBackColor Then
                        Using b As New SolidBrush(part.BackColor)
                            g.FillRectangle(b, rx, y, partialW, ItemHeight)
                        End Using
                    End If
                    Using b As New SolidBrush(part.ForeColor)
                        g.DrawString(part2, part.Font, b, rx, y + (ItemHeight - part.Font.Height) / 2.0F, fmt)
                    End Using
                End If
                truncated = True
            End If
        Next

        ' "..." fix la capătul drept al zonei
        Using b As New SolidBrush(ellipsisColor)
            g.DrawString("...", ellipsisFont, b,
                     startX + budget - ellipsisWidth,
                     y + (ItemHeight - ellipsisFont.Height) / 2.0F, fmt)
        End Using

        Return startX + budget
    End Function

    ' Desenează o listă de RichTextParts fără trunchiere, pornind de la startX.
    Private Sub DrawRichPartsSimple(g As Graphics, parts As List(Of RichTextPart),
                                 y As Integer, startX As Single, fmt As StringFormat)
        Dim rx As Single = startX
        For Each part In parts
            Dim sz As SizeF = g.MeasureString(part.Text, part.Font, PointF.Empty, fmt)
            If part.HasBackColor Then
                Using b As New SolidBrush(part.BackColor)
                    g.FillRectangle(b, rx, y, sz.Width, ItemHeight)
                End Using
            End If
            Using b As New SolidBrush(part.ForeColor)
                g.DrawString(part.Text, part.Font, b, rx, y + (ItemHeight - part.Font.Height) / 2.0F, fmt)
            End Using
            rx += sz.Width
        Next
    End Sub
    Private Function GetRoundedRect(rect As Rectangle, radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath()
        Dim diameter As Integer = radius * 2

        ' Evităm erorile dacă dreptunghiul e prea mic pentru rază
        If diameter > rect.Width Then diameter = rect.Width
        If diameter > rect.Height Then diameter = rect.Height

        Dim arc As New Rectangle(rect.X, rect.Y, diameter, diameter)

        path.AddArc(arc, 180, 90) ' Stânga sus
        arc.X = rect.Right - diameter
        path.AddArc(arc, 270, 90) ' Dreapta sus
        arc.Y = rect.Bottom - diameter
        path.AddArc(arc, 0, 90)   ' Dreapta jos
        arc.X = rect.X
        path.AddArc(arc, 90, 90)  ' Stânga jos
        path.CloseFigure()

        Return path
    End Function

    Friend Shared Function ParseColor(val As String, defaultColor As Color) As Color
        Try
            If String.IsNullOrEmpty(val) Then Return defaultColor
            If val.StartsWith("#"c) Then Return ColorTranslator.FromHtml(val)
            Return Color.FromName(val)
        Catch
            Return defaultColor
        End Try
    End Function

    ''' <summary>
    ''' Deseneaza randul de headere al coloanelor, fix sub header-ul principal si search bar.
    ''' Apelata din OnPaint DUPA items, astfel incat acopera orice bleeding.
    ''' </summary>
    Private Sub DrawColumnHeaders(g As Graphics)
        If Not _treeListViewEnabled OrElse Not _treeListView OrElse _columns.Count = 0 Then Return
        Try
            Dim headerOff As Integer = If(_headerVisible, _headerHeight, 0) +
                                       If(_isSearchMode, _searchBarHeight, 0)
            Dim hdrY As Integer = headerOff

            Try
                Using bgBrush As New SolidBrush(ControlPaint.Dark(Me.BackColor, 0.05F))
                    g.FillRectangle(bgBrush, 0, hdrY, Me.Width, COLUMN_HEADER_HEIGHT)
                End Using
            Catch
            End Try

            Try
                Using borderPen As New Pen(LineColor, 1)
                    g.DrawLine(borderPen, 0, hdrY + COLUMN_HEADER_HEIGHT - 1,
                                           Me.Width, hdrY + COLUMN_HEADER_HEIGHT - 1)
                End Using
            Catch
            End Try

            Dim cx As Integer = _captionColumnEndX
            Dim fmt As New StringFormat()
            fmt.LineAlignment = StringAlignment.Center
            fmt.Trimming = StringTrimming.EllipsisCharacter

            Try
                For i As Integer = 0 To _columns.Count - 1
                    Try
                        Dim cd = _columns(i)
                        Dim colRect As New Rectangle(cx, hdrY, cd.Width, COLUMN_HEADER_HEIGHT)

                        ' ── 1. BackColor per-coloana ─────────────────────────────────────────
                        ' LIPSEA COMPLET
                        If cd.HeaderBackColor <> Color.Empty Then
                            Using hdrBg As New SolidBrush(cd.HeaderBackColor)
                                g.FillRectangle(hdrBg, colRect)
                            End Using
                        End If

                        ' ── 2. Separator vertical ────────────────────────────────────────────
                        Using sepPen As New Pen(Color.FromArgb(COLUMN_SEPARATOR_COLOR_ALPHA, LineColor), 1)
                            g.DrawLine(sepPen, cx, hdrY, cx, hdrY + COLUMN_HEADER_HEIGHT)
                        End Using

                        ' ── 3. Aliniere efectiva ─────────────────────────────────────────────
                        ' Select Case cd.Align cu HorizontalAlignment era mort (suprascris imediat)
                        ' → eliminat complet
                        Dim effAlign As en_ColAlign = If(cd.HeaderAlign = en_ColAlign.ColAlign_Inherit, cd.Align, cd.HeaderAlign)
                        fmt.Alignment = ColAlignToStringAlign(effAlign)

                        ' ── 4. Font dinamic: Bold / Italic / Underline ───────────────────────
                        ' LIPSEA COMPLET — intotdeauna Me.Font
                        Dim hdrStyle As FontStyle = FontStyle.Regular
                        If cd.HeaderBold Then hdrStyle = hdrStyle Or FontStyle.Bold
                        If cd.HeaderItalic Then hdrStyle = hdrStyle Or FontStyle.Italic
                        If cd.HeaderUnderline Then hdrStyle = hdrStyle Or FontStyle.Underline
                        Dim hdrFont As Font = If(hdrStyle = FontStyle.Regular,
                                 Me.Font,
                                 New Font(Me.Font, hdrStyle))

                        ' ── 5. ForeColor per-coloana ─────────────────────────────────────────
                        ' LIPSEA COMPLET — intotdeauna Me.ForeColor
                        Dim hdrFore As Color = If(cd.HeaderForeColor <> Color.Empty,
                                  cd.HeaderForeColor,
                                  Me.ForeColor)

                        ' ── 6. Desenare text ─────────────────────────────────────────────────
                        Dim cellRect As New Rectangle(cx + 4, hdrY, cd.Width - 8, COLUMN_HEADER_HEIGHT)
                        Using fgBrush As New SolidBrush(hdrFore)
                            g.DrawString(cd.Header, hdrFont, fgBrush, cellRect, fmt)
                        End Using

                        ' ── 7. Dispose font creat dinamic ────────────────────────────────────
                        If hdrStyle <> FontStyle.Regular Then hdrFont.Dispose()

                        ' ── Filter indicator ● ───────────────────────────────────────────────
                        If _activeColFilters.ContainsKey(cd.Name) Then
                            Using indicBrush As New SolidBrush(Color.FromArgb(210, 55, 55))
                                g.FillEllipse(indicBrush,
                                              colRect.Right - 13,
                                              hdrY + (COLUMN_HEADER_HEIGHT - 8) \ 2,
                                              8, 8)
                            End Using
                        End If

                        cx += cd.Width
                    Catch
                        cx += If(i < _columns.Count, _columns(i).Width, 0)
                    End Try
                Next
            Finally
                fmt.Dispose()
            End Try
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Converteste en_ColAlign la StringAlignment pentru GDI+.
    ''' ColAlign_Inherit se trateaza ca Left (nu ar trebui sa ajunga aici).
    ''' </summary>
    Private Shared Function ColAlignToStringAlign(a As en_ColAlign) As StringAlignment
        Select Case a
            Case en_ColAlign.ColAlign_Center : Return StringAlignment.Center
            Case en_ColAlign.ColAlign_Right : Return StringAlignment.Far
            Case Else : Return StringAlignment.Near
        End Select
    End Function
End Class
