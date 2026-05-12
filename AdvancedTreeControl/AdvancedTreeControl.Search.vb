Imports System.DirectoryServices
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

Partial Public Class AdvancedTreeControl

    ' ══════════════════════════════════════════════════════════════════
    ' HEADER — DRAWING
    ' ══════════════════════════════════════════════════════════════════

    Friend Sub DrawHeader(g As Graphics)
        ' Background
        Using bg As New SolidBrush(_headerBackColor)
            g.FillRectangle(bg, 0, 0, Me.Width, _headerHeight)
        End Using

        Dim midY As Integer = _headerHeight \ 2

        ' ── Left icon ────────────────────────────────────────────────
        Dim x As Integer = PADDING_TREE_START
        If _headerLeftIcon IsNot Nothing Then
            Dim iy = midY - (_headerIconSize.Height \ 2)
            g.DrawImage(_headerLeftIcon, x, iy, _headerIconSize.Width, _headerIconSize.Height)
            x += _headerIconSize.Width + PADDING_ICON_GAP
        End If

        ' ── Right side: RightIcon then SearchIcon (built right-to-left) ──
        Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
        Dim rx As Integer = Me.Width - PADDING_TREE_END - scrollW

        _headerRightIconRect = Rectangle.Empty
        If _headerRightIcon IsNot Nothing Then
            rx -= _headerIconSize.Width
            Dim iy = midY - (_headerIconSize.Height \ 2)
            _headerRightIconRect = New Rectangle(rx, iy, _headerIconSize.Width, _headerIconSize.Height)
            g.DrawImage(_headerRightIcon, _headerRightIconRect)
            rx -= PADDING_ICON_GAP
        End If

        _headerSearchIconRect = Rectangle.Empty
        If _headerSearchIcon IsNot Nothing Then
            rx -= _headerIconSize.Width
            Dim iy = midY - (_headerIconSize.Height \ 2)
            _headerSearchIconRect = New Rectangle(rx, iy, _headerIconSize.Width, _headerIconSize.Height)
            g.DrawImage(_headerSearchIcon, _headerSearchIconRect)
            rx -= PADDING_ICON_GAP
        End If

        ' ── Caption (rich text, in remaining space) ───────────────────
        Dim captionRight As Integer = rx
        Dim availW As Integer = Math.Max(0, captionRight - x)
        If Not String.IsNullOrEmpty(_headerCaption) AndAlso availW > 0 Then
            Dim fmt = StringFormat.GenericTypographic
            fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces
            Dim parts = ParseRichText(_headerCaption, Me.TreeFont, _headerForeColor)
            Dim oldClip = g.Clip.Clone()
            g.SetClip(New Rectangle(x, 0, availW, _headerHeight))
            Dim cx As Single = x
            For Each part In parts
                Dim sz = g.MeasureString(part.Text, part.Font, PointF.Empty, fmt)
                If cx + sz.Width > x + availW Then Exit For
                If part.HasBackColor Then
                    Using b As New SolidBrush(part.BackColor)
                        g.FillRectangle(b, cx, 0, sz.Width, _headerHeight)
                    End Using
                End If
                Using b As New SolidBrush(part.ForeColor)
                    g.DrawString(part.Text, part.Font, b,
                                 cx, (_headerHeight - part.Font.Height) / 2.0F, fmt)
                End Using
                cx += sz.Width
            Next
            g.Clip = oldClip
        End If

        ' ── Bottom separator ─────────────────────────────────────────
        Using sep As New Pen(Color.FromArgb(60, _headerForeColor))
            g.DrawLine(sep, 0, _headerHeight - 1, Me.Width, _headerHeight - 1)
        End Using
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' HEADER — ICON KEY RESOLUTION (called from Tree.Builder after cache load)
    ' ══════════════════════════════════════════════════════════════════

    Public Sub ResolveHeaderIcons(cache As Dictionary(Of String, Image))
        Dim img As Image = Nothing
        If Not String.IsNullOrEmpty(_headerLeftIconKey) Then
            If cache.TryGetValue(_headerLeftIconKey, img) Then _headerLeftIcon = img
        End If
        If Not String.IsNullOrEmpty(_headerRightIconKey) Then
            If cache.TryGetValue(_headerRightIconKey, img) Then _headerRightIcon = img
        End If
        If Not String.IsNullOrEmpty(_headerSearchIconKey) Then
            If cache.TryGetValue(_headerSearchIconKey, img) Then _headerSearchIcon = img
        End If
        Me.Invalidate()
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — OPEN / CLOSE
    ' ══════════════════════════════════════════════════════════════════

    Private Sub OpenSearchMode()
        If _searchTextBox Is Nothing Then
            _searchTextBox = New TextBox() With {
                .BorderStyle = BorderStyle.None,
                .Font = Me.Font
            }
            AddHandler _searchTextBox.TextChanged, AddressOf OnSearchTextChanged
            Me.Controls.Add(_searchTextBox)
        End If

        _searchTextBox.BackColor = Me.BackColor
        _searchTextBox.ForeColor = Me.ForeColor
        _searchTextBox.Text = ""
        PositionSearchTextBox()
        _searchTextBox.Visible = True
        _searchTextBox.BringToFront()
        _searchTextBox.Focus()

        _isSearchMode = True
        _searchResultHoveredIdx = -1
        _searchResults.Clear()
        Me.Invalidate()
    End Sub

    Friend Sub CloseSearchMode()
        If _searchTextBox IsNot Nothing Then
            _searchTextBox.Visible = False
        End If
        _isSearchMode = False
        _searchResultHoveredIdx = -1
        _searchResults.Clear()
        Me.Invalidate()
    End Sub

    ' Compensates for AutoScrollPosition so TextBox stays fixed below header
    Private Sub PositionSearchTextBox()
        If _searchTextBox Is Nothing Then Return
        Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)

        ' Left boundary: after header left icon (if any)
        Dim left As Integer = PADDING_TREE_START
        If _headerLeftIcon IsNot Nothing Then left += _headerIconSize.Width + PADDING_ICON_GAP

        ' Right boundary: before header right icon (if any)
        Dim right As Integer = Me.Width - PADDING_TREE_END - scrollW
        If _headerRightIcon IsNot Nothing Then right -= _headerIconSize.Width + PADDING_ICON_GAP
        If _headerSearchIcon IsNot Nothing Then right -= _headerIconSize.Width + PADDING_ICON_GAP

        ' Compensate for AutoScrollPosition (ScrollableControl offsets children)
        _searchTextBox.Left = left - Me.AutoScrollPosition.X
        _searchTextBox.Top = (_headerHeight + 4) - Me.AutoScrollPosition.Y
        _searchTextBox.Width = Math.Max(40, right - left)
        _searchTextBox.Height = ItemHeight
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — TEXTBOX EVENTS
    ' ══════════════════════════════════════════════════════════════════

    Private Sub OnSearchTextChanged(sender As Object, e As EventArgs)
        _searchDebounceTimer.Stop()
        If _searchTextBox Is Nothing Then Return
        Dim txt = _searchTextBox.Text
        If txt.Length < 3 Then
            _searchResults.Clear()
            _searchResultHoveredIdx = -1
            Me.Invalidate()
        Else
            _searchDebounceTimer.Start()
        End If
    End Sub

    Private Sub OnSearchDebounceTimerTick(sender As Object, e As EventArgs) Handles _searchDebounceTimer.Tick
        _searchDebounceTimer.Stop()
        If _searchTextBox IsNot Nothing Then
            PerformSearch(_searchTextBox.Text)
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — CORE LOGIC
    ' ══════════════════════════════════════════════════════════════════

    Private Sub PerformSearch(searchText As String)
        _searchResults.Clear()
        If String.IsNullOrEmpty(searchText) OrElse searchText.Length < 3 Then
            Me.Invalidate()
            Return
        End If

        If _searchMode = en_Tree_SearchMode.SearchMode_Tree Then
            BuildTreeSearchResults(searchText)
        Else
            BuildListSearchResults(searchText)
        End If

        _searchResultHoveredIdx = -1
        Me.Invalidate()

        ' Ridică evenimentul cu nodurile găsite (non-dimmed = match real, nu ancestor)
        If searchText.Length >= 3 Then
            Dim matching = _searchResults.
                Where(Function(r) Not r.IsDimmed).
                Select(Function(r) r.Item).
                ToList()
            RaiseEvent SearchFinished(matching, searchText)
        End If
    End Sub

    Private Function MatchesSearch(it As TreeItem, searchText As String) As Boolean
        Dim lower As String = searchText.ToLowerInvariant()

        ' Strip mini-html tags for caption search
        Dim plainCaption As String = Regex.Replace(
            If(it.Caption, ""), "<[^>]+>", "", RegexOptions.IgnoreCase
        ).ToLowerInvariant()

        Dim tagText As String = If(it.Tag IsNot Nothing, it.Tag.ToString(), "").ToLowerInvariant()

        Dim toSearch As String
        Select Case _searchIn
            Case en_Tree_SearchIn.SearchIn_Tag : toSearch = tagText
            Case en_Tree_SearchIn.SearchIn_Both : toSearch = plainCaption & " " & tagText
            Case Else : toSearch = plainCaption
        End Select

        Return If(_searchType = en_Tree_SearchType.SearchType_StartsWith,
                  toSearch.StartsWith(lower),
                  toSearch.Contains(lower))
    End Function

    ' ── List mode ────────────────────────────────────────────────────
    Private Sub BuildListSearchResults(searchText As String)
        CollectListResultsRecursive(Items, searchText)
    End Sub

    Private Sub CollectListResultsRecursive(nodes As List(Of TreeItem), searchText As String)
        For Each it In nodes
            If MatchesSearch(it, searchText) Then
                _searchResults.Add(New SearchResultItem(it, False))
            End If
            CollectListResultsRecursive(it.Children, searchText)
        Next
    End Sub

    ' ── Tree mode ────────────────────────────────────────────────────
    Private Sub BuildTreeSearchResults(searchText As String)
        Dim matchSet As New HashSet(Of TreeItem)()
        CollectMatchingNodes(Items, searchText, matchSet)
        If matchSet.Count = 0 Then Return

        ' Collect all ancestors
        Dim ancestorSet As New HashSet(Of TreeItem)()
        For Each node In matchSet
            Dim p = node.Parent
            While p IsNot Nothing
                ancestorSet.Add(p)
                p = p.Parent
            End While
        Next

        ' DFS traversal, same order as tree rendering, keeping only relevant nodes
        BuildTreeResultsOrdered(Items, matchSet, ancestorSet)
    End Sub

    Private Sub CollectMatchingNodes(nodes As List(Of TreeItem),
                                     searchText As String,
                                     result As HashSet(Of TreeItem))
        For Each it In nodes
            If MatchesSearch(it, searchText) Then result.Add(it)
            CollectMatchingNodes(it.Children, searchText, result)
        Next
    End Sub

    Private Sub BuildTreeResultsOrdered(nodes As List(Of TreeItem),
                                        matchSet As HashSet(Of TreeItem),
                                        ancestorSet As HashSet(Of TreeItem))
        For Each it In nodes
            Dim isMatch = matchSet.Contains(it)
            Dim isAncestor = ancestorSet.Contains(it)
            If Not isMatch AndAlso Not isAncestor Then Continue For

            ' Dimmed = ancestor-only (not itself a match)
            _searchResults.Add(New SearchResultItem(it, isAncestor AndAlso Not isMatch))

            ' Always recurse into children of ancestors (forced expand)
            If isAncestor Then
                BuildTreeResultsOrdered(it.Children, matchSet, ancestorSet)
            End If
        Next
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — OVERLAY PAINTING
    ' ══════════════════════════════════════════════════════════════════

    ' Constants for overlay layout (derived, not configurable — keeps surface small)
    Private ReadOnly Property SearchResultsTop As Integer
        Get
            Return _headerHeight + ItemHeight + 8
        End Get
    End Property

    Friend Sub DrawSearchOverlay(g As Graphics)
        Dim overlayRect As New Rectangle(0, _headerHeight, Me.Width, _searchDropdownHeight)

        ' ── Background ───────────────────────────────────────────────
        Using bg As New SolidBrush(Me.BackColor)
            g.FillRectangle(bg, overlayRect)
        End Using

        ' ── Input row border (around TextBox) ────────────────────────
        If _searchTextBox IsNot Nothing AndAlso _searchTextBox.Visible Then
            Dim tbr As New Rectangle(
                _searchTextBox.Left + Me.AutoScrollPosition.X - 1,
                _searchTextBox.Top + Me.AutoScrollPosition.Y - 1,
                _searchTextBox.Width + 2,
                _searchTextBox.Height + 2)
            Using pen As New Pen(Color.FromArgb(150, LineColor))
                g.DrawRectangle(pen, tbr)
            End Using
        End If

        ' ── Separator between input and results ───────────────────────
        Dim sepY = SearchResultsTop - 2
        Using sep As New Pen(Color.FromArgb(60, LineColor))
            g.DrawLine(sep, PADDING_TREE_START, sepY, Me.Width - PADDING_TREE_END, sepY)
        End Using

        ' ── Bottom border of overlay ──────────────────────────────────
        Dim bottomY = _headerHeight + _searchDropdownHeight - 1
        Using sep As New Pen(Color.FromArgb(80, LineColor))
            g.DrawLine(sep, 0, bottomY, Me.Width, bottomY)
        End Using

        ' ── No results / hint text ────────────────────────────────────
        If _searchResults.Count = 0 Then
            Dim hint As String = ""
            If _searchTextBox IsNot Nothing Then
                If _searchTextBox.Text.Length > 0 AndAlso _searchTextBox.Text.Length < 3 Then
                    hint = "Minim 3 caractere..."
                ElseIf _searchTextBox.Text.Length >= 3 Then
                    hint = "Niciun rezultat."
                End If
            End If
            If Not String.IsNullOrEmpty(hint) Then
                Using hintBrush As New SolidBrush(Color.Gray)
                    g.DrawString(hint, Me.Font, hintBrush, PADDING_TREE_START, SearchResultsTop + 4)
                End Using
            End If
            Return
        End If

        ' ── Results ───────────────────────────────────────────────────
        Dim maxBottom As Integer = _headerHeight + _searchDropdownHeight
        Dim oldClip = g.Clip.Clone()
        g.SetClip(New Rectangle(0, SearchResultsTop, Me.Width, maxBottom - SearchResultsTop))

        Dim isListMode As Boolean = (_searchMode = en_Tree_SearchMode.SearchMode_List)
        Dim y As Integer = SearchResultsTop
        For idx As Integer = 0 To _searchResults.Count - 1
            If y + ItemHeight > maxBottom Then Exit For
            Dim r = _searchResults(idx)
            DrawSearchResultRow(g, r.Item, y,
                                isHovered:=(idx = _searchResultHoveredIdx),
                                isDimmed:=r.IsDimmed,
                                flatMode:=isListMode)
            y += ItemHeight
        Next

        g.Clip = oldClip
    End Sub

    Private Sub DrawSearchResultRow(g As Graphics, it As TreeItem, y As Integer,
                                    isHovered As Boolean, isDimmed As Boolean,
                                    flatMode As Boolean)
        ' Indentation
        Dim level As Integer = If(flatMode, 0, it.Level)
        Dim gridLeft As Integer = level * Indent + PADDING_TREE_START
        Dim x As Integer = If(level = 0 AndAlso Not _RootExpander,
                              gridLeft,
                              gridLeft + Indent + PADDING_EXPANDER_GAP)

        ' ── Hover background (only for selectable = non-dimmed) ──────
        If isHovered AndAlso Not isDimmed Then
            Dim hw = Math.Max(0, Me.ClientSize.Width - gridLeft - PADDING_TREE_END)
            Using hb As New SolidBrush(HoverBackColor)
                g.FillRectangle(hb, gridLeft, y, hw, ItemHeight)
            End Using
        End If

        ' ── Left icon ────────────────────────────────────────────────
        If _hasNodeIcons Then
            Dim icon As Image = If(it.Expanded, it.LeftIconOpen, it.LeftIconClosed)
            If icon IsNot Nothing Then
                Dim iy = y + (ItemHeight - LeftIconSize.Height) \ 2
                If isDimmed Then
                    ' Draw dimmed icon using semi-transparent ImageAttributes
                    Dim ia As New System.Drawing.Imaging.ImageAttributes()
                    Dim cm As New System.Drawing.Imaging.ColorMatrix() With {.Matrix33 = 0.35F}
                    ia.SetColorMatrix(cm)
                    g.DrawImage(icon, New Rectangle(x, iy, LeftIconSize.Width, LeftIconSize.Height),
                                0, 0, icon.Width, icon.Height,
                                GraphicsUnit.Pixel, ia)
                Else
                    g.DrawImage(icon, x, iy, LeftIconSize.Width, LeftIconSize.Height)
                End If
                x += LeftIconSize.Width + PADDING_ICON_GAP
            End If
        End If

        ' ── Text boundaries ──────────────────────────────────────────
        Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
        Dim maxRightX As Integer = Me.Width - scrollW - PADDING_TREE_END
        If it.RightIcon IsNot Nothing AndAlso Not isDimmed Then
            maxRightX -= RightIconSize.Width + PADDING_RIGHT_ICON_GAP
        End If
        Dim availW As Integer = Math.Max(0, maxRightX - x)

        ' ── Caption ──────────────────────────────────────────────────
        Dim baseColor As Color = If(isDimmed,
            Color.FromArgb(150, 150, 150),
            If(it.NodeForeColor <> Color.Empty, it.NodeForeColor,
               If(Me.ForeColor <> Color.Empty, Me.ForeColor, Color.Black)))

        Dim style As FontStyle = Me.TreeFont.Style
        If Not isDimmed Then
            If it.Bold Then style = style Or FontStyle.Bold
            If it.Italic Then style = style Or FontStyle.Italic
        End If
        Dim nodeFont As Font = If(style <> Me.Font.Style, New Font(Me.Font, style), Me.TreeFont)

        Dim oldClip = g.Clip.Clone()
        g.SetClip(New Rectangle(x, y, availW, ItemHeight))
        DrawRichText(g, it.Caption, x, y, nodeFont, baseColor, availW)
        g.Clip = oldClip

        ' ── Right icon (non-dimmed only) ─────────────────────────────
        If it.RightIcon IsNot Nothing AndAlso Not isDimmed Then
            Dim rx As Integer = Me.Width - RightIconSize.Width - PADDING_RIGHT_ICON_GAP - PADDING_TREE_END - scrollW
            Dim ry As Integer = y + (ItemHeight - RightIconSize.Height) \ 2
            g.DrawImage(it.RightIcon, rx, ry, RightIconSize.Width, RightIconSize.Height)
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — OVERLAY HIT TESTING
    ' ══════════════════════════════════════════════════════════════════

    ' Returns the result index under the point, or -1 if not in results area
    Private Function OverlayResultIndexAt(p As Point) As Integer
        If Not _isSearchMode Then Return -1
        If p.Y < SearchResultsTop Then Return -1
        If p.Y >= _headerHeight + _searchDropdownHeight Then Return -1
        Dim idx = (p.Y - SearchResultsTop) \ ItemHeight
        Return If(idx >= 0 AndAlso idx < _searchResults.Count, idx, -1)
    End Function

End Class