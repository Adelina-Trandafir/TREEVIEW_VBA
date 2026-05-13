Imports System.DirectoryServices
Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

Partial Public Class AdvancedTreeControl

    ' ── Filter state (inline search — no overlay) ────────────────────────
    Private _filterActive As Boolean = False
    Private _filterSet As New HashSet(Of TreeItem)()

    ' Search bar row state
    Private _searchBarHeight As Integer = 0
    Private _searchBarLabel As Label = Nothing
    Private _searchPlaceholderActive As Boolean = False

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

        ' ── Search bar row (below header, when search is active and no caption) ──
        If _isSearchMode AndAlso String.IsNullOrEmpty(_headerCaption) AndAlso _searchBarHeight > 0 Then
            Using bg As New SolidBrush(_headerBackColor)
                g.FillRectangle(bg, 0, _headerHeight, Me.Width, _searchBarHeight)
            End Using
            Using sep As New Pen(Color.FromArgb(60, _headerForeColor))
                g.DrawLine(sep, 0, _headerHeight + _searchBarHeight - 1,
                           Me.Width, _headerHeight + _searchBarHeight - 1)
            End Using
        End If
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

        ' Auto-open: SearchShow = True și nu există iconiță toggle
        If _searchShow AndAlso _headerSearchIcon Is Nothing Then
            OpenSearchMode()
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
            AddHandler _searchTextBox.GotFocus, AddressOf OnSearchTextBoxGotFocus
            AddHandler _searchTextBox.LostFocus, AddressOf OnSearchTextBoxLostFocus
            AddHandler _searchTextBox.KeyDown, AddressOf OnSearchTextBoxKeyDown
            Me.Controls.Add(_searchTextBox)
        End If
        UpdateSearchTextBoxFont()

        _searchTextBox.BackColor = Me.BackColor
        _searchTextBox.ForeColor = Me.ForeColor
        _searchTextBox.Text = ""

        If String.IsNullOrEmpty(_headerCaption) Then
            ' ── Ramura fără caption: label + textbox în rândul de sub header ──
            If _searchBarLabel Is Nothing Then
                _searchBarLabel = New Label() With {
                    .AutoSize = True,
                    .Text = _searchBarLabelText,
                    .ForeColor = If(_searchBarLabelForeColor <> Color.Empty, _searchBarLabelForeColor, _headerForeColor),
                    .BackColor = _headerBackColor
                }
                UpdateSearchBarLabelFont()
                Me.Controls.Add(_searchBarLabel)
            End If
            _searchBarHeight = Math.Max(ItemHeight + 8, Me.Font.Height + 10)
            If _searchBarHeight > _headerHeight Then _headerHeight = _searchBarHeight

            _searchBarLabel.Left = PADDING_TREE_START
            _searchBarLabel.Top = _headerHeight + (_searchBarHeight - _searchBarLabel.Height) \ 2
            _searchBarLabel.Visible = True
            _searchBarLabel.BringToFront()

            _searchTextBox.Top = _headerHeight + (_searchBarHeight - _searchTextBox.PreferredHeight) \ 2
            _searchTextBox.Left = _searchBarLabel.Right + 4
            _searchTextBox.Width = Math.Max(40, Me.Width - _searchBarLabel.Right - 4 - PADDING_TREE_END)
            _searchTextBox.Height = _searchTextBox.PreferredHeight
        Else
            ' ── Ramura cu caption: textbox în header row, dreapta, 1/4 lățime ──
            If _searchBarLabel IsNot Nothing Then _searchBarLabel.Visible = False
        End If

        PositionSearchTextBox()
        _searchTextBox.Visible = True
        _searchTextBox.BringToFront()

        _isSearchMode = True
        _searchResults.Clear()
        ApplySearchPlaceholder()
        Me.Invalidate()
        _searchTextBox.Focus()
    End Sub

    Friend Sub CloseSearchMode()
        ' Guard: persistent dacă SearchShow = True și nu există iconiță toggle
        If _searchShow AndAlso _headerSearchIcon Is Nothing Then Return

        If _searchTextBox IsNot Nothing Then _searchTextBox.Visible = False
        If _searchBarLabel IsNot Nothing Then _searchBarLabel.Visible = False
        _filterActive = False
        _filterSet.Clear()
        _isSearchMode = False
        _searchPlaceholderActive = False
        _searchResults.Clear()
        Me.Invalidate()
    End Sub

    Private Sub PositionSearchTextBox()
        If _searchTextBox Is Nothing Then Return
        Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)

        If String.IsNullOrEmpty(_headerCaption) Then
            ' ── Ramura 1: fără caption — textbox în header, lățime maximă disponibilă ──
            Dim left As Integer = PADDING_TREE_START
            If _headerLeftIcon IsNot Nothing Then left += _headerIconSize.Width + PADDING_ICON_GAP

            Dim right As Integer = Me.Width - PADDING_TREE_END - scrollW
            If _headerRightIcon IsNot Nothing Then right -= _headerIconSize.Width + PADDING_ICON_GAP
            If _headerSearchIcon IsNot Nothing Then right -= _headerIconSize.Width + PADDING_ICON_GAP

            _searchTextBox.Left = left - Me.AutoScrollPosition.X
            _searchTextBox.Top = (_headerHeight - _searchTextBox.PreferredHeight) \ 2
            _searchTextBox.Width = Math.Max(40, right - left)
            _searchTextBox.Height = _searchTextBox.PreferredHeight

            If _searchBarLabel IsNot Nothing Then _searchBarLabel.Visible = False
        Else
            ' ── Ramura 2: cu caption — textbox în header, dreapta, 1/4 lățime ──
            Dim total As Integer = Me.Width - PADDING_TREE_END - scrollW
            Dim tbWidth As Integer

            If _headerRightIcon IsNot Nothing Then
                Dim available As Integer = total - _headerIconSize.Width - PADDING_ICON_GAP
                tbWidth = available \ 3
            Else
                tbWidth = total \ 4
            End If

            Dim tbLeft As Integer = Me.Width - PADDING_TREE_END - scrollW - tbWidth
            If _headerRightIcon IsNot Nothing Then tbLeft -= _headerIconSize.Width + PADDING_ICON_GAP

            _searchTextBox.Left = tbLeft
            _searchTextBox.Top = (_headerHeight - _searchTextBox.PreferredHeight) \ 2
            _searchTextBox.Width = tbWidth
            _searchTextBox.Height = _searchTextBox.PreferredHeight

            If _searchBarLabel IsNot Nothing Then _searchBarLabel.Visible = False
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — TEXTBOX EVENTS
    ' ══════════════════════════════════════════════════════════════════

    Private Sub OnSearchTextChanged(sender As Object, e As EventArgs)
        If _searchPlaceholderActive Then Return
        _searchDebounceTimer.Stop()
        If _searchTextBox Is Nothing Then Return
        Dim txt = _searchTextBox.Text
        If txt.Length < 3 Then
            _filterActive = False
            _filterSet.Clear()
            _searchResults.Clear()
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
        _filterSet.Clear()

        If String.IsNullOrEmpty(searchText) OrElse searchText.Length < 3 Then
            _filterActive = False
            Me.Invalidate()
            Return
        End If

        ' 1. Găsește nodurile care se potrivesc direct
        Dim matchSet As New HashSet(Of TreeItem)()
        CollectMatchingNodes(Items, searchText, matchSet)

        ' 2. filterSet = matches + toți ancestorii lor
        For Each node In matchSet
            _filterSet.Add(node)
            Dim p = node.Parent
            While p IsNot Nothing
                _filterSet.Add(p)
                p = p.Parent
            End While
        Next

        ' 3. Populează _searchResults pentru date suplimentare
        BuildTreeSearchResults(searchText)

        ' 4. Activează filtrul
        _filterActive = (_filterSet.Count > 0)

        ' 5. Ridică evenimentul
        RaiseEvent SearchFinished(matchSet.ToList(), searchText)

        Me.Invalidate()
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
    ' SEARCH — PLACEHOLDER
    ' ══════════════════════════════════════════════════════════════════

    Friend Sub ApplySearchPlaceholder()
        If _searchTextBox Is Nothing OrElse String.IsNullOrEmpty(_searchDefaultText) Then Return
        If _searchTextBox.Focused Then Return
        _searchTextBox.Text = _searchDefaultText
        _searchTextBox.ForeColor = Color.Gray
        _searchPlaceholderActive = True
    End Sub

    Private Sub RemoveSearchPlaceholder()
        If Not _searchPlaceholderActive Then Return
        _searchTextBox.Text = ""
        _searchTextBox.ForeColor = Me.ForeColor
        _searchPlaceholderActive = False
    End Sub

    Private Sub OnSearchTextBoxGotFocus(sender As Object, e As EventArgs)
        RemoveSearchPlaceholder()
    End Sub

    Private Sub OnSearchTextBoxLostFocus(sender As Object, e As EventArgs)
        If _searchTextBox IsNot Nothing AndAlso _searchTextBox.Text = "" Then
            ApplySearchPlaceholder()
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — KEYBOARD NAVIGATION
    ' ══════════════════════════════════════════════════════════════════

    Private Sub OnSearchTextBoxKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode <> Keys.Down AndAlso e.KeyCode <> Keys.Up Then Return

        Dim visible = GetVisibleItems()
        If visible.Count = 0 Then Return

        pSelectedItem = If(e.KeyCode = Keys.Down, visible.First(), visible.Last())

        e.Handled = True
        Me.Focus()
        Dim itemY = GetItemY(pSelectedItem)
        If itemY >= 0 Then
            Me.AutoScrollPosition = New Point(0, itemY - _headerHeight - _searchBarHeight)
        End If
        Me.Invalidate()
    End Sub

End Class
