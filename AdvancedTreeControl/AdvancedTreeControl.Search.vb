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

    Private _searchClearBtn As Label = Nothing
    Private Const CLEAR_BTN_WIDTH As Integer = 18

    ' ── Win32 CueBanner — placeholder nativ, fără race conditions ────────────
    Private Const EM_SETCUEBANNER As Integer = &H1501

    <System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    Private Shared Function SendMessage(hWnd As IntPtr, msg As Integer, wParam As IntPtr, lParam As String) As IntPtr
    End Function

    Private Sub SetSearchCueBanner()
        If _searchTextBox Is Nothing OrElse String.IsNullOrEmpty(_searchDefaultText) Then Return
        If _searchTextBox.IsHandleCreated Then
            ' wParam=0: banner dispare când textbox-ul primește focus (comportament standard)
            SendMessage(_searchTextBox.Handle, EM_SETCUEBANNER,
                    New IntPtr(0), _searchDefaultText)
        Else
            ' Handle nu e creat încă — aplicăm la HandleCreated
            AddHandler _searchTextBox.HandleCreated, AddressOf OnSearchTextBoxHandleCreated
        End If
    End Sub

    Private Sub OnSearchTextBoxHandleCreated(sender As Object, e As EventArgs)
        RemoveHandler _searchTextBox.HandleCreated, AddressOf OnSearchTextBoxHandleCreated
        If Not String.IsNullOrEmpty(_searchDefaultText) Then
            SendMessage(_searchTextBox.Handle, EM_SETCUEBANNER,
                    New IntPtr(0), _searchDefaultText)
        End If
    End Sub
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
        Dim scrollW As Integer = ScrollBarWidth 'If(_vScroll.Visible, _vScroll.Width, 0)
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

        ' Auto-open: SearchShow = True și nu există iconiță toggle
        If _searchShow AndAlso _headerSearchIcon Is Nothing Then
            OpenSearchMode()
        End If

        Me.Invalidate()
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH 
    ' ══════════════════════════════════════════════════════════════════
    Friend Sub DrawSearchBar(g As Graphics)
        Dim barTop As Integer = If(_headerVisible, _headerHeight, 0)

        ' Background cu culoarea proprie a benzii de search
        Using bg As New SolidBrush(_searchBackColor)
            g.FillRectangle(bg, 0, barTop, Me.Width, _searchBarHeight)
        End Using

        ' Separator inferior
        Using sep As New Pen(Color.FromArgb(80, Color.Black))
            g.DrawLine(sep, 0, barTop + _searchBarHeight - 1,
                   Me.Width, barTop + _searchBarHeight - 1)
        End Using
    End Sub

    Private Sub OpenSearchMode()
        If _searchTextBox Is Nothing Then
            _searchTextBox = New TextBox() With {
            .BorderStyle = BorderStyle.None,
            .Font = Me.Font,
            .TabStop = False,
            .TextAlign = HorizontalAlignment.Center
        }
            AddHandler _searchTextBox.TextChanged, AddressOf OnSearchTextChanged
            AddHandler _searchTextBox.KeyDown, AddressOf OnSearchTextBoxKeyDown
            Me.Controls.Add(_searchTextBox)
        End If
        UpdateSearchTextBoxFont()

        _searchTextBox.BackColor = If(_searchBoxBackColor = Color.Empty, Me.BackColor, _searchBoxBackColor)
        _searchTextBox.ForeColor = Me.ForeColor
        _searchTextBox.Text = ""

        ' ── Search bar este ÎNTOTDEAUNA o bandă separată ──────────────────
        _searchBarHeight = Math.Max(ItemHeight + 8, Me.Font.Height + 10)

        If Not String.IsNullOrEmpty(_searchBarLabelText) Then
            If _searchBarLabel Is Nothing Then
                _searchBarLabel = New Label() With {
                .AutoSize = True,
                .Text = _searchBarLabelText,
                .ForeColor = If(_searchBarLabelForeColor <> Color.Empty,
                                _searchBarLabelForeColor, _headerForeColor),
                .BackColor = _searchBackColor,
                .TabStop = False
            }
                UpdateSearchBarLabelFont()
                Me.Controls.Add(_searchBarLabel)
            Else
                _searchBarLabel.BackColor = _searchBackColor
            End If
            _searchBarLabel.Visible = True
            _searchBarLabel.BringToFront()
        End If

        ' ── Clear button (✕) — opțional, vizual în interiorul textbox-ului ────
        If _searchClearButton Then
            If _searchClearBtn Is Nothing Then
                _searchClearBtn = New Label() With {
            .Text = "✕",
            .AutoSize = False,
            .Width = CLEAR_BTN_WIDTH,
            .TextAlign = ContentAlignment.MiddleCenter,
            .Cursor = Cursors.Hand,
            .Visible = False,
            .TabStop = False
        }
                AddHandler _searchClearBtn.Click, AddressOf OnSearchClearBtnClick
                Me.Controls.Add(_searchClearBtn)
            End If
            Dim btnBack As Color = If(_searchBoxBackColor = Color.Empty, Me.BackColor, _searchBoxBackColor)
            _searchClearBtn.BackColor = btnBack
            _searchClearBtn.ForeColor = Me.ForeColor
            _searchClearBtn.Font = Me.Font
            _searchClearBtn.BringToFront()
        End If

        PositionSearchTextBox()
        _searchTextBox.Visible = True
        _searchTextBox.BringToFront()

        _isSearchMode = True
        _searchResults.Clear()
        _searchPlaceholderActive = False
        ApplySearchPlaceholder()
        Me.Invalidate()
        Me.Focus()
    End Sub

    Friend Sub CloseSearchMode()
        If _searchShow AndAlso _headerSearchIcon Is Nothing Then Return
        If _searchClearBtn IsNot Nothing Then _searchClearBtn.Visible = False
        If _searchTextBox IsNot Nothing Then _searchTextBox.Visible = False
        If _searchBarLabel IsNot Nothing Then _searchBarLabel.Visible = False
        _filterActive = False
        _filterSet.Clear()
        _isSearchMode = False
        _searchPlaceholderActive = False
        _searchResults.Clear()
        _searchBarHeight = 0    ' ← reset explicit — headerOff din OnPaint/GetItemY devine corect imediat
        Me.Invalidate()
    End Sub

    Private Sub PositionSearchTextBox()
        If _searchTextBox Is Nothing Then Return
        'Dim scrollW As Integer = ScrollBarWidth 'If(_vScroll.Visible, _vScroll.Width, 0)
        ' barTop vizual REAL = poziție fixă + compensare scroll
        Dim barTop As Integer = If(_headerVisible, _headerHeight, 0)
        Dim tbTop As Integer = barTop + (_searchBarHeight - _searchTextBox.PreferredHeight) \ 2

        ' Spațiu rezervat pentru ✕ — DOAR când butonul e vizibil
        Dim clearW As Integer = If(_searchClearButton AndAlso
                                _searchClearBtn IsNot Nothing AndAlso
                                _searchClearBtn.Visible, CLEAR_BTN_WIDTH, 0)

        Dim tbLeft As Integer
        Dim tbWidth As Integer

        If _searchBarLabel IsNot Nothing AndAlso _searchBarLabel.Visible Then
            _searchBarLabel.Left = PADDING_TREE_START
            _searchBarLabel.Top = barTop + (_searchBarHeight - _searchBarLabel.Height) \ 2
            tbLeft = _searchBarLabel.Right + 4
            tbWidth = Math.Max(40, Me.Width - tbLeft - PADDING_TREE_END - clearW)
        Else
            tbLeft = PADDING_TREE_START
            tbWidth = Math.Max(40, Me.Width - PADDING_TREE_START - PADDING_TREE_END - clearW)
        End If

        _searchTextBox.Left = tbLeft
        _searchTextBox.Top = tbTop
        _searchTextBox.Width = tbWidth
        _searchTextBox.Height = _searchTextBox.PreferredHeight

        ' ── Poziționare ✕ imediat la dreapta textbox-ului, aceeași înălțime ──
        If _searchClearButton AndAlso _searchClearBtn IsNot Nothing AndAlso _searchClearBtn.Visible Then
            _searchClearBtn.Left = _searchTextBox.Right
            _searchClearBtn.Top = _searchTextBox.Top
            _searchClearBtn.Height = _searchTextBox.Height
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════════════
    ' SEARCH — TEXTBOX EVENTS
    ' ══════════════════════════════════════════════════════════════════

    Private Sub OnSearchTextChanged(sender As Object, e As EventArgs)
        If _searchPlaceholderActive Then Return
        SearchDebounceTimer.Stop()
        If _searchTextBox Is Nothing Then Return
        Dim txt = _searchTextBox.Text
        UpdateClearBtnVisibility()      ' ← adăugat
        If txt.Length < 3 Then
            _filterActive = False
            _filterSet.Clear()
            _searchResults.Clear()
            Me.Invalidate()
        Else
            SearchDebounceTimer.Start()
        End If
    End Sub

    Private Sub OnSearchDebounceTimerTick(sender As Object, e As EventArgs) Handles SearchDebounceTimer.Tick
        SearchDebounceTimer.Stop()
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
            _vScroll.Value = 0                    ' ← reset
            Me.Invalidate()
            Return
        End If

        Dim matchSet As New HashSet(Of TreeItem)()
        CollectMatchingNodes(Items, searchText, matchSet)

        For Each node In matchSet
            _filterSet.Add(node)
            Dim p = node.Parent
            While p IsNot Nothing
                _filterSet.Add(p)
                p = p.Parent
            End While
        Next

        BuildTreeSearchResults(searchText)
        _filterActive = (_filterSet.Count > 0)
        RaiseEvent SearchFinished(matchSet.ToList(), searchText)

        _vScroll.Value = 0                        ' ← reset după filter nou
        Me.BeginInvoke(New Action(AddressOf RefreshScrollVisibility))
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
            Case En_Tree_SearchIn.SearchIn_Tag : toSearch = tagText
            Case En_Tree_SearchIn.SearchIn_Both : toSearch = plainCaption & " " & tagText
            Case Else : toSearch = plainCaption
        End Select

        Return If(_searchType = En_Tree_SearchType.SearchType_StartsWith,
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
        SetSearchCueBanner()
    End Sub

    Private Sub RemoveSearchPlaceholder()
        _searchPlaceholderActive = False
    End Sub

    Private Sub UpdateClearBtnVisibility()
        If Not _searchClearButton OrElse _searchClearBtn Is Nothing Then Return
        Dim shouldShow As Boolean = _searchTextBox IsNot Nothing AndAlso
                                Not _searchPlaceholderActive AndAlso
                                _searchTextBox.Text.Length > 0
        If _searchClearBtn.Visible = shouldShow Then Return
        _searchClearBtn.Visible = shouldShow

        ' Poziționare doar a butonului × — TextBox rămâne neatins
        If shouldShow AndAlso _searchTextBox IsNot Nothing Then
            _searchClearBtn.Left = _searchTextBox.Right - CLEAR_BTN_WIDTH
            _searchClearBtn.Top = _searchTextBox.Top
            _searchClearBtn.Height = _searchTextBox.Height
            _searchClearBtn.BackColor = _searchTextBox.BackColor
            _searchClearBtn.BringToFront()
        End If
    End Sub

    Private Sub OnSearchClearBtnClick(sender As Object, e As EventArgs)
        If _headerSearchIcon IsNot Nothing Then
            ' Se comportă identic cu click pe icona de search (toggle)
            CloseSearchMode()
        Else
            ' Curăță textul — OnSearchTextChanged resetează filtrul automat
            ' UpdateClearBtnVisibility ascunde × și relărgește textbox-ul
            If _searchTextBox IsNot Nothing Then
                _searchTextBox.Text = ""
                _searchTextBox.Focus()
            End If
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
