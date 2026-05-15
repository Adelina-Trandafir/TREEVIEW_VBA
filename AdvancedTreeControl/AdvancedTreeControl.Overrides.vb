Imports System.Drawing.Drawing2D

Partial Public Class AdvancedTreeControl
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        RecalculateItemHeight()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)
        e.Graphics.Clear(Me.BackColor)
        e.Graphics.SmoothingMode = SmoothingMode.None
        e.Graphics.PixelOffsetMode = PixelOffsetMode.Half

        Dim columnHdrH As Integer = If(_treeListView AndAlso _columns.Count > 0, COLUMN_HEADER_HEIGHT, 0)
        Dim headerOff As Integer = If(_headerVisible, _headerHeight, 0) +
                               If(_isSearchMode, _searchBarHeight, 0) +
                               columnHdrH

        ' ── 1. Items (cu clip) — se desenează PRIMII ──────────────────────
        Dim visibleItems = GetVisibleItems()
        Dim contentH As Integer = visibleItems.Count * ItemHeight + PADDING_TREE_TOP

        Dim oldClip = e.Graphics.Clip.Clone()
        e.Graphics.SetClip(New Rectangle(0, headerOff, Me.Width, Me.Height - headerOff))

        Dim y As Integer = -_vScroll.Value + PADDING_TREE_TOP + headerOff
        For Each it In visibleItems
            If y + ItemHeight > headerOff AndAlso y < Me.Height Then
                DrawItem(e.Graphics, it, y)
            End If
            y += ItemHeight
        Next

        e.Graphics.Clip = oldClip

        ' ── 2. Header + SearchBar desenate DUPĂ items — acoperă orice bleeding ──
        If _headerVisible Then DrawHeader(e.Graphics)
        If _isSearchMode Then DrawSearchBar(e.Graphics)
        ' ── 2b. Column headers (TreeListView) — deseneaza DUPA items, DUPA header ──
        If _treeListView Then DrawColumnHeaders(e.Graphics)

        ' ── 3. Scrollbar visibility (BeginInvoke — nu din interiorul OnPaint) ──
        Dim viewport As Integer = Math.Max(1, Me.Height - headerOff)
        Dim needsScroll As Boolean = contentH > viewport
        If _vScroll.Visible <> needsScroll Then
            Me.BeginInvoke(New Action(AddressOf RefreshScrollVisibility))
        ElseIf needsScroll Then
            ' Actualizează Maximum și Value fără să schimbe Visible
            _vScroll.LargeChange = viewport
            _vScroll.Maximum = Math.Max(viewport, contentH - 1)   ' ← FIX bug 3
            Dim maxVal As Integer = Math.Max(0, contentH - viewport)
            If _vScroll.Value > maxVal Then _vScroll.Value = maxVal
        End If

        ' ── 4. Disabled mask ──────────────────────────────────────────────
        If Not Me.Enabled Then
            Using brush As New SolidBrush(Color.FromArgb(120, Color.WhiteSmoke))
                e.Graphics.FillRectangle(brush, Me.ClientRectangle)
            End Using
        End If

        ' ── 5. Border ─────────────────────────────────────────────────────
        If Me.BorderColor <> Color.Transparent Then
            Using pen As New Pen(Me.BorderColor, 1)
                e.Graphics.DrawRectangle(pen, 1, 1, Me.Width - 1, Me.Height - 1)
            End Using
        End If
    End Sub

    ' ======================================================
    ' 7. LOGICA MOUSE & INTERACȚIUNE
    ' ======================================================
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        Me.Focus()

        ' ── Header area clicks ───────────────────────────────────────────────
        If _headerVisible AndAlso e.Y < _headerHeight Then
            If _headerSearchIconRect.Contains(e.Location) AndAlso _headerSearchIcon IsNot Nothing Then
                If _isSearchMode Then CloseSearchMode() Else OpenSearchMode()
            ElseIf _headerRightIconRect.Contains(e.Location) AndAlso _headerRightIcon IsNot Nothing Then
                RaiseEvent HeaderRightIconClicked(e)
            End If
            Return
        End If

        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then
            pSelectedItem = Nothing
            Me.Invalidate()
            Return
        End If

        Try
            If it IsNot Nothing Then
                it.LastClickedColumnIndex = -1
                it.LastClickedColumnName = ""
            End If
        Catch
        End Try

        ' --- 1. LOGICĂ ZONĂ MOARTĂ (Folosind constantele din AdvancedTreeControl.vb) ---
        If it IsNot Nothing Then
            ' Calculăm punctul de start exact ca în Painting.vb
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

            ' Considerăm zona activă începând de la linia expanderului/indentării
            ' Tot ce e în stânga alinierii nivelului → Toggle Expand dacă are copii/LazyNode
            If e.X < gridLeft Then
                If it.Children.Count > 0 OrElse it.LazyNode Then
                    ' A. Protecție Root
                    If it.Level = 0 AndAlso Not _RootExpander Then
                        Return
                    End If

                    ' B. Lazy Load (identic cu logica expanderului)
                    If it.LazyNode AndAlso it.Children.Count = 0 Then
                        Dim loader As New TreeItem With {
                    .Key = "LOADER_" & Guid.NewGuid().ToString(),
                    .Caption = "Loading...",
                    .Level = it.Level + 1,
                    .Parent = it,
                    .IsLoader = True
                }
                        it.Children.Add(loader)
                        it.Expanded = True
                        If Not LoadingTimer.Enabled Then LoadingTimer.Start()
                        Me.Invalidate()
                        RaiseEvent RequestLazyLoad(Me, it)
                        Return
                    End If

                    ' C. Toggle standard
                    it.Expanded = Not it.Expanded
                    Me.Invalidate()
                End If
                ' NICIODATĂ selecție în zona moartă
                Return
            End If
        End If
        ' -------------------------------------------------------------------------------

        ' --- Logica Nod Loader (Nu permitem interacțiuni cu nodurile loader) ---
        If it.IsLoader Then Return

        ' =================================================================
        ' 2. PRIORITATE ZERO: EXPANDER (+/-)
        ' =================================================================
        ' GetExpanderRect folosește deja PADDING_TREE_START intern
        Dim expRect = GetExpanderRect(it)

        ' Verificăm dacă click-ul e în zona expanderului (și dacă are copii)
        If expRect.Contains(e.Location) AndAlso (it.Children.Count > 0 OrElse it.LazyNode) Then

            ' A. Verificare Protecție Root (dacă e activă)
            If it.Level = 0 AndAlso Not _RootExpander Then
                Return
            End If

            If pSelectedItem IsNot it Then
                pSelectedItem = it
                RaiseEvent NodeMouseUp(it, e)
            End If

            ' B. LOGICĂ LAZY LOAD (Interception)
            If it.LazyNode AndAlso it.Children.Count = 0 Then

                ' 1. Creăm nodul temporar de încărcare
                Dim loader As New TreeItem With {
                    .Key = "LOADER_" & Guid.NewGuid().ToString(),
                    .Caption = "Loading...",
                    .Level = it.Level + 1,
                    .Parent = it,
                    .IsLoader = True
                }
                it.Children.Add(loader)

                ' 2. Expandăm vizual părintele imediat
                it.Expanded = True

                ' 3. Pornim animația dacă nu merge deja
                If Not LoadingTimer.Enabled Then LoadingTimer.Start()

                ' 4. Forțăm redesenarea pentru a apărea loaderul
                Me.Invalidate()

                ' 5. Trimitem cererea la VBA
                RaiseEvent RequestLazyLoad(Me, it)

                Return
            End If

            ' C. Acțiunea propriu-zisă (Standard) 
            it.Expanded = Not it.Expanded
            Me.Invalidate()

            ' D. CRITIC: Oprim execuția aici!  
            Return
        End If

        ' =================================================================
        ' 3. PRIORITATE UNU: CHECKBOX / RADIOBUTTON
        ' =================================================================
        If NodeHasCheckControl(it) Then
            Dim chkRect = GetCheckBoxRect(it)
            If chkRect.Contains(e.Location) Then

                If _radioButtonLevel >= 0 AndAlso it.Level = _radioButtonLevel Then
                    ' --- RADIO: deselectăm frații, ștergem checkboxurile copiilor nodeOff ---
                    Dim siblings As List(Of TreeItem) = If(it.Parent IsNot Nothing, it.Parent.Children, Me.Items)

                    ' Capturăm nodeOff ÎNAINTE
                    Dim nodeOff As TreeItem = Nothing
                    For Each sibling In siblings
                        If sibling.Level = _radioButtonLevel AndAlso sibling IsNot it AndAlso sibling.IsRadioSelected Then
                            nodeOff = sibling
                            Exit For
                        End If
                    Next

                    ' Ștergem checkboxurile copiilor lui nodeOff
                    If nodeOff IsNot Nothing Then
                        ClearChildrenCheckboxes(nodeOff)
                    End If

                    ' Deselectăm toți frații
                    For Each sibling In siblings
                        If sibling.Level = _radioButtonLevel Then
                            sibling.IsRadioSelected = False
                        End If
                    Next

                    ' Selectăm nodul curent
                    CheckChildrenRecursive(it)
                    it.IsRadioSelected = True

                    RaiseEvent NodeRadioSelected(it, nodeOff)
                    Me.Invalidate()

                Else
                    ' --- CHECKBOX STANDARD ---
                    Dim newState As TreeCheckState = If(it.CheckState = TreeCheckState.Checked,
                                                 TreeCheckState.Unchecked,
                                                 TreeCheckState.Checked)
                    SetNodeStateWithPropagation(it, newState)
                    RaiseEvent NodeChecked(it)
                    Me.Invalidate()
                End If

                pSelectedItem = it
                If pSelectedItem IsNot pOldSelectedItem Then RaiseEvent NodeMouseDown(it, e)
                Return
            End If
        End If

        ' =================================================================
        ' 4. PRIORITATE DOI: SELECȚIE RÂND (TEXT / ICON)
        ' =================================================================
        pSelectedItem = it
        RaiseEvent NodeMouseDown(it, e)

        ' =================================================================
        ' 5. PRIORITATE: RIGHT ICON CLICK
        ' =================================================================
        If it.RightIcon IsNot Nothing Then
            Dim scrollW As Integer = ScrollBarWidth 'If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
            ' Reconstituim dreptunghiul iconiței exact ca în Painting.vb
            Dim rIconRect As New Rectangle(Me.Width - RightIconSize.Width - 6 - scrollW,
                                           (it.Level * Indent) + Me.AutoScrollPosition.Y + (ItemHeight - RightIconSize.Height) \ 2, ' Aici trebuie calculat Y-ul vizual, nu logic
                                           RightIconSize.Width,
                                           RightIconSize.Height)

            ' Nota: Calculul Y de mai sus e complex pentru ca OnMouseDown nu ne da Y-ul desenat direct.
            ' Mai simplu: stim ca e in dreapta. Verificam doar X-ul.
            Dim minX As Integer = Me.Width - RightIconSize.Width - 6 - scrollW

            If e.X >= minX AndAlso e.X <= (minX + RightIconSize.Width) Then
                ' Aici ridici un eveniment special
                RaiseEvent RightIconClicked(it, e)
                'Return ' Oprim selecția rândului
            End If
        End If


        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)

        Dim it = HitTestItem(e.Location)

        ' --- Logică Zonă Moartă ---
        If it IsNot Nothing Then
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START
            If e.X < gridLeft Then
                it = Nothing
            End If
        End If
        ' --------------------------

        Dim expRect = GetExpanderRect(it)
        If expRect.Contains(e.Location) AndAlso it.Children.Count > 0 Then
            Return
        End If

        If Not Me.ReRaiseClickOnSameNode Then
            If pSelectedItem Is pOldSelectedItem AndAlso e.Button = MouseButtons.Left Then Return
        End If

        ' SCHIMBARE:
        ' Daca e activat RaiseLeftClickOnRightClick, se permite trigger-ul orice ar fi
        ' Daca NU e activat, atunci la Click Dreapta, daca nodul selectat NU e cel pe care am dat click,
        ' atunci se ridica trigger-ul left, orice ar fi, ca sa se selecteze nodul curent mai intai  
        'If Not RaiseLeftClickOnRightClick AndAlso e.Button = MouseButtons.Right Then
        '    If it IsNot Nothing AndAlso it IsNot pSelectedItem Then
        '        pSelectedItem = it
        '        RaiseEvent NodeMouseDown(it, New MouseEventArgs(MouseButtons.Left, 1, e.X, e.Y, 0))
        '    End If
        'End If

        If it IsNot Nothing Then
            ' ── TreeListView: detecteaza daca click-ul a cazut pe o coloana ─────────
            If _treeListView Then
                Try
                    Dim clickedColIdx As Integer = GetColumnAtX(e.X)
                    it.LastClickedColumnIndex = clickedColIdx
                    it.LastClickedColumnName = If(clickedColIdx >= 0 AndAlso clickedColIdx < _columns.Count,
                                                   _columns(clickedColIdx).Name, "")
                Catch
                    it.LastClickedColumnIndex = -1
                    it.LastClickedColumnName = ""
                End Try
            End If
            ' ─────────────────────────────────────────────────────────────────────────

            _pendingClickItem = it
            _pendingMouseArgs = e
            ClickDelayTimer.Start()
        End If
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)
        MyBase.OnMouseDoubleClick(e)
        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        ' --- 1. LOGICĂ ZONĂ MOARTĂ (Folosind constantele din AdvancedTreeControl.vb) ---
        If it IsNot Nothing Then
            ' Calculăm punctul de start exact ca în Painting.vb
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

            ' Considerăm zona activă începând de la linia expanderului/indentării
            ' Tot ce e în stânga alinierii nivelului este ignorat
            If e.X < gridLeft Then
                Return
            End If
        End If
        ' -------------------------------------------------------------------------------

        ' Dublu click oriunde pe rând face Toggle Expand
        If it.Children.Count > 0 OrElse it.LazyNode Then

            ' --- PROTECȚIE ROOT ---
            If it.Level = 0 AndAlso Not _RootExpander Then
                Return
            End If

            ' --- LOGICĂ LAZY LOAD (Interception și la Dublu Click) ---
            If it.LazyNode AndAlso it.Children.Count = 0 Then
                RaiseEvent RequestLazyLoad(Me, it)
                Return
            End If
            ' ---------------------------------------------------------

            it.Expanded = Not it.Expanded
            Me.Invalidate()
        End If

        ' Nu triggerează dublu click dacă e în zona moartă sau pe expander, dar dacă e dublu click pe text/icon, atunci da:
        If Not Me.IsPopupTree Then RaiseEvent NodeDoubleClicked(it, e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)

        Dim it = HitTestItem(e.Location)

        ' --- Logică Zonă Moartă ---
        If it IsNot Nothing Then
            ' Folosim constanta PADDING_TREE_START definită în AdvancedTreeControl.vb
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

            ' Opțional: Dacă vrei să ignori mouse-over chiar și pe indentare:
            ' Dim activeAreaStart As Integer = gridLeft ' Sau gridLeft + Indent

            If e.X < gridLeft Then
                it = Nothing
            End If
        End If
        ' --------------------------

        If it IsNot pHoveredItem Then
            pHoveredItem = it
            ResetTooltip(it, e.X)
            Me.Invalidate()
        End If

        ' Și stocăm întotdeauna X-ul curent (indiferent de hover change):
        _lastMouseX = e.X
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        pHoveredItem = Nothing
        HideAllTooltips()
        pTooltipTimer.Stop()
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnScroll(se As ScrollEventArgs)
        MyBase.OnScroll(se)
        If _isSearchMode Then PositionSearchTextBox()
    End Sub

    Protected Overrides Sub OnResize(e As EventArgs)
        MyBase.OnResize(e)
        _vScroll.Width = SystemInformation.VerticalScrollBarWidth
        _vScroll.Left = Math.Max(0, Me.Width - _vScroll.Width)
        _vScroll.Top = 0
        _vScroll.Height = Me.Height
        RefreshScrollVisibility()
        If _isSearchMode Then PositionSearchTextBox()
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        Dim headerOff As Integer = If(_headerVisible, _headerHeight, 0) +
                               If(_isSearchMode, _searchBarHeight, 0)
        Dim viewport As Integer = Math.Max(1, Me.Height - headerOff)
        Dim contentH As Integer = GetVisibleItems().Count * ItemHeight + PADDING_TREE_TOP
        If contentH <= viewport Then Return

        Dim lines As Integer = SystemInformation.MouseWheelScrollLines
        Dim delta As Integer = -(e.Delta \ 120) * lines * ItemHeight
        Dim maxVal As Integer = Math.Max(0, contentH - viewport)
        _vScroll.Value = Math.Max(0, Math.Min(_vScroll.Value + delta, maxVal))
        Me.Invalidate()
    End Sub

End Class