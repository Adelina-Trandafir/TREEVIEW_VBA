Partial Public Class AdvancedTreeControl

    Private _keyNavPending As Boolean = False

    ' ── IsInputKey ─────────────────────────────────────────────────────────────
    ' Fără asta, Form-ul consumă Up/Down/Left/Right/PgUp/PgDn/Home/End ca
    ' "dialog keys" înainte ca controlul să le vadă vreodată.
    Protected Overrides Function IsInputKey(keyData As Keys) As Boolean
        Select Case keyData And Keys.KeyCode
            Case Keys.Up, Keys.Down, Keys.Left, Keys.Right,
                 Keys.PageUp, Keys.PageDown, Keys.Home, Keys.End
                Return True
        End Select
        Return MyBase.IsInputKey(keyData)
    End Function

    ' ── OnKeyDown ───────────────────────────────────────────────────────────────
    ' Navigare vizuală la fiecare repeat al tastei.
    ' _keyNavPending = True DOAR dacă selecția chiar s-a schimbat.
    Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
        MyBase.OnKeyDown(e)
        If Not Me.Enabled Then Return

        Dim visible As List(Of TreeItem) = GetVisibleItems()
        If visible.Count = 0 Then Return

        Dim prevItem As TreeItem = pSelectedItem
        Dim currentIdx As Integer = If(pSelectedItem IsNot Nothing,
                                       visible.IndexOf(pSelectedItem), -1)
        Dim handled As Boolean = False

        Select Case e.KeyCode

            Case Keys.Up
                handled = True
                pSelectedItem = visible(Math.Max(0, If(currentIdx < 0, 0, currentIdx - 1)))

            Case Keys.Down
                handled = True
                pSelectedItem = visible(Math.Min(visible.Count - 1,
                                                  If(currentIdx < 0, 0, currentIdx + 1)))

            Case Keys.Home
                handled = True
                pSelectedItem = visible(0)

            Case Keys.End
                handled = True
                pSelectedItem = visible(visible.Count - 1)

            Case Keys.PageUp
                handled = True
                Dim pgUpSize As Integer = Math.Max(1, Me.Height \ ItemHeight)
                pSelectedItem = visible(Math.Max(0,
                                                  If(currentIdx < 0, 0, currentIdx - pgUpSize)))

            Case Keys.PageDown
                handled = True
                Dim pgDnSize As Integer = Math.Max(1, Me.Height \ ItemHeight)
                pSelectedItem = visible(Math.Min(visible.Count - 1,
                                                  If(currentIdx < 0, 0, currentIdx + pgDnSize)))

            Case Keys.Left
                handled = True
                If pSelectedItem IsNot Nothing Then
                    If pSelectedItem.Expanded AndAlso pSelectedItem.Children.Count > 0 Then
                        ' Collapse — respectăm RootExpander
                        If pSelectedItem.Level > 0 OrElse _RootExpander Then
                            pSelectedItem.Expanded = False
                        End If
                        ' selecția rămâne pe același nod
                    ElseIf pSelectedItem.Parent IsNot Nothing Then
                        pSelectedItem = pSelectedItem.Parent
                    End If
                End If

            Case Keys.Right
                handled = True
                If pSelectedItem IsNot Nothing Then
                    Dim canExpand As Boolean =
                        (pSelectedItem.Children.Count > 0 OrElse pSelectedItem.LazyNode) AndAlso
                        (pSelectedItem.Level > 0 OrElse _RootExpander)

                    If Not pSelectedItem.Expanded AndAlso canExpand Then
                        ' Expand (+ LazyLoad dacă e cazul)
                        If pSelectedItem.LazyNode AndAlso pSelectedItem.Children.Count = 0 Then
                            Dim loader As New TreeItem With {
                                .Key = "LOADER_" & Guid.NewGuid().ToString(),
                                .Caption = "Loading...",
                                .Level = pSelectedItem.Level + 1,
                                .Parent = pSelectedItem,
                                .IsLoader = True
                            }
                            pSelectedItem.Children.Add(loader)
                            pSelectedItem.Expanded = True
                            If Not loadingTimer.Enabled Then loadingTimer.Start()
                            RaiseEvent RequestLazyLoad(Me, pSelectedItem)
                        Else
                            pSelectedItem.Expanded = True
                        End If
                        ' selecția rămâne pe nodul expandat

                    ElseIf pSelectedItem.Expanded AndAlso pSelectedItem.Children.Count > 0 Then
                        ' Deja expandat → mergi la primul copil non-loader
                        Dim firstChild As TreeItem = pSelectedItem.Children(0)
                        If Not firstChild.IsLoader Then
                            pSelectedItem = firstChild
                        End If
                    End If
                End If

        End Select

        If handled Then
            e.Handled = True
            e.SuppressKeyPress = True

            If pSelectedItem IsNot Nothing Then
                EnsureNodeVisible(pSelectedItem)
            End If
            Me.Invalidate()

            ' _keyNavPending = True doar dacă selecția s-a schimbat efectiv
            If Not Object.ReferenceEquals(pSelectedItem, prevItem) Then
                _keyNavPending = True
            End If
        End If
    End Sub

    ' ── OnKeyUp ─────────────────────────────────────────────────────────────────
    ' Trimite evenimentul la VBA O SINGURĂ DATĂ, la ridicarea tastei.
    ' Indiferent cât timp s-a ținut apăsată, VBA primește un singur CLICK.
    Protected Overrides Sub OnKeyUp(e As KeyEventArgs)
        MyBase.OnKeyUp(e)

        Select Case e.KeyCode
            Case Keys.Up, Keys.Down, Keys.Left, Keys.Right,
             Keys.PageUp, Keys.PageDown, Keys.Home, Keys.End
                ' Navigare vizuală — NU ridicăm NodeMouseUp; utilizatorul confirmă cu Enter
                If _keyNavPending Then
                    _keyNavPending = False
                    Me.Invalidate()
                End If

            Case Keys.Enter
                ' Enter = confirmare selecție → un singur NodeMouseUp
                If pSelectedItem IsNot Nothing Then
                    RaiseEvent NodeMouseUp(pSelectedItem,
                    New MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0))
                    pOldSelectedItem = pSelectedItem
                End If
        End Select
    End Sub

    ' ── EnsureNodeVisible ────────────────────────────────────────────────────────
    ' Scroll minim necesar ca nodul să fie vizibil — nu sare la mijloc,
    ' mișcă doar dacă e în afara viewport-ului.
    Private Sub EnsureNodeVisible(node As TreeItem)
        Dim visible As List(Of TreeItem) = GetVisibleItems()
        Dim idx As Integer = visible.IndexOf(node)
        If idx < 0 Then Return

        Dim headerOff As Integer = If(_headerVisible, _headerHeight, 0) +
                                   If(_isSearchMode, _searchBarHeight, 0)
        Dim viewport As Integer = Math.Max(1, Me.Height - headerOff)
        Dim scrollY As Integer = _vScroll.Value
        Dim nodeTop As Integer = PADDING_TREE_TOP + idx * ItemHeight
        Dim nodeBot As Integer = nodeTop + ItemHeight

        If nodeTop < scrollY Then
            _vScroll.Value = nodeTop
        ElseIf nodeBot > scrollY + viewport Then
            _vScroll.Value = Math.Max(0, nodeBot - viewport)
        End If
        ' Dacă e deja în viewport → nu mișcăm nimic
    End Sub
End Class