Imports System.Drawing.Drawing2D

Partial Public Class AdvancedTreeControl
    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        RecalculateItemHeight()
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        ' Setări pentru linii clare
        e.Graphics.SmoothingMode = SmoothingMode.None
        e.Graphics.PixelOffsetMode = PixelOffsetMode.Half

        Dim y As Integer = Me.AutoScrollPosition.Y
        Dim visibleItems = GetVisibleItems()

        ' Ajustăm scrollbar-ul virtual
        Me.AutoScrollMinSize = New Size(0, visibleItems.Count * ItemHeight)

        For Each it In visibleItems
            ' Desenăm doar ce este vizibil pe ecran (Clipping manual pentru performanță)
            If y + ItemHeight > 0 AndAlso y < Me.Height Then
                DrawItem(e.Graphics, it, y)
            End If
            y += ItemHeight
        Next
    End Sub

    ' ======================================================
    ' 7. LOGICA MOUSE & INTERACȚIUNE
    ' ======================================================
    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        Me.Focus()

        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then
            pSelectedItem = Nothing
            Me.Invalidate()
            Return
        End If

        ' --- 1. LOGICĂ ZONĂ MOARTĂ (Folosind constantele din AdvancedTreeControl.vb) ---
        If it IsNot Nothing Then
            ' Calculăm punctul de start exact ca în Painting.vb
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

            ' Considerăm zona activă începând de la linia expanderului/indentării
            ' Tot ce e în stânga alinierii nivelului este ignorat
            If e.X < gridLeft Then
                it = Nothing
            End If
        End If
        ' -------------------------------------------------------------------------------

        ' =================================================================
        ' 2. PRIORITATE ZERO: EXPANDER (+/-)
        ' =================================================================
        ' GetExpanderRect folosește deja PADDING_TREE_START intern
        Dim expRect = GetExpanderRect(it)

        ' Verificăm dacă click-ul e în zona expanderului (și dacă are copii)
        If expRect.Contains(e.Location) AndAlso (it.Children.Count > 0 OrElse it.LazyNode) Then

            ' A. Verificare Protecție Root (dacă e activă)
            If it.Level = 0 AndAlso Not _rootButton Then
                Return
            End If

            ' B. LOGICĂ LAZY LOAD (Interception)
            ' Verificăm dacă încercăm să deschidem un nod nescărcat
            If it.LazyNode AndAlso it.Children.Count = 0 Then
                RaiseEvent RequestLazyLoad(Me, it)
                Return ' STOP! Nu expandăm vizual (nu avem ce arăta încă)
            End If

            ' C. Acțiunea propriu-zisă (Standard)
            it.Expanded = Not it.Expanded
            Me.Invalidate()

            ' D. CRITIC: Oprim execuția aici! 
            Return
        End If

        ' =================================================================
        ' 3. PRIORITATE UNU: CHECKBOX (Dacă există)
        ' =================================================================
        If _checkBoxes Then
            ' GetCheckBoxRect folosește deja PADDING_TREE_START și PADDING_EXPANDER_GAP intern
            Dim chkRect = GetCheckBoxRect(it)

            If chkRect.Contains(e.Location) Then
                ' Toggle CheckState
                Dim newState As TreeCheckState = TreeCheckState.Checked
                If it.CheckState = TreeCheckState.Checked Then
                    newState = TreeCheckState.Unchecked
                End If

                ' Aplică logica recursivă
                SetNodeStateWithPropagation(it, newState)

                RaiseEvent NodeChecked(it)
                Me.Invalidate()

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

        If pSelectedItem Is pOldSelectedItem AndAlso e.Button = MouseButtons.Left Then Return
        If pSelectedItem IsNot it AndAlso e.Button = MouseButtons.Right Then Return

        If it IsNot Nothing Then
            _pendingClickItem = it
            _pendingMouseArgs = e
            ClickDelayTimer.Start()
        End If
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)
        MyBase.OnMouseDoubleClick(e)
        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        ' Dublu click oriunde pe rând face Toggle Expand
        If it.Children.Count > 0 OrElse it.LazyNode Then

            ' --- PROTECȚIE ROOT ---
            If it.Level = 0 AndAlso Not _rootButton Then
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
        RaiseEvent NodeDoubleClicked(it, e)
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
            ResetTooltip(it)
            Me.Invalidate()
        End If
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        pHoveredItem = Nothing
        pToolTip.Hide(Me)
        pTooltipTimer.Stop()
        Me.Invalidate()
    End Sub

End Class