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

        If it IsNot Nothing Then
            Dim chkWidth As Integer = If(_checkBoxes, _checkBoxSize, 0)
            Dim leftMargin As Integer = Indent + 5 - chkWidth
            Dim xStart As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin
            If e.X < xStart Then
                it = Nothing
            End If
        End If

        ' =================================================================
        ' 1. PRIORITATE ZERO: EXPANDER
        ' Dacă am dat click aici, facem Toggle și IEȘIM (Return).
        ' Nu vrem să se selecteze rândul.
        ' =================================================================
        Dim expRect = GetExpanderRect(it)

        ' Verificăm dacă click-ul e în zona expanderului (și dacă are copii)
        If expRect.Contains(e.Location) AndAlso it.Children.Count > 0 Then

            ' A. Verificare Protecție Root (dacă e activă)
            ' Dacă e root și nu are expander vizual, ignorăm click-ul AICI.
            ' Dar punem Return ca să NU ajungă la selecție (zona fiind goală/invizibilă).
            If it.Level = 0 AndAlso Not _rootButton Then
                Return
            End If

            ' B. Acțiunea propriu-zisă
            it.Expanded = Not it.Expanded
            Me.Invalidate()

            ' C. CRITIC: Oprim execuția aici! 
            ' Astfel nu se execută codul de mai jos (selecție/checkbox).
            Return
        End If

        ' =================================================================
        ' 2. PRIORITATE UNU: CHECKBOX (Dacă există)
        ' =================================================================
        If _checkBoxes Then
            Dim chkRect = GetCheckBoxRect(it)
            If chkRect.Contains(e.Location) Then

                ' Toggle CheckState (Unchecked <-> Checked)
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
        ' 3. PRIORITATE DOI: SELECȚIE RÂND (TEXT / ICON)
        ' Ajungem aici DOAR dacă nu s-a dat click pe Expander sau Checkbox
        ' =================================================================
        'If pSelectedItem IsNot it Then

        pSelectedItem = it
        RaiseEvent NodeMouseDown(it, e)
        'End If
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)

        Dim it = HitTestItem(e.Location)

        If it IsNot Nothing Then
            Dim chkWidth As Integer = If(_checkBoxes, _checkBoxSize, 0)
            Dim leftMargin As Integer = Indent + 5 - chkWidth
            Dim xStart As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin
            If e.X < xStart Then
                it = Nothing
            End If
        End If

        Dim expRect = GetExpanderRect(it)
        If expRect.Contains(e.Location) AndAlso it.Children.Count > 0 Then
            Return
        End If

        If pSelectedItem Is pOldSelectedItem AndAlso e.Button = MouseButtons.Left Then Return
        If pSelectedItem IsNot it AndAlso e.Button = MouseButtons.Right Then Return

        If it IsNot Nothing Then
            ' NU trimitem evenimentul imediat. Îl salvăm pentru mai târziu.
            _pendingClickItem = it
            _pendingMouseArgs = e

            ' Pornim cronometrul. Dacă utilizatorul dă al doilea click repede, 
            ' acest timer va fi oprit în OnMouseDoubleClick înainte să apuce să ticăie.
            ClickDelayTimer.Start()
        End If
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)
        MyBase.OnMouseDoubleClick(e)
        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        ' Dublu click oriunde pe rând face Toggle Expand
        If it.Children.Count > 0 Then

            ' --- PROTECȚIE ROOT ---
            ' Dacă e root și nu are expander, NU permitem collapse/expand
            If it.Level = 0 AndAlso Not _rootButton Then
                Return
            End If
            ' ----------------------

            it.Expanded = Not it.Expanded
            Me.Invalidate()
        End If
        RaiseEvent NodeDoubleClicked(it, e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)

        ' 1. Aflăm nodul de sub cursor (pe axa Y)
        Dim it = HitTestItem(e.Location)

        ' 2. --- LOGICĂ NOUĂ: ZONA MOARTĂ (LINII) ---
        If it IsNot Nothing Then
            ' Recalculăm marginea exact cum am făcut la DrawItem
            Dim leftMargin As Integer = Indent + 5

            ' Calculăm unde începe zona activă a acestui nod specific
            ' (Tot ce e la stânga lui xStart sunt linii ierarhice sau spațiu gol)
            Dim xStart As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin

            ' Dacă checkbox-urile sunt active, zona activă începe chiar de la checkbox? 
            ' Sau vrei ca nici checkbox-ul să nu se activeze dacă ești prea în stânga?
            ' De obicei, xStart definit mai sus e linia de unde începe Checkbox-ul sau Expander-ul.

            ' Dacă mouse-ul este în stânga indentării acestui nivel -> IGNORĂM
            If e.X < xStart Then
                it = Nothing
            End If
        End If
        ' -------------------------------------------

        ' 3. Gestionarea Hover-ului (Standard)
        If it IsNot pHoveredItem Then
            pHoveredItem = it
            ResetTooltip(it) ' Resetare tooltip
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
