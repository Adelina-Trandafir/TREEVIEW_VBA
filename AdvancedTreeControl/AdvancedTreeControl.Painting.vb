Imports System.Drawing.Drawing2D

Partial Public Class AdvancedTreeControl
    Private Sub DrawItem(g As Graphics, it As TreeItem, y As Integer)
        ' Forțăm expandarea root-ului dacă nu are expander permis
        If it.Level = 0 AndAlso Not _rootButton AndAlso Not it.Expanded Then
            it.Expanded = True
        End If

        ' 1. Punctul de start al grilei pentru nivelul curent (linia din stânga a nivelului)
        Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

        ' 2. Expander-ul este centrat în coloana de indentare
        Dim expanderCenterX As Integer = gridLeft + (Indent \ 2)
        Dim midY As Integer = y + (ItemHeight \ 2)
        Dim expanderRect As New Rectangle(expanderCenterX - (ExpanderSize \ 2), midY - (ExpanderSize \ 2), ExpanderSize, ExpanderSize)

        ' 3. Conținutul (Checkbox/Text) începe DUPĂ indentare + SPAȚIUL SUPLIMENTAR (PADDING_EXPANDER_GAP)
        ' Aici se aplică distanțarea cerută
        Dim xBase As Integer = gridLeft + Indent + PADDING_EXPANDER_GAP

        ' -- [PASUL 1] SELECȚIE & HOVER (FULL ROW) --
        ' Calculăm selecția să înceapă de la limita vizuală a nivelului
        Dim selStartX As Integer = gridLeft + ExpanderSize * 2 + 2 ' +2 ca să nu acoperim linia punctată a părintelui
        Dim selWidth As Integer = Me.ClientSize.Width - selStartX
        If selWidth < 0 Then selWidth = 0

        Dim fullRowRect As New Rectangle(selStartX, y, selWidth, ItemHeight)

        If it Is pSelectedItem Then
            Using brush As New SolidBrush(SelectedBackColor)
                g.FillRectangle(brush, fullRowRect)
            End Using
            Using pen As New Pen(SelectedBorderColor)
                Dim borderRect As New Rectangle(selStartX, y, selWidth - 1, ItemHeight - 1)
                g.DrawRectangle(pen, borderRect)
            End Using
        ElseIf it Is pHoveredItem Then
            Using brush As New SolidBrush(HoverBackColor)
                g.FillRectangle(brush, fullRowRect)
            End Using
        End If

        ' -- [PASUL 2] LINII (TREE LINES) --
        ' -- [PASUL 2] LINII (TREE LINES) --
        DrawTreeLines(g, it, y, expanderCenterX, midY, gridLeft)

        ' === LOGICĂ DESENARE LOADER (REVIZUITĂ) ===
        If it.IsLoader Then
            ' 1. Recalculăm poziția X exactă bazată pe nivelul nodului curent
            ' gridLeft este marginea stângă a nivelului. Adăugăm Indentarea și Spațiul de Expander.
            Dim currentGridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START
            Dim loaderX As Integer = currentGridLeft + Indent + PADDING_EXPANDER_GAP

            ' 2. Calculăm poziția Y Centrată (pentru spinner și text)
            Dim loaderY As Integer = y + (ItemHeight - 14) \ 2
            Dim textY_Local As Integer = y + (ItemHeight - Me.Font.Height) \ 2 + 1

            ' 3. Setăm Grafica
            Dim oldSmoothing = g.SmoothingMode
            g.SmoothingMode = SmoothingMode.AntiAlias

            ' 4. Desenăm Spinner-ul (La poziția loaderX)
            Using p As New Pen(Color.DimGray, 2)
                g.DrawArc(p, loaderX, loaderY, 14, 14, _loadingAngle, 300)
            End Using

            ' 5. Desenăm Textul (La loaderX + 20px spațiu)
            g.DrawString("Se încarcă...", Me.Font, Brushes.Gray, loaderX + 20, textY_Local)

            ' 6. Restore și Ieșire
            g.SmoothingMode = oldSmoothing
            Return
        End If
        ' ===========================================
        ' -- [PASUL 3] CHECKBOX MODERN --
        Dim chkRect As Rectangle
        If _checkBoxes Then
            Dim chkSize As Integer = _checkBoxSize

            Dim chkY As Integer = midY - (chkSize \ 2)
            chkRect = New Rectangle(xBase, chkY, chkSize, chkSize)

            Dim oldSmoothing = g.SmoothingMode
            g.SmoothingMode = SmoothingMode.AntiAlias

            Dim accentColor As Color = Color.DodgerBlue
            Dim borderColor As Color = Color.FromArgb(180, 180, 180)

            If it.CheckState = TreeCheckState.Checked Then
                Using brush As New SolidBrush(accentColor)
                    g.FillRectangle(brush, chkRect)
                End Using
                Using pen As New Pen(accentColor)
                    g.DrawRectangle(pen, chkRect)
                End Using
                Using penTick As New Pen(Color.White, 2)
                    Dim p1 As New Point(chkRect.X + 3, chkRect.Y + 8)
                    Dim p2 As New Point(chkRect.X + 6, chkRect.Y + 11)
                    Dim p3 As New Point(chkRect.X + 12, chkRect.Y + 4)
                    g.DrawLines(penTick, {p1, p2, p3})
                End Using

            ElseIf it.CheckState = TreeCheckState.Indeterminate Then
                Using brush As New SolidBrush(accentColor)
                    g.FillRectangle(brush, chkRect)
                End Using
                Using pen As New Pen(accentColor)
                    g.DrawRectangle(pen, chkRect)
                End Using
                Using penDash As New Pen(Color.White, 2)
                    Dim yMidLine As Integer = chkRect.Y + (chkRect.Height \ 2)
                    g.DrawLine(penDash, chkRect.X + 3, yMidLine, chkRect.Right - 3, yMidLine)
                End Using
            Else
                g.FillRectangle(Brushes.White, chkRect)
                Using pen As New Pen(borderColor, 1)
                    g.DrawRectangle(pen, chkRect)
                End Using
            End If

            g.SmoothingMode = oldSmoothing

            ' Avansăm cu lățimea checkbox-ului + PADDING_CHECKBOX_GAP
            xBase += chkSize + PADDING_CHECKBOX_GAP
        End If

        ' -- [PASUL 4] CALCUL CONȚINUT (Icon + Caption) --
        Dim leftIconY As Integer = y + (ItemHeight - LeftIconSize.Height) \ 2
        Dim leftIconRect As New Rectangle(xBase, leftIconY, LeftIconSize.Width, LeftIconSize.Height)

        If it.TextWidth = -1 Then it.TextWidth = CInt(g.MeasureString(it.Caption, Me.Font).Width)

        ' Calcul poziție text cu PADDING_ICON_GAP
        Dim textX As Integer = If(it.LeftIconClosed IsNot Nothing, leftIconRect.Right + PADDING_ICON_GAP, xBase)
        Dim textY As Integer = y + (ItemHeight - Me.Font.Height) \ 2 + 1

        ' -- [PASUL 5] EXPANDER (+/-) --
        Dim showExpander As Boolean = (it.Children.Count > 0 OrElse it.LazyNode)
        If it.Level = 0 AndAlso Not _rootButton Then showExpander = False

        If showExpander Then
            g.FillRectangle(Brushes.White, expanderRect)
            g.DrawRectangle(New Pen(LineColor), expanderRect)
            g.DrawLine(Pens.Black, expanderRect.Left + 2, midY, expanderRect.Right - 2, midY)
            If Not it.Expanded Then
                g.DrawLine(Pens.Black, expanderCenterX, expanderRect.Top + 2, expanderCenterX, expanderRect.Bottom - 2)
            End If
        End If

        ' -- [PASUL 6] DESENARE CONȚINUT FINAL --
        If it.Expanded Then
            If it.LeftIconOpen IsNot Nothing Then g.DrawImage(it.LeftIconOpen, leftIconRect)
        Else
            If it.LeftIconClosed IsNot Nothing Then g.DrawImage(it.LeftIconClosed, leftIconRect)
        End If

        g.DrawString(it.Caption, Me.Font, Brushes.Black, textX, textY)

        ' -- [PASUL 7] ICONIȚĂ DREAPTA --
        If it.RightIcon IsNot Nothing Then
            Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)
            Dim rx As Integer = Me.Width - RightIconSize.Width - 6 - scrollW
            Dim ry As Integer = y + (ItemHeight - RightIconSize.Height) \ 2
            g.DrawImage(it.RightIcon, rx, ry, RightIconSize.Width, RightIconSize.Height)
        End If
    End Sub

    ' Am adăugat parametrul 'gridLeft' pentru a nu-l recalcula degeaba
    Private Sub DrawTreeLines(g As Graphics, it As TreeItem, y As Integer, expCenterX As Integer, midY As Integer, currentGridLeft As Integer)
        Using p As New Pen(LineColor)
            p.DashStyle = DashStyle.Dot

            ' 1. Linia Orizontală (Expander -> Conținut)
            Dim startH As Integer = expCenterX + (ExpanderSize \ 2) + 2
            If it.Children.Count = 0 Then startH = expCenterX

            ' Linia se duce până unde începe conținutul (xBase) minus 2 pixeli
            Dim endH As Integer = currentGridLeft + Indent + PADDING_EXPANDER_GAP - 2
            g.DrawLine(p, startH, midY, endH, midY)

            ' 2. Linia Verticală Sus
            If it.Parent IsNot Nothing Then
                g.DrawLine(p, expCenterX, y, expCenterX, midY)
            End If

            ' 3. Linia Verticală Jos
            If it.Parent IsNot Nothing AndAlso Not it.IsLastSibling Then
                g.DrawLine(p, expCenterX, midY, expCenterX, y + ItemHeight)
            End If

            ' 4. Liniile Strămoșilor
            Dim ancestor As TreeItem = it.Parent
            While ancestor IsNot Nothing
                If ancestor.Parent IsNot Nothing AndAlso Not ancestor.IsLastSibling Then
                    ' Recalculăm poziția pentru nivelul strămoșului folosind constantele
                    Dim ancGridLeft As Integer = (ancestor.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START
                    Dim ancExpCenterX As Integer = ancGridLeft + (Indent \ 2)

                    g.DrawLine(p, ancExpCenterX, y, ancExpCenterX, y + ItemHeight)
                End If
                ancestor = ancestor.Parent
            End While
        End Using
    End Sub
End Class
