Imports System.Drawing.Drawing2D
Imports System.Text.RegularExpressions

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

        ' === CALCUL LIMITĂ TEXT (Clipping) ===
        Dim scrollW As Integer = If(Me.VerticalScroll.Visible, SystemInformation.VerticalScrollBarWidth, 0)

        ' Limita din dreapta a controlului (minus padding 6px)
        Dim maxRightX As Integer = Me.Width - 6 - scrollW

        ' Dacă există RightIcon, limita se mută mai la stânga (lățime icon + încă 6px padding)
        If it.RightIcon IsNot Nothing Then
            maxRightX -= (RightIconSize.Width + 6)
        End If

        ' Calculăm lățimea disponibilă pentru text
        Dim availableTextWidth As Integer = maxRightX - textX
        If availableTextWidth < 0 Then availableTextWidth = 0

        ' Salvăm starea curentă a "foarfecii" (Clip)
        Dim oldClip As Region = g.Clip.Clone()

        ' Setăm noua zonă de tăiere: Textul se va desena DOAR în acest dreptunghi
        Dim clipRect As New Rectangle(textX, y, availableTextWidth, ItemHeight)
        g.SetClip(clipRect)

        ' --- DESENARE TEXT ---
        Dim baseTextColor As Color = Color.Black
        DrawRichText(g, it.Caption, textX, y, Me.Font, baseTextColor)
        ' ---------------------

        ' Restaurăm "foarfeca" originală pentru a putea desena RightIcon (care e în afara zonei de text)
        g.Clip = oldClip


        ' -- [PASUL 7] ICONIȚĂ DREAPTA --
        If it.RightIcon IsNot Nothing Then
            ' Recalculăm poziția exact cum am calculat limita mai sus
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

    ' =================================================================================
    ' SUPORT RICH TEXT (BOLD, ITALIC, COLOR)
    ' =================================================================================
    ' Desenează textul formatat și returnează lățimea totală (pentru calcule)
    Private Sub DrawRichText(g As Graphics, text As String, x As Integer, y As Integer, defaultFont As Font, defaultColor As Color)
        Dim parts As List(Of RichTextPart) = ParseRichText(text, defaultFont, defaultColor)
        Dim currentX As Single = x

        ' Setăm formatarea pentru a elimina spațierea extra a GDI+
        Dim fmt As StringFormat = StringFormat.GenericTypographic
        fmt.FormatFlags = fmt.FormatFlags Or StringFormatFlags.MeasureTrailingSpaces

        For Each part In parts
            ' 1. Desenăm fundalul (dacă există)
            Dim size As SizeF = g.MeasureString(part.Text, part.Font, PointF.Empty, fmt)

            If part.HasBackColor Then
                Using b As New SolidBrush(part.BackColor)
                    ' Ajustăm puțin rect-ul pe verticală pentru a arăta bine
                    g.FillRectangle(b, currentX, y, size.Width, ItemHeight)
                End Using
            End If

            ' 2. Desenăm Textul
            Using b As New SolidBrush(part.ForeColor)
                ' Ajustăm Y pentru centrare verticală în funcție de font
                Dim textY As Single = y + (ItemHeight - part.Font.Height) / 2
                g.DrawString(part.Text, part.Font, b, currentX, textY, fmt)
            End Using

            ' 3. Avansăm cursorul X
            currentX += size.Width
        Next
    End Sub

    ' Parser simplu bazat pe Regex
    Private Function ParseRichText(rawText As String, baseFont As Font, baseColor As Color) As List(Of RichTextPart)
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

    Private Function ParseColor(val As String, defaultColor As Color) As Color
        Try
            If String.IsNullOrEmpty(val) Then Return defaultColor
            If val.StartsWith("#") Then Return ColorTranslator.FromHtml(val)
            Return Color.FromName(val)
        Catch
            Return defaultColor
        End Try
    End Function
End Class
