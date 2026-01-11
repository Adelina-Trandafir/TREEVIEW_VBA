Imports System.Drawing.Drawing2D
Imports System.Reflection
''' <summary>
''' AdvancedTreeControl
''' Control personalizat tip TreeView:
''' - Desenare GDI+ completă (fără WndProc complex)
''' - Linii punctate continue (stil ierarhic)
''' - Expanderi (+/-) interactivi
''' - Suport Iconiță Stânga + Iconiță Dreapta
''' - Selecție vizuală doar pe Caption+Icon
''' - Tooltip inteligent (doar dacă textul e trunchiat)
''' - Auto-Height (calculat după font/iconițe)
''' </summary>
Public Class AdvancedTreeControl
    Inherits ScrollableControl
    Public Enum TreeCheckState
        Unchecked = 0       ' Nebifat
        Checked = 1         ' Bifat complet
        Indeterminate = 2   ' Parțial bifat (pătrățel plin sau liniuță)
    End Enum

    ' ======================================================
    ' 1. MODELUL DE DATE (TreeItem)
    ' ======================================================
    Public Class TreeItem
        Public Key As String
        Public Caption As String
        Public Children As New List(Of TreeItem)
        Public Expanded As Boolean = True
        Public Parent As TreeItem
        Public Level As Integer
        Public CheckState As TreeCheckState = TreeCheckState.Unchecked
        Public LeftIconClosed As Image
        Public LeftIconOpen As Image
        Public RightIcon As Image

        Private _tag As Object

        ' Cache pentru lățimea textului (performanță la desenare)
        Friend TextWidth As Integer = -1

        ' Proprietate critică pentru desenarea corectă a liniilor verticale
        Public ReadOnly Property IsLastSibling As Boolean
            Get
                If Parent Is Nothing Then
                    ' Dacă e root, verificăm dacă e ultimul în lista principală a controlului
                    ' (Necesită referință la control, dar pentru simplitate desenăm standard)
                    Return True
                End If
                Return Parent.Children.LastOrDefault() Is Me
            End Get
        End Property

        Public Property Tag As Object
            Get
                Return _tag
            End Get
            Set(value As Object)
                _tag = value
            End Set
        End Property
    End Class

    Public ReadOnly Items As New List(Of TreeItem)

    ' ======================================================
    ' 2. CONFIGURARE & PROPRIETĂȚI
    ' ======================================================
    Private _checkBoxSize As Integer = 16
    Public Property CheckBoxSize As Integer
        Get
            Return _checkBoxSize
        End Get
        Set(value As Integer)
            _checkBoxSize = value
            Me.Invalidate()
        End Set
    End Property

    ' Înălțimea rândului (calculată automat sau setată manual)
    Private _itemHeight As Integer = 22
    Public Property ItemHeight As Integer
        Get
            Return _itemHeight
        End Get
        Set(value As Integer)
            _itemHeight = value
            Me.Invalidate()
        End Set
    End Property

    Public Indent As Integer = 20
    Public ExpanderSize As Integer = 12

    ' Iconițe - Setarea lor declanșează recalcularea înălțimii rândului
    Private _leftIconSize As New Size(24, 24)
    Public Property LeftIconSize As Size
        Get
            Return _leftIconSize
        End Get
        Set(value As Size)
            _leftIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rightIconSize As New Size(24, 24)
    Public Property RightIconSize As Size
        Get
            Return _rightIconSize
        End Get
        Set(value As Size)
            _rightIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rootHasExpander As Boolean = True
    Public Property RootHasExpander As Boolean
        Get
            Return _rootHasExpander
        End Get
        Set(value As Boolean)
            _rootHasExpander = value
            Me.Invalidate()
        End Set
    End Property

    ' Culori
    Public LineColor As Color = Color.FromArgb(160, 160, 160)
    Public HoverBackColor As Color = Color.FromArgb(230, 240, 255)
    Public SelectedBackColor As Color = Color.FromArgb(200, 220, 255)
    Public SelectedBorderColor As Color = Color.FromArgb(150, 180, 255)

    ' Tooltip
    Public TooltipDelayMs As Integer = 600

    ' ======================================================
    ' 3. STARE INTERNĂ (STATE)
    ' ======================================================
    Private pHoveredItem As TreeItem = Nothing
    Private pSelectedItem As TreeItem = Nothing

    Private ReadOnly pToolTip As New ToolTip()
    Private ReadOnly pTooltipTimer As New Timer()
    Private pTooltipItem As TreeItem = Nothing

    ' ======================================================
    ' 4. EVENIMENTE PUBLICE
    ' ======================================================
    Public Event NodeMouseDown(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeMouseUp(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeDoubleClicked(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeChecked(pNode As TreeItem)

    ' Timer pentru a diferenția Click de DoubleClick
    Private WithEvents ClickDelayTimer As New Timer()
    Private _pendingClickItem As TreeItem = Nothing
    Private _pendingMouseArgs As MouseEventArgs = Nothing

    ' ======================================================
    ' 5. INIȚIALIZARE
    ' ======================================================
    Public Sub New()
        Me.DoubleBuffered = True
        Me.AutoScroll = True
        Me.BackColor = Color.White
        Me.Cursor = Cursors.Default
        Me.Font = New Font("Segoe UI", 9)
        Me._rightIconSize = Me._rightIconSize * CInt(Me.DeviceDpi) / 96
        pToolTip.ShowAlways = False
        pTooltipTimer.Interval = TooltipDelayMs
        AddHandler pTooltipTimer.Tick, AddressOf TooltipTimerTick

        RecalculateItemHeight()

        ClickDelayTimer.Interval = SystemInformation.DoubleClickTime
    End Sub

    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        RecalculateItemHeight()
    End Sub

    Private Sub OnClickDelayTimerTick(sender As Object, e As EventArgs) Handles ClickDelayTimer.Tick
        ClickDelayTimer.Stop()

        ' Dacă timer-ul a expirat, înseamnă că nu a urmat un al doilea click.
        ' Putem trimite evenimentul de Click (MouseUp) acum.
        If _pendingClickItem IsNot Nothing AndAlso _pendingMouseArgs IsNot Nothing Then
            RaiseEvent NodeMouseUp(_pendingClickItem, _pendingMouseArgs)
        End If

        _pendingClickItem = Nothing
        _pendingMouseArgs = Nothing
    End Sub

    Private Sub RecalculateItemHeight()
        ' Înălțimea = Maximul dintre Font și Iconițe + Padding
        Dim hFont As Integer = CInt(Me.Font.Height)
        Dim hIcon As Integer = Math.Max(_leftIconSize.Height, _rightIconSize.Height)
        Dim hMax As Integer = Math.Max(hFont, hIcon)

        _itemHeight = hMax + 6
        Me.Invalidate()
    End Sub

    ' ======================================================
    ' 6. ENGINE DE DESENARE (PAINT)
    ' ======================================================
    ' Marginea globală din stânga a întregului arbore (să nu fie lipit de margine)
    Private Const PADDING_TREE_START As Integer = 5

    ' SPAȚIUL DINTRE EXPANDER/LINIE ȘI CONȚINUT (Checkbox sau Icon)
    ' Mărește această valoare pentru a depărta bifa de liniile punctate!
    Private Const PADDING_EXPANDER_GAP As Integer = 12

    ' Spațiu între Checkbox și următorul element (Icon/Text)
    Private Const PADDING_CHECKBOX_GAP As Integer = 8

    ' Spațiu între Iconiță (stânga) și Text
    Private Const PADDING_ICON_GAP As Integer = 4
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

    Private Sub DrawItem(g As Graphics, it As TreeItem, y As Integer)
        ' Forțăm expandarea root-ului dacă nu are expander permis
        If it.Level = 0 AndAlso Not _rootHasExpander AndAlso Not it.Expanded Then
            it.Expanded = True
        End If

        ' -- COORDONATE DE BAZĂ --
        ' xBase este punctul unde începe zona nodului (după indentare)
        Dim leftMargin As Integer = Indent + 5
        Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin

        ' Centrul expanderului pe axa X
        Dim expanderCenterX As Integer = xBase - Indent + (Indent \ 2)
        Dim midY As Integer = y + (ItemHeight \ 2)
        Dim expanderRect As New Rectangle(expanderCenterX - (ExpanderSize \ 2), midY - (ExpanderSize \ 2), ExpanderSize, ExpanderSize)

        ' -----------------------------------------------------------------------
        ' -- [PASUL 1] SELECȚIE & HOVER (FULL ROW - LOGICĂ CORECTĂ) --
        ' Desenăm asta PRIMA DATĂ, ca să fie în spatele tuturor elementelor.
        ' -----------------------------------------------------------------------
        ' 1. Calculăm punctul de start logic al acestui nod (include indentarea)
        '    Adăugăm un mic padding (5px) ca să nu fie lipit de liniile părinților
        Dim selStartX As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + 5

        ' 2. Calculăm lățimea: De la punctul de start până la marginea dreaptă a ferestrei
        Dim selWidth As Integer = Me.ClientSize.Width - selStartX

        ' Protecție: Să nu avem lățime negativă dacă intrăm mult în scroll
        If selWidth < 0 Then selWidth = 0

        Dim fullRowRect As New Rectangle(selStartX, y, selWidth, ItemHeight)

        If it Is pSelectedItem Then
            ' Desenăm fundalul selecției
            Using brush As New SolidBrush(SelectedBackColor)
                g.FillRectangle(brush, fullRowRect)
            End Using
            ' Desenăm chenarul selecției
            Using pen As New Pen(SelectedBorderColor)
                Dim borderRect As New Rectangle(selStartX, y, selWidth - 1, ItemHeight - 1)
                g.DrawRectangle(pen, borderRect)
            End Using
        ElseIf it Is pHoveredItem Then
            ' Desenăm fundalul de hover
            Using brush As New SolidBrush(HoverBackColor)
                g.FillRectangle(brush, fullRowRect)
            End Using
        End If
        ' -----------------------------------------------------------------------

        ' -- [PASUL 2] LINII (TREE LINES) --
        ' Le desenăm ACUM, peste fundalul de selecție
        DrawTreeLines(g, it, y, expanderCenterX, midY)

        ' -- [PASUL 3] CHECKBOX MODERN --
        Dim chkRect As Rectangle
        If _checkBoxes Then
            Dim chkSize As Integer = _checkBoxSize

            ' Centrat vertical
            Dim chkY As Integer = midY - (chkSize \ 2)
            chkRect = New Rectangle(xBase, chkY, chkSize, chkSize)

            ' Setări grafice pentru calitate înaltă (AntiAlias)
            Dim oldSmoothing = g.SmoothingMode
            g.SmoothingMode = SmoothingMode.AntiAlias

            ' Culori
            Dim accentColor As Color = Color.DodgerBlue
            Dim borderColor As Color = Color.FromArgb(180, 180, 180)

            If it.CheckState = TreeCheckState.Checked Then
                ' --- STARE BIFATĂ (CHECKED) ---
                Using brush As New SolidBrush(accentColor)
                    g.FillRectangle(brush, chkRect)
                End Using
                Using pen As New Pen(accentColor)
                    g.DrawRectangle(pen, chkRect)
                End Using

                ' Bifa (v)
                Using penTick As New Pen(Color.White, 2)
                    Dim p1 As New Point(chkRect.X + 3, chkRect.Y + 8)
                    Dim p2 As New Point(chkRect.X + 6, chkRect.Y + 11)
                    Dim p3 As New Point(chkRect.X + 12, chkRect.Y + 4)
                    g.DrawLines(penTick, {p1, p2, p3})
                End Using

            ElseIf it.CheckState = TreeCheckState.Indeterminate Then
                ' --- STARE NEDEFINITĂ (INDETERMINATE) ---
                Using brush As New SolidBrush(accentColor)
                    g.FillRectangle(brush, chkRect)
                End Using
                Using pen As New Pen(accentColor)
                    g.DrawRectangle(pen, chkRect)
                End Using

                ' Linie orizontală (-) albă la mijloc
                Using penDash As New Pen(Color.White, 2)
                    Dim yMidLine As Integer = chkRect.Y + (chkRect.Height \ 2)
                    g.DrawLine(penDash, chkRect.X + 3, yMidLine, chkRect.Right - 3, yMidLine)
                End Using

            Else
                ' --- STARE NEBIFATĂ (UNCHECKED) ---
                g.FillRectangle(Brushes.White, chkRect)
                Using pen As New Pen(borderColor, 1)
                    g.DrawRectangle(pen, chkRect)
                End Using
            End If

            ' Restaurăm setările grafice
            g.SmoothingMode = oldSmoothing

            ' Împingem conținutul (Icon + Caption) mai la dreapta
            xBase += chkSize + 8
        End If

        ' -- [PASUL 4] CALCUL CONȚINUT (Icon + Caption) --
        Dim leftIconY As Integer = y + (ItemHeight - LeftIconSize.Height) \ 2
        Dim leftIconRect As New Rectangle(xBase, leftIconY, LeftIconSize.Width, LeftIconSize.Height)

        If it.TextWidth = -1 Then it.TextWidth = CInt(g.MeasureString(it.Caption, Me.Font).Width)

        Dim textX As Integer = If(it.LeftIconClosed IsNot Nothing, leftIconRect.Right + 4, xBase)
        Dim textY As Integer = y + (ItemHeight - Me.Font.Height) \ 2 + 1

        ' -- [PASUL 5] EXPANDER (+/-) --
        Dim showExpander As Boolean = (it.Children.Count > 0)
        ' Dacă e Root (Level 0) și am dezactivat RootHasExpander, NU îl afișăm
        If it.Level = 0 AndAlso Not _rootHasExpander Then
            showExpander = False
        End If

        If showExpander Then
            g.FillRectangle(Brushes.White, expanderRect)
            g.DrawRectangle(New Pen(LineColor), expanderRect)

            ' Linia orizontală (-)
            g.DrawLine(Pens.Black, expanderRect.Left + 2, midY, expanderRect.Right - 2, midY)

            ' Linia verticală (|) -> devine +
            If Not it.Expanded Then
                g.DrawLine(Pens.Black, expanderCenterX, expanderRect.Top + 2, expanderCenterX, expanderRect.Bottom - 2)
            End If
        End If

        ' -- [PASUL 6] DESENARE CONȚINUT FINAL --
        If it.Expanded Then
            If it.LeftIconOpen IsNot Nothing Then g.DrawImage(it.LeftIconOpen, leftIconRect)
            Debug.Print("EXP:" & it.Caption)
        Else
            If it.LeftIconClosed IsNot Nothing Then g.DrawImage(it.LeftIconClosed, leftIconRect)
            Debug.Print("COL:" & it.Caption)
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

    Private Sub DrawTreeLines(g As Graphics, it As TreeItem, y As Integer, expCenterX As Integer, midY As Integer)
        Using p As New Pen(LineColor)
            p.DashStyle = DashStyle.Dot

            ' --- CORECȚIE: Sincronizare margine cu DrawItem ---
            ' Înainte era hardcodat + 10. Acum trebuie să fie Indent + 5,
            ' exact cum am definit leftMargin în DrawItem.
            Dim leftMargin As Integer = Indent + 5
            ' --------------------------------------------------

            ' 1. Linia Orizontală (de la Expander/Linia Verticală spre Caption)
            Dim startH As Integer = expCenterX + (ExpanderSize \ 2) + 2
            If it.Children.Count = 0 Then startH = expCenterX ' Dacă n-are expander, linia pleacă din centru

            ' FIX: Folosim leftMargin în loc de 10 pentru a nimeri exact textul/checkbox-ul
            Dim endH As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin - 2
            g.DrawLine(p, startH, midY, endH, midY)

            ' 2. Linia Verticală Sus (către Părinte)
            If it.Parent IsNot Nothing Then
                g.DrawLine(p, expCenterX, y, expCenterX, midY)
            End If

            ' 3. Linia Verticală Jos (către Frate Următor)
            ' Se desenează DOAR dacă nu sunt ultimul frate
            If it.Parent IsNot Nothing AndAlso Not it.IsLastSibling Then
                g.DrawLine(p, expCenterX, midY, expCenterX, y + ItemHeight)
            End If

            ' 4. Liniile Verticale ale Strămoșilor (Liniile lungi din stânga)
            Dim ancestor As TreeItem = it.Parent
            While ancestor IsNot Nothing
                ' Dacă un strămoș nu e ultimul, linia lui trebuie să continue vizual în jos prin dreptul nostru
                If ancestor.Parent IsNot Nothing AndAlso Not ancestor.IsLastSibling Then

                    ' FIX: Recalculăm poziția expanderului strămoșului folosind NOUA margine (leftMargin)
                    ' Formula trebuie să fie identică cu cea din DrawItem pentru a se alinia perfect
                    Dim ancExpCenterX As Integer = (ancestor.Level * Indent) + Me.AutoScrollPosition.X + leftMargin - Indent + (Indent \ 2)

                    g.DrawLine(p, ancExpCenterX, y, ancExpCenterX, y + ItemHeight)
                End If
                ancestor = ancestor.Parent
            End While
        End Using
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
            If it.Level = 0 AndAlso Not _rootHasExpander Then
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

                ' CRITIC: Returnăm și aici. 
                ' De obicei, când bifezi, NU vrei să schimbi și selecția rândului.
                ' Dacă totuși vrei să se și selecteze rândul când bifezi, șterge linia de mai jos.
                Return
            End If
        End If

        ' =================================================================
        ' 3. PRIORITATE DOI: SELECȚIE RÂND (TEXT / ICON)
        ' Ajungem aici DOAR dacă nu s-a dat click pe Expander sau Checkbox
        ' =================================================================
        pSelectedItem = it
        RaiseEvent NodeMouseDown(it, e)
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
            If it.Level = 0 AndAlso Not _rootHasExpander Then
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

        ' Opțional: Schimbăm cursorul dacă e peste checkbox/expander, 
        ' dar dacă it devine Nothing mai sus, cursorul va rămâne Default, ceea ce e corect.
    End Sub

    Protected Overrides Sub OnMouseLeave(e As EventArgs)
        MyBase.OnMouseLeave(e)
        pHoveredItem = Nothing
        pToolTip.Hide(Me)
        pTooltipTimer.Stop()
        Me.Invalidate()
    End Sub

    ' ======================================================
    ' 8. HELPERS (HitTest, VisibleItems, Rects)
    ' ======================================================
    Private Function HitTestItem(p As Point) As TreeItem
        Dim yRel = p.Y - Me.AutoScrollPosition.Y
        Dim idx As Integer = yRel \ ItemHeight
        Dim visible = GetVisibleItems()
        If idx < 0 OrElse idx >= visible.Count Then Return Nothing
        Return visible(idx)
    End Function

    Private Function GetCheckBoxRect(it As TreeItem) As Rectangle
        If Not _checkBoxes Then Return Rectangle.Empty

        Dim y As Integer = GetItemY(it)
        If y = -1 Then Return Rectangle.Empty ' Item invizibil

        ' FIX: Folosim aceeași margine ca la desenare
        Dim leftMargin As Integer = Indent + 5
        Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin

        Dim midY As Integer = y + (ItemHeight \ 2)

        ' FIX: Dimensiune 16 (la fel ca în DrawItem modern)
        Return New Rectangle(xBase, midY - (_checkBoxSize \ 2), _checkBoxSize, _checkBoxSize)
    End Function

    Private Function GetExpanderRect(it As TreeItem) As Rectangle
        Dim y As Integer = GetItemY(it)
        If y = -1 Then Return Rectangle.Empty ' Protecție extra

        ' FIX: Folosim aceeași margine
        Dim leftMargin As Integer = Indent + 5
        Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin

        Dim cx As Integer = xBase - Indent + (Indent \ 2)
        Dim cy As Integer = y + (ItemHeight \ 2)

        Return New Rectangle(cx - (ExpanderSize \ 2), cy - (ExpanderSize \ 2), ExpanderSize, ExpanderSize)
    End Function

    Private Function GetItemY(it As TreeItem) As Integer
        Dim idx = GetVisibleItems().IndexOf(it)
        If idx < 0 Then Return -1
        Return Me.AutoScrollPosition.Y + idx * ItemHeight
    End Function

    ' Returnează lista plată a nodurilor vizibile (ținând cont de expandare)
    Private Function GetVisibleItems() As List(Of TreeItem)
        Dim result As New List(Of TreeItem)
        For Each it In Items
            AddVisible(it, result)
        Next
        Return result
    End Function

    Private Shared Sub AddVisible(it As TreeItem, list As List(Of TreeItem))
        list.Add(it)
        If it.Expanded Then
            For Each c In it.Children
                AddVisible(c, list)
            Next
        End If
    End Sub

    ' ======================================================
    ' 9. TOOLTIP LOGIC
    ' ======================================================
    Private Sub ResetTooltip(it As TreeItem)
        pToolTip.Hide(Me)
        pTooltipTimer.Stop()
        pTooltipItem = Nothing

        If it Is Nothing Then Return
        If TextFits(it) Then Return ' Nu afișăm dacă încape

        pTooltipItem = it
        pTooltipTimer.Start()
    End Sub

    Private Sub TooltipTimerTick(sender As Object, e As EventArgs)
        pTooltipTimer.Stop()
        If pTooltipItem Is Nothing OrElse pTooltipItem IsNot pHoveredItem Then Return

        Dim pt As Point = Me.PointToClient(Cursor.Position)
        pToolTip.Show(pTooltipItem.Caption, Me, pt.X, pt.Y + 20, 4000)
    End Sub

    Private Function TextFits(it As TreeItem) As Boolean
        Using g As Graphics = Me.CreateGraphics()
            Dim textSize = g.MeasureString(it.Caption, Me.Font)

            ' FIX: Margine actualizată
            Dim leftMargin As Integer = Indent + 5
            Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + leftMargin

            ' Luăm în calcul și lățimea checkbox-ului dacă există
            Dim chkOffset As Integer = If(_checkBoxes, _checkBoxSize + 8, 0) '+ 8 padding

            Dim leftIconW As Integer = If(it.LeftIconClosed IsNot Nothing, LeftIconSize.Width + 4, 0)

            Dim endX As Integer = xBase + chkOffset + leftIconW + CInt(textSize.Width)
            Dim visibleWidth As Integer = Me.Width - RightIconSize.Width - 20
            If Me.VerticalScroll.Visible Then visibleWidth -= SystemInformation.VerticalScrollBarWidth

            Return endX <= visibleWidth
        End Using
    End Function

    ' Setează starea unui nod, a copiilor săi și actualizează părinții
    Private Shared Sub SetNodeStateWithPropagation(node As TreeItem, newState As TreeCheckState)
        ' 1. Setează starea nodului curent
        node.CheckState = newState

        ' 2. Propagă în jos (toți copiii iau aceeași stare)
        SetChildrenStateRecursive(node, newState)

        ' 3. Propagă în sus (părinții își recalculează starea)
        UpdateParentStateRecursive(node.Parent)
    End Sub

    ' Setează recursiv toți descendenții la o anumită stare
    Private Shared Sub SetChildrenStateRecursive(node As TreeItem, state As TreeCheckState)
        For Each child In node.Children
            child.CheckState = state
            SetChildrenStateRecursive(child, state)
        Next
    End Sub

    ' Verifică starea fraților și actualizează părintele
    Private Shared Sub UpdateParentStateRecursive(parent As TreeItem)
        If parent Is Nothing Then Return

        Dim anyChecked As Boolean = False
        Dim anyUnchecked As Boolean = False
        Dim anyIndeterminate As Boolean = False

        For Each child In parent.Children
            Select Case child.CheckState
                Case TreeCheckState.Checked
                    anyChecked = True
                Case TreeCheckState.Unchecked
                    anyUnchecked = True
                Case TreeCheckState.Indeterminate
                    anyIndeterminate = True
            End Select
        Next

        ' Reguli pentru starea părintelui:
        If anyIndeterminate Then
            parent.CheckState = TreeCheckState.Indeterminate
        ElseIf anyChecked AndAlso anyUnchecked Then
            parent.CheckState = TreeCheckState.Indeterminate ' Mixt -> Nedefinit
        ElseIf anyChecked Then
            parent.CheckState = TreeCheckState.Checked       ' Toți bifati
        Else
            parent.CheckState = TreeCheckState.Unchecked     ' Nimeni bifat
        End If

        ' Continuăm urcarea spre rădăcină
        UpdateParentStateRecursive(parent.Parent)
    End Sub

    ' ======================================================
    ' 10. API PUBLIC
    ' ======================================================
    Public Function AddItem(pKey As String, pCaption As String,
                            Optional pParent As TreeItem = Nothing,
                            Optional pLeftIconClosed As Image = Nothing,
                            Optional pLeftIconOpen As Image = Nothing,
                            Optional pRightIcon As Image = Nothing,
                            Optional pTag As String = Nothing,
                            Optional pExpanded As Boolean = False) As TreeItem
        Dim it As New TreeItem With {
            .Key = pKey,
            .Tag = pTag,
            .Caption = pCaption,
            .Parent = pParent,
            .LeftIconClosed = pLeftIconClosed,
            .LeftIconOpen = pLeftIconOpen,
            .RightIcon = pRightIcon,
            .Expanded = pExpanded
        }

        If pParent Is Nothing Then
            it.Level = 0
            Items.Add(it)
        Else
            it.Level = pParent.Level + 1
            pParent.Children.Add(it)
        End If

        Me.Invalidate()
        Return it
    End Function

    ' Funcția care primește string-ul din VBA și returnează valoarea
    Public Function ProcessPropertyRequest(cmd As String) As String
        ' Format așteptat: "GET_PROPERTY||PropName||[OptionalNodeID]"
        Dim parts() As String = cmd.Split(separator, StringSplitOptions.None)

        If parts.Length < 2 Then Return "ERROR: Invalid Format"

        Dim propName As String = parts(1)
        Dim result As String = "NOT_FOUND"

        Try
            ' === CAZUL 1: PROPRIETATE A CONTROLULUI (GLOBAL) ===
            If parts.Length = 2 Then
                ' Căutăm proprietatea în clasa AdvancedTreeControl (Me)
                Dim propInfo As PropertyInfo = Me.GetType().GetProperty(propName, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)

                If propInfo IsNot Nothing Then
                    Dim val = propInfo.GetValue(Me, Nothing)
                    result = FormatValue(val)
                Else
                    result = "ERROR: Property '" & propName & "' not found on Tree."
                End If

                ' === CAZUL 2: PROPRIETATE A UNUI NOD ===
            ElseIf parts.Length = 3 Then
                Dim nodeID As String = parts(2)

                ' 1. Găsim nodul după ID (care e Key în VBA)
                Dim node As TreeItem = FindNodeByID(nodeID)

                If node IsNot Nothing Then
                    ' 2. Căutăm proprietatea în clasa TreeItem
                    Dim propInfo As PropertyInfo = node.GetType().GetProperty(propName, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)

                    If propInfo IsNot Nothing Then
                        Dim val = propInfo.GetValue(node, Nothing)
                        result = FormatValue(val)
                    Else
                        result = "ERROR: Property '" & propName & "' not found on Node."
                    End If
                Else
                    result = "ERROR: Node with ID '" & nodeID & "' not found."
                End If
            End If

        Catch ex As Exception
            result = "ERROR: " & ex.Message
        End Try

        Return result
    End Function

    ' --- HELPERS ---

    ' Funcție recursivă pentru a găsi un nod după ID
    Private Function FindNodeByID(id As String) As TreeItem
        Return FindNodeRecursive(Me.Items, id)
    End Function

    Private Shared Function FindNodeRecursive(collection As List(Of TreeItem), id As String) As TreeItem
        For Each it As TreeItem In collection
            ' Verificăm ID-ul (care corespunde cu Key din VBA)
            ' Asigură-te că ai proprietatea ID definită în TreeItem (sau folosește _tag dacă acolo ții ID-ul)
            ' Presupunând că ai: Public ID As String
            If it.Key = id Then
                Return it
            End If

            ' Căutare în adâncime
            Dim foundChild = FindNodeRecursive(it.Children, id)
            If foundChild IsNot Nothing Then Return foundChild
        Next
        Return Nothing
    End Function

    ' Funcție pentru a converti orice obiect (Boolean, Color, Enum) în String pentru VBA
    Private Shared Function FormatValue(val As Object) As String
        If val Is Nothing Then Return ""

        If TypeOf val Is Boolean Then
            ' Returnăm "True"/"False" sau "-1"/"0" cum preferă VBA
            Return If(DirectCast(val, Boolean), "True", "False")

        ElseIf TypeOf val Is Color Then
            ' Pentru culori, returnăm codul ARGB sau numele
            Return DirectCast(val, Color).Name

        ElseIf TypeOf val Is [Enum] Then
            ' Pentru Enum-uri (ex: CheckState), returnăm valoarea numerică (0, 1, 2)
            Return CInt(val).ToString()

        Else
            ' Pentru String, Integer, etc.
            Return val.ToString()
        End If
    End Function

    Public Property SelectedNode As TreeItem
        Get
            Return pSelectedItem
        End Get
        Set(value As TreeItem)
            If pSelectedItem IsNot value Then
                pSelectedItem = value
                ' Invalidate to trigger redraw and show the new selection
                Me.Invalidate()
            End If
        End Set
    End Property

    Private _checkBoxes As Boolean = False
    Public Property CheckBoxes As Boolean
        Get
            Return _checkBoxes
        End Get
        Set(value As Boolean)
            _checkBoxes = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property

    Private Shared ReadOnly separator As String() = New String() {"||"}

    Public Sub Clear()
        Items.Clear()
        pSelectedItem = Nothing
        pHoveredItem = Nothing
        Me.Invalidate()
    End Sub

End Class