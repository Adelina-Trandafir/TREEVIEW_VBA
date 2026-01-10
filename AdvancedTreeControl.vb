Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

''' <summary>
''' AdvancedTreeControl
''' Control personalizat tip TreeView:
''' - Desenare GDI+ completă (fără WndProc complex)
''' - Linii punctate continue (stil ierarhic)
''' - Expanderi (+/-) interactivi
''' - Suport Iconiță Stânga + Iconiță Dreapta
''' - Selecție vizuală doar pe Text+Icon
''' - Tooltip inteligent (doar dacă textul e trunchiat)
''' - Auto-Height (calculat după font/iconițe)
''' </summary>
Public Class AdvancedTreeControl
    Inherits ScrollableControl

    ' ======================================================
    ' 1. MODELUL DE DATE (TreeItem)
    ' ======================================================
    Public Class TreeItem
        Public Text As String
        Public Children As New List(Of TreeItem)
        Public Expanded As Boolean = True
        Public Parent As TreeItem
        Public Level As Integer

        Public LeftIcon As Image
        Public RightIcon As Image

        Public Tag As Object

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
    End Class

    Public ReadOnly Items As New List(Of TreeItem)

    ' ======================================================
    ' 2. CONFIGURARE & PROPRIETĂȚI
    ' ======================================================

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
    Public ExpanderSize As Integer = 9

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

    ' ======================================================
    ' 5. INIȚIALIZARE
    ' ======================================================
    Public Sub New()
        Me.DoubleBuffered = True
        Me.AutoScroll = True
        Me.BackColor = Color.White
        Me.Cursor = Cursors.Default
        Me.Font = New Font("Segoe UI", 9)

        pToolTip.ShowAlways = False
        pTooltipTimer.Interval = TooltipDelayMs
        AddHandler pTooltipTimer.Tick, AddressOf TooltipTimerTick

        RecalculateItemHeight()
    End Sub

    Protected Overrides Sub OnFontChanged(e As EventArgs)
        MyBase.OnFontChanged(e)
        RecalculateItemHeight()
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
        ' -- COORDONATE --
        ' xBase este punctul unde începe zona nodului (după indentare)
        ' +10 este marginea din stânga controlului
        Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + 10

        ' Centrul expanderului pe axa X (pentru alinierea liniilor verticale)
        Dim expanderCenterX As Integer = xBase - Indent + (Indent \ 2)
        Dim midY As Integer = y + (ItemHeight \ 2)

        ' Dreptunghiuri
        Dim expanderRect As New Rectangle(expanderCenterX - (ExpanderSize \ 2), midY - (ExpanderSize \ 2), ExpanderSize, ExpanderSize)
        Dim leftIconY As Integer = y + (ItemHeight - LeftIconSize.Height) \ 2
        Dim leftIconRect As New Rectangle(xBase, leftIconY, LeftIconSize.Width, LeftIconSize.Height)

        ' Calcul text
        If it.TextWidth = -1 Then it.TextWidth = CInt(g.MeasureString(it.Text, Me.Font).Width)

        Dim textX As Integer = If(it.LeftIcon IsNot Nothing, leftIconRect.Right + 4, xBase)
        Dim textY As Integer = y + (ItemHeight - Me.Font.Height) \ 2 + 1
        Dim textRect As New Rectangle(textX, y, it.TextWidth + 4, ItemHeight)

        ' -- A. LINII (TREE LINES) --
        DrawTreeLines(g, it, y, expanderCenterX, midY)

        ' -- B. SELECȚIE / HOVER (Doar Text + Icon) --
        Dim contentRect As Rectangle = textRect
        If it.LeftIcon IsNot Nothing Then
            contentRect = Rectangle.Union(leftIconRect, textRect)
            contentRect.Inflate(2, 0)
        Else
            contentRect.X -= 2
            contentRect.Width += 4
        End If

        ' Ajustare fină pe verticală
        contentRect.Y += 1
        contentRect.Height -= 2

        If it Is pSelectedItem Then
            g.FillRectangle(New SolidBrush(SelectedBackColor), contentRect)
            g.DrawRectangle(New Pen(SelectedBorderColor), contentRect)
        ElseIf it Is pHoveredItem Then
            g.FillRectangle(New SolidBrush(HoverBackColor), contentRect)
        End If

        ' -- C. EXPANDER (+/-) --
        If it.Children.Count > 0 Then
            g.FillRectangle(Brushes.White, expanderRect)
            g.DrawRectangle(New Pen(LineColor), expanderRect)

            ' Linia orizontală (-)
            g.DrawLine(Pens.Black, expanderRect.Left + 2, midY, expanderRect.Right - 2, midY)

            ' Linia verticală (|) -> devine +
            If Not it.Expanded Then
                g.DrawLine(Pens.Black, expanderCenterX, expanderRect.Top + 2, expanderCenterX, expanderRect.Bottom - 2)
            End If
        End If

        ' -- D. CONȚINUT --
        If it.LeftIcon IsNot Nothing Then
            g.DrawImage(it.LeftIcon, leftIconRect)
        End If
        g.DrawString(it.Text, Me.Font, Brushes.Black, textX, textY)

        ' -- E. ICONIȚĂ DREAPTA --
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

            ' 1. Linia Orizontală (de la Expander/Linia Verticală spre Text)
            Dim startH As Integer = expCenterX + (ExpanderSize \ 2) + 2
            If it.Children.Count = 0 Then startH = expCenterX ' Dacă n-are expander, linia pleacă din centru

            Dim endH As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + 10 - 2
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
                    Dim ancExpCenterX As Integer = (ancestor.Level * Indent) + Me.AutoScrollPosition.X + 10 - Indent + (Indent \ 2)
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

        ' Verificăm Click pe Expander
        Dim expRect = GetExpanderRect(it)
        If expRect.Contains(e.Location) AndAlso it.Children.Count > 0 Then
            it.Expanded = Not it.Expanded
            Me.Invalidate()
            Return ' Click pe expander nu selectează nodul
        End If

        ' Altfel, Selectăm Nodul
        pSelectedItem = it
        RaiseEvent NodeMouseDown(it, e)
        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        Dim it = HitTestItem(e.Location)
        If it IsNot Nothing Then
            RaiseEvent NodeMouseUp(it, e)
        End If
    End Sub

    Protected Overrides Sub OnMouseDoubleClick(e As MouseEventArgs)
        MyBase.OnMouseDoubleClick(e)
        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        ' Dublu click oriunde pe rând face Toggle Expand
        If it.Children.Count > 0 Then
            it.Expanded = Not it.Expanded
            Me.Invalidate()
        End If
        RaiseEvent NodeDoubleClicked(it, e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim it = HitTestItem(e.Location)
        If it IsNot pHoveredItem Then
            pHoveredItem = it
            ResetTooltip(it) ' Resetare tooltip la schimbarea rândului
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

    Private Function GetExpanderRect(it As TreeItem) As Rectangle
        Dim y As Integer = GetItemY(it)
        Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + 10
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

    Private Sub AddVisible(it As TreeItem, list As List(Of TreeItem))
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
        pToolTip.Show(pTooltipItem.Text, Me, pt.X, pt.Y + 20, 4000)
    End Sub

    Private Function TextFits(it As TreeItem) As Boolean
        Using g As Graphics = Me.CreateGraphics()
            Dim textSize = g.MeasureString(it.Text, Me.Font)
            Dim xBase As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + 10
            Dim leftIconW As Integer = If(it.LeftIcon IsNot Nothing, LeftIconSize.Width + 4, 0)

            Dim endX As Integer = xBase + leftIconW + CInt(textSize.Width)
            Dim visibleWidth As Integer = Me.Width - RightIconSize.Width - 20 ' Margine siguranță
            If Me.VerticalScroll.Visible Then visibleWidth -= SystemInformation.VerticalScrollBarWidth

            Return endX <= visibleWidth
        End Using
    End Function

    ' ======================================================
    ' 10. API PUBLIC
    ' ======================================================
    Public Function AddItem(pText As String, Optional pParent As TreeItem = Nothing, Optional pLeftIcon As Image = Nothing, Optional pRightIcon As Image = Nothing) As TreeItem
        Dim it As New TreeItem With {
            .Text = pText,
            .Parent = pParent,
            .LeftIcon = pLeftIcon,
            .RightIcon = pRightIcon
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

    Public ReadOnly Property SelectedNode As TreeItem
        Get
            Return pSelectedItem
        End Get
    End Property

    Public Sub Clear()
        Items.Clear()
        pSelectedItem = Nothing
        pHoveredItem = Nothing
        Me.Invalidate()
    End Sub

End Class