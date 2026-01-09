Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' AdvancedTreeControl
''' EXTINS cu:
''' + MouseDown / MouseUp PE NODE
''' + Tooltip cu delay
''' + Tooltip afișat DOAR dacă textul NU încape vizibil
''' </summary>
Public Class AdvancedTreeControl
    Inherits ScrollableControl

    ' ======================================================
    ' MODEL
    ' ======================================================

    Public Class TreeItem
        Public Text As String
        Public Children As New List(Of TreeItem)
        Public Expanded As Boolean = True
        Public Parent As TreeItem
        Public Level As Integer
        Public RightIcon As Image
        Public Tag As Object
    End Class

    Public ReadOnly Items As New List(Of TreeItem)

    ' ======================================================
    ' CONFIG
    ' ======================================================

    Public ItemHeight As Integer = 22
    Public Indent As Integer = 20
    Public RightIconHeight As Integer = 16
    Public LineColor As Color = Color.LightGray
    Public HoverBackColor As Color = Color.FromArgb(230, 240, 255)
    Public SelectedBackColor As Color = Color.FromArgb(200, 220, 255)
    Public TooltipDelayMs As Integer = 600

    ' ======================================================
    ' STATE
    ' ======================================================

    Private pHoveredItem As TreeItem = Nothing
    Private pSelectedItem As TreeItem = Nothing

    Private ReadOnly pToolTip As New ToolTip()
    Private ReadOnly pTooltipTimer As New Timer()
    Private pTooltipItem As TreeItem = Nothing

    ' ======================================================
    ' EVENTS
    ' ======================================================

    Public Event TreeMouseDown(e As MouseEventArgs)
    Public Event TreeMouseUp(e As MouseEventArgs)
    Public Event TreeMouseLeave()

    Public Event NodeMouseDown(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeMouseUp(pNode As TreeItem, e As MouseEventArgs)

    ' ======================================================
    ' INIT
    ' ======================================================

    Public Sub New()
        Me.DoubleBuffered = True
        Me.AutoScroll = True
        Me.BackColor = Color.White

        pToolTip.ShowAlways = False
        pTooltipTimer.Interval = TooltipDelayMs
        AddHandler pTooltipTimer.Tick, AddressOf TooltipTimerTick
    End Sub

    ' ======================================================
    ' PAINT
    ' ======================================================

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        MyBase.OnPaint(e)

        Dim y As Integer = Me.AutoScrollPosition.Y
        For Each it In GetVisibleItems()
            DrawItem(e.Graphics, it, y)
            y += ItemHeight
        Next
    End Sub

    Private Sub DrawItem(g As Graphics, it As TreeItem, y As Integer)
        Dim xBase As Integer = it.Level * Indent + Me.AutoScrollPosition.X

        If it Is pSelectedItem Then
            g.FillRectangle(New SolidBrush(SelectedBackColor), 0, y, Me.Width, ItemHeight)
        ElseIf it Is pHoveredItem Then
            g.FillRectangle(New SolidBrush(HoverBackColor), 0, y, Me.Width, ItemHeight)
        End If

        DrawLines(g, it, y)

        If it.Children.Count > 0 Then
            Dim r As New Rectangle(xBase - 10, y + (ItemHeight \ 2) - 5, 10, 10)
            g.DrawRectangle(Pens.Gray, r)
            g.DrawLine(Pens.Black, r.Left + 2, r.Top + 5, r.Right - 2, r.Top + 5)
            If Not it.Expanded Then
                g.DrawLine(Pens.Black, r.Left + 5, r.Top + 2, r.Left + 5, r.Bottom - 2)
            End If
        End If

        g.DrawString(it.Text, Me.Font, Brushes.Black, xBase, y + 3)

        If it.RightIcon IsNot Nothing Then
            Dim s As Integer = RightIconHeight
            Dim rx As Integer = Me.Width - s - 6
            Dim ry As Integer = y + (ItemHeight - s) \ 2
            g.DrawImage(it.RightIcon, rx, ry, s, s)
        End If
    End Sub

    Private Sub DrawLines(g As Graphics, it As TreeItem, y As Integer)
        Dim midY As Integer = y + ItemHeight \ 2
        Dim x As Integer = it.Level * Indent + Me.AutoScrollPosition.X - 10

        If it.Parent IsNot Nothing Then
            g.DrawLine(New Pen(LineColor), x, y, x, y + ItemHeight)
        End If

        g.DrawLine(New Pen(LineColor), x, midY, x + 10, midY)
    End Sub

    ' ======================================================
    ' MOUSE – CONTROL + NODE
    ' ======================================================

    Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        RaiseEvent TreeMouseDown(e)

        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        pSelectedItem = it
        RaiseEvent NodeMouseDown(it, e)

        Dim toggleRect = GetToggleRect(it)
        If toggleRect.Contains(e.Location) AndAlso it.Children.Count > 0 Then
            it.Expanded = Not it.Expanded
        End If

        Me.Invalidate()
    End Sub

    Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
        MyBase.OnMouseUp(e)
        RaiseEvent TreeMouseUp(e)

        Dim it = HitTestItem(e.Location)
        If it Is Nothing Then Return

        RaiseEvent NodeMouseUp(it, e)
    End Sub

    Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
        MyBase.OnMouseMove(e)

        Dim it = HitTestItem(e.Location)
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
        RaiseEvent TreeMouseLeave()
        Me.Invalidate()
    End Sub

    ' ======================================================
    ' TOOLTIP LOGIC
    ' ======================================================

    Private Sub ResetTooltip(it As TreeItem)
        pToolTip.Hide(Me)
        pTooltipTimer.Stop()
        pTooltipItem = Nothing

        If it Is Nothing Then Return
        If TextFits(it) Then Return

        pTooltipItem = it
        pTooltipTimer.Start()
    End Sub

    Private Sub TooltipTimerTick(sender As Object, e As EventArgs)
        pTooltipTimer.Stop()

        If pTooltipItem Is Nothing OrElse pTooltipItem IsNot pHoveredItem Then Return

        Dim y As Integer = GetItemY(pTooltipItem)
        If y < 0 Then Return

        pToolTip.Show(pTooltipItem.Text, Me, 10, y + ItemHeight, 4000)
    End Sub

    Private Function TextFits(it As TreeItem) As Boolean
        Using g As Graphics = Me.CreateGraphics()
            Dim textSize = g.MeasureString(it.Text, Me.Font)
            Dim xBase As Integer = it.Level * Indent + Me.AutoScrollPosition.X
            Dim availableWidth As Integer = Me.Width - xBase - RightIconHeight - 10
            Return textSize.Width <= availableWidth
        End Using
    End Function

    ' ======================================================
    ' HIT TEST
    ' ======================================================

    Private Function HitTestItem(p As Point) As TreeItem
        Dim idx As Integer = (p.Y - Me.AutoScrollPosition.Y) \ ItemHeight
        Dim visible = GetVisibleItems()
        If idx < 0 OrElse idx >= visible.Count Then Return Nothing
        Return visible(idx)
    End Function

    Private Function GetToggleRect(it As TreeItem) As Rectangle
        Dim y As Integer = GetItemY(it)
        Dim x As Integer = it.Level * Indent + Me.AutoScrollPosition.X - 10
        Return New Rectangle(x, y + (ItemHeight \ 2) - 5, 10, 10)
    End Function

    Private Function GetItemY(it As TreeItem) As Integer
        Dim idx = GetVisibleItems().IndexOf(it)
        If idx < 0 Then Return -1
        Return Me.AutoScrollPosition.Y + idx * ItemHeight
    End Function

    ' ======================================================
    ' VISIBLE ITEMS
    ' ======================================================

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
    ' API
    ' ======================================================

    Public Function AddItem(pText As String, Optional pParent As TreeItem = Nothing, Optional pRightIcon As Image = Nothing) As TreeItem
        Dim it As New TreeItem With {
            .Text = pText,
            .Parent = pParent,
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

End Class
