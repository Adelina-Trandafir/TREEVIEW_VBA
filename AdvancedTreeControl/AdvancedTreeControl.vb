Imports System.Drawing.Drawing2D
Imports System.Reflection

Partial Public Class AdvancedTreeControl
    Inherits ScrollableControl

    ' STARE INTERNĂ (STATE)
    Private pHoveredItem As TreeItem = Nothing
    Private pSelectedItem As TreeItem = Nothing
    Private pOldSelectedItem As TreeItem = Nothing

    Private ReadOnly pToolTip As New ToolTip()
    Private ReadOnly pTooltipTimer As New Timer()
    Private pTooltipItem As TreeItem = Nothing

    ' Timer pentru a diferenția Click de DoubleClick
    Private WithEvents ClickDelayTimer As New Timer()
    Private _pendingClickItem As TreeItem = Nothing
    Private _pendingMouseArgs As MouseEventArgs = Nothing

    ' Marginea globală din stânga a întregului arbore (să nu fie lipit de margine)
    Private Const PADDING_TREE_START As Integer = 5

    ' SPAȚIUL DINTRE EXPANDER/LINIE ȘI CONȚINUT (Checkbox sau Icon)
    ' Mărește această valoare pentru a depărta bifa de liniile punctate!
    Private Const PADDING_EXPANDER_GAP As Integer = 12

    ' Spațiu între Checkbox și următorul element (Icon/Text)
    Private Const PADDING_CHECKBOX_GAP As Integer = 8

    ' Spațiu între Iconiță (stânga) și Text
    Private Const PADDING_ICON_GAP As Integer = 4

    ' Separator pentru comanda de procesare venita din VBA
    Private Shared ReadOnly separator As String() = New String() {"||"}

    ' Timer pentru animația de încărcare / Nod
    Private WithEvents _loadingTimer As New Timer() With {.Interval = 50} ' 20 FPS
    Private _loadingAngle As Single = 0

    Private Structure RichTextPart
        Public Text As String
        Public Font As Font
        Public ForeColor As Color
        Public BackColor As Color
        Public HasBackColor As Boolean
    End Structure

    ' INIȚIALIZARE
    Public Sub New()
        Me.DoubleBuffered = True
        Me.AutoScroll = True
        Me.BackColor = Color.White
        Me.Cursor = Cursors.Default
        Me.Font = New Font("Segoe UI", 9)
        Me.Enabled = True
        'Me._rightIconSize = Me._rightIconSize * CInt(Me.DeviceDpi) / 96
        pToolTip.ShowAlways = False
        pTooltipTimer.Interval = TooltipDelayMs
        AddHandler pTooltipTimer.Tick, AddressOf TooltipTimerTick

        RecalculateItemHeight()

        ClickDelayTimer.Interval = 50
    End Sub

    Private Sub OnClickDelayTimerTick(sender As Object, e As EventArgs) Handles ClickDelayTimer.Tick
        ClickDelayTimer.Stop()

        If _pendingClickItem IsNot Nothing AndAlso _pendingMouseArgs IsNot Nothing Then
            RaiseEvent NodeMouseUp(_pendingClickItem, _pendingMouseArgs)
            pOldSelectedItem = pSelectedItem
        End If

        _pendingClickItem = Nothing
        _pendingMouseArgs = Nothing
    End Sub

    Private Sub RecalculateItemHeight()
        ' Dacă utilizatorul a setat manual înălțimea, NU mai recalculăm
        If Not _autoHeight Then Return

        ' Înălțimea = Maximul dintre Font și Iconițe + Padding
        Dim hFont As Integer = CInt(Me.Font.Height)
        Dim hIcon As Integer = Math.Max(_leftIconSize.Height, _rightIconSize.Height)
        Dim hMax As Integer = Math.Max(hFont, hIcon)

        _itemHeight = hMax + 6
        Me.Invalidate()
    End Sub

    ' Opțional: Adaugă o metodă pentru a reveni la Auto
    Public Sub SetAutoHeight()
        _autoHeight = True
        RecalculateItemHeight()
    End Sub

    Private Function HitTestItem(p As Point) As TreeItem
        Dim yRel = p.Y - Me.AutoScrollPosition.Y
        Dim idx As Integer = yRel \ ItemHeight
        Dim visible = GetVisibleItems()
        If idx < 0 OrElse idx >= visible.Count Then Return Nothing
        Return visible(idx)
    End Function

    Private Function GetCheckBoxRect(it As TreeItem) As Rectangle
        If Not NodeHasCheckControl(it) Then Return Rectangle.Empty

        Dim y As Integer = GetItemY(it)
        If y = -1 Then Return Rectangle.Empty ' Item invizibil

        ' --- ACTUALIZARE LOGICĂ POZIȚIONARE ---
        ' 1. Punctul de start al grilei (același ca la DrawItem)
        Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

        ' 2. Checkbox-ul începe după Indent + PADDING_EXPANDER_GAP
        Dim xChk As Integer = gridLeft + Indent + PADDING_EXPANDER_GAP

        Dim midY As Integer = y + (ItemHeight \ 2)
        Dim chkSize As Integer = _checkBoxSize

        Return New Rectangle(xChk, midY - (chkSize \ 2), chkSize, chkSize)
    End Function

    Private Function GetExpanderRect(it As TreeItem) As Rectangle
        Dim y As Integer = GetItemY(it)
        If y = -1 Then Return Rectangle.Empty

        ' --- ACTUALIZARE LOGICĂ POZIȚIONARE ---
        ' 1. Punctul de start al grilei
        Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

        ' 2. Centrul expanderului este la jumătatea indentării curente
        Dim cx As Integer = gridLeft + (Indent \ 2)
        Dim cy As Integer = y + (ItemHeight \ 2)

        Return New Rectangle(cx - (ExpanderSize \ 2), cy - (ExpanderSize \ 2), ExpanderSize, ExpanderSize)
    End Function

    Private Function GetItemY(it As TreeItem) As Integer
        Dim idx = GetVisibleItems().IndexOf(it)
        If idx < 0 Then Return -1
        Return Me.AutoScrollPosition.Y + idx * ItemHeight
    End Function

    ' Găsește ancestorul de pe RadioButtonLevel al unui nod
    Private Function GetRadioAncestor(it As TreeItem) As TreeItem
        Dim current As TreeItem = it.Parent
        While current IsNot Nothing
            If current.Level = _radioButtonLevel Then Return current
            current = current.Parent
        End While
        Return Nothing
    End Function

    ' Determină dacă un nod trebuie să aibă checkbox/radio desenat și activ
    Private Function NodeHasCheckControl(it As TreeItem) As Boolean
        If _radioButtonLevel >= 0 Then
            If it.Level < _radioButtonLevel Then Return False                    ' deasupra: niciodată
            If it.Level = _radioButtonLevel Then Return True                     ' nivelul radio: întotdeauna
            ' sub nivel radio: doar dacă ancestorul radio e selectat
            Dim radioAnc As TreeItem = GetRadioAncestor(it)
            Return radioAnc IsNot Nothing AndAlso radioAnc.IsRadioSelected
        Else
            Return _checkBoxes                                                    ' mod normal
        End If
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

        Try
            Dim pt As Point = Me.PointToClient(Cursor.Position)
            pToolTip.Show(pTooltipItem.Caption, Me, pt.X, pt.Y + 20, 4000)

        Catch ex As Exception
            TreeLogger.Ex(ex, "TooltipTimerTick")
        End Try
    End Sub

    Private Function TextFits(it As TreeItem) As Boolean
        Using g As Graphics = Me.CreateGraphics()
            Dim textSize = g.MeasureString(it.Caption, Me.Font)

            ' 1. Calculăm punctul de start al grilei (Sincronizat cu DrawItem / Helpers)
            Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START

            ' 2. Calculăm poziția curentă X (cursorul virtual de desenare)
            '    Pornim de la zona de după Expander
            Dim currentX As Integer = gridLeft + Indent + PADDING_EXPANDER_GAP

            ' 3. Adăugăm lățimea Checkbox-ului + Spațiul de după el (dacă e activ)
            If NodeHasCheckControl(it) Then
                currentX += _checkBoxSize + PADDING_CHECKBOX_GAP
            End If

            ' 4. Adăugăm lățimea Iconiței din stânga + Spațiul de după ea
            '    Verificăm dacă există iconiță (Closed sau Open, dimensiunea e dată de LeftIconSize)
            If it.LeftIconClosed IsNot Nothing OrElse it.LeftIconOpen IsNot Nothing Then
                currentX += LeftIconSize.Width + PADDING_ICON_GAP
            End If

            ' 5. Adăugăm lățimea Textului pentru a afla punctul final
            Dim endX As Integer = currentX + CInt(textSize.Width)

            ' 6. Calculăm limita vizibilă a ferestrei
            '    Scădem zona rezervată iconiței din dreapta și o marjă de siguranță (20px)
            Dim visibleWidth As Integer = Me.Width - RightIconSize.Width - 20

            '    Scădem și lățimea barei de scroll vertical dacă este vizibilă
            If Me.VerticalScroll.Visible Then visibleWidth -= SystemInformation.VerticalScrollBarWidth

            ' Verificăm dacă textul încape
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

    ' Resetează recursiv checkboxurile tuturor descendenților unui nod
    Private Sub ClearChildrenCheckboxes(node As TreeItem)
        For Each child In node.Children
            child.CheckState = TreeCheckState.Unchecked
            ClearChildrenCheckboxes(child)
        Next
    End Sub

    ' Bifează recursiv toți descendenții unui nod
    Private Sub CheckChildrenRecursive(node As TreeItem)
        For Each child In node.Children
            child.CheckState = TreeCheckState.Checked
            CheckChildrenRecursive(child)
        Next
    End Sub

    Private Sub LoadingTimer_Tick(sender As Object, e As EventArgs) Handles _loadingTimer.Tick
        _loadingAngle += 15
        If _loadingAngle >= 360 Then _loadingAngle = 0

        ' Invalidăm doar zona vizibilă pentru a redesena animația
        ' Optimizare: Am putea invalida doar nodurile loader, dar Invalidate() e suficient pentru început
        Me.Invalidate()
    End Sub
End Class