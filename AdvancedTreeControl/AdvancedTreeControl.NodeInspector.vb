
Partial Public Class AdvancedTreeControl

    ' ==========================================================================
    ' ENTRY POINT  —  apelat din OnMouseDown cand Right-Click este pe expander
    ' ==========================================================================
    Friend Sub ShowNodeInspector(it As TreeItem)
        Dim info As New NodeDebugInfo()

        ' ──────────────────────────────────────────────────────────────────────
        ' 1. MODEL — toate campurile din TreeItem
        ' ──────────────────────────────────────────────────────────────────────
        info.Key = If(it.Key, "")
        info.Caption = If(it.Caption, "")
        info.Level = it.Level
        info.Expanded = it.Expanded
        info.ParentKey = If(it.Parent IsNot Nothing, If(it.Parent.Key, ""), "")
        info.ChildCount = it.Children.Count
        info.CheckState = it.CheckState.ToString()
        info.HasCheckBox = it.HasCheckBox
        info.Bold = it.Bold
        info.Italic = it.Italic
        info.Tooltip = If(it.Tooltip, "")
        info.Tag = If(it.Tag IsNot Nothing, it.Tag.ToString(), "")
        info.LazyNode = it.LazyNode
        info.ShowRightIconOnHover = it.ShowRightIconOnHover
        info.IsLoader = it.IsLoader
        info.IsRadioSelected = it.IsRadioSelected
        info.IsLastSibling = it.IsLastSibling
        info.ColHeaderText = If(it.ColHeaderText, "")
        info.HasLeftIconClosed = (it.LeftIconClosed IsNot Nothing)
        info.HasLeftIconOpen = (it.LeftIconOpen IsNot Nothing)
        info.HasRightIcon = (it.RightIcon IsNot Nothing)
        info.TextWidth_Cache = it.TextWidth
        info.LastClickedColumnIndex = it.LastClickedColumnIndex
        info.LastClickedColumnName = If(it.LastClickedColumnName, "")

        info.NodeForeColor = If(it.NodeForeColor = Color.Empty,
                               "Empty (mosteneste control)",
                               $"#{it.NodeForeColor.R:X2}{it.NodeForeColor.G:X2}{it.NodeForeColor.B:X2}  (R:{it.NodeForeColor.R} G:{it.NodeForeColor.G} B:{it.NodeForeColor.B})")

        info.NodeBackColor = If(it.NodeBackColor = Color.Empty,
                               "Empty (transparent)",
                               $"#{it.NodeBackColor.R:X2}{it.NodeBackColor.G:X2}{it.NodeBackColor.B:X2}  (R:{it.NodeBackColor.R} G:{it.NodeBackColor.G} B:{it.NodeBackColor.B})")

        ' ──────────────────────────────────────────────────────────────────────
        ' 2. LAYOUT — calculat exact ca in DrawItem / Painting.vb
        ' ──────────────────────────────────────────────────────────────────────
        Dim y As Integer = GetItemY(it)
        Dim gridLeft As Integer = (it.Level * Indent) + Me.AutoScrollPosition.X + PADDING_TREE_START
        Dim expCX As Integer = gridLeft + (Indent \ 2)   ' centrul expanderului pe X
        Dim midY As Integer = y + (ItemHeight \ 2)

        ' xBase — exact ca in DrawItem
        Dim xBase As Integer = If(it.Level = 0 AndAlso Not _RootExpander,
                                  gridLeft,
                                  gridLeft + Indent + PADDING_EXPANDER_GAP)

        ' NodeBounds
        info.NodeBounds = If(y = -1, Rectangle.Empty,
                             New Rectangle(0, y, Me.ClientSize.Width, ItemHeight))

        ' ExpanderBounds — din GetExpanderRect (exact)
        info.ExpanderBounds = GetExpanderRect(it)

        ' CheckBoxBounds — din GetCheckBoxRect (exact)
        info.CheckBoxBounds = GetCheckBoxRect(it)

        ' LeftIconBounds — replica logica din DrawContent
        Dim leftIconBounds As Rectangle = Rectangle.Empty
        If _hasNodeIcons Then
            Dim xIcon As Integer = xBase
            ' Daca exista checkbox, iconita incepe dupa el
            If info.CheckBoxBounds <> Rectangle.Empty Then
                xIcon = info.CheckBoxBounds.Right + PADDING_CHECKBOX_GAP
            End If
            If it.LeftIconClosed IsNot Nothing OrElse it.LeftIconOpen IsNot Nothing Then
                leftIconBounds = New Rectangle(xIcon,
                                               y + (ItemHeight - _leftIconSize.Height) \ 2,
                                               _leftIconSize.Width, _leftIconSize.Height)
            End If
        End If
        info.LeftIconBounds = leftIconBounds

        ' TextBounds — replica logica din DrawContent
        Dim textX As Integer = If(leftIconBounds <> Rectangle.Empty,
                                  leftIconBounds.Right + PADDING_ICON_GAP,
                                  xBase)
        ' Daca exista checkbox dar nu iconita, textul incepe dupa checkbox
        If leftIconBounds = Rectangle.Empty AndAlso info.CheckBoxBounds <> Rectangle.Empty Then
            textX = info.CheckBoxBounds.Right + PADDING_CHECKBOX_GAP
        End If

        Dim scrollW As Integer = ScrollBarWidth
        Dim maxRightX As Integer = Me.Width - scrollW - PADDING_TREE_END
        If it.RightIcon IsNot Nothing Then maxRightX -= (RightIconSize.Width + PADDING_RIGHT_ICON_GAP)
        info.TextBounds = If(y = -1, Rectangle.Empty,
                             New Rectangle(textX, y, Math.Max(0, maxRightX - textX), ItemHeight))

        ' RightIconBounds — replica logica din DrawRightIcon
        Dim rightIconBounds As Rectangle = Rectangle.Empty
        If it.RightIcon IsNot Nothing Then
            Dim rx As Integer = Me.Width - RightIconSize.Width - PADDING_RIGHT_ICON_GAP - PADDING_TREE_END - scrollW
            rightIconBounds = New Rectangle(rx, y + (ItemHeight - RightIconSize.Height) \ 2,
                                            RightIconSize.Width, RightIconSize.Height)
        End If
        info.RightIconBounds = rightIconBounds

        info.GridLeft = gridLeft
        info.XBase = xBase
        info.MidY = midY
        info.ExpanderCenterX = expCX

        Dim visibleItems = GetVisibleItems()
        info.IndexInVisibleList = visibleItems.IndexOf(it)
        info.IsInViewport = (y <> -1) AndAlso (y + ItemHeight > 0) AndAlso (y < Me.Height)

        ' SelectionBounds — replica logica din DrawSelection
        Dim selStartX As Integer
        If it.Level = 0 AndAlso Not _RootExpander Then
            selStartX = gridLeft
        Else
            selStartX = gridLeft + ExpanderSize * 2 - 3
        End If
        info.SelectionBounds = If(y = -1, Rectangle.Empty,
                                    New Rectangle(selStartX, y,
                                                  Math.Max(0, Me.ClientSize.Width - selStartX - PADDING_TREE_END),
                                                  ItemHeight))

        ' ──────────────────────────────────────────────────────────────────────
        ' 3. CELLS — celulele TreeListView
        ' ──────────────────────────────────────────────────────────────────────
        info.CellCount = it.Cells.Count
        If it.Cells.Count > 0 Then
            Dim sb As New System.Text.StringBuilder()
            For Each kvp In it.Cells
                Dim cd = kvp.Value
                sb.Append($"{kvp.Key}: ""{cd.Value}""")
                If cd.BackColor <> Color.Empty Then sb.Append($"  bg={cd.BackColor.Name}")
                If cd.ForeColor <> Color.Empty Then sb.Append($"  fg={cd.ForeColor.Name}")
                sb.AppendLine()
            Next
            info.CellsData = sb.ToString().TrimEnd()
        End If

        ' ──────────────────────────────────────────────────────────────────────
        ' 4. RENDERER — setarile controlului
        ' ──────────────────────────────────────────────────────────────────────
        info.ItemHeight = Me.ItemHeight
        info.Indent = Me.Indent
        info.ExpanderSize = Me.ExpanderSize
        info.CheckBoxSize = Me._checkBoxSize
        info.LeftIconSize = $"{_leftIconSize.Width} × {_leftIconSize.Height}"
        info.RightIconSize = $"{RightIconSize.Width} × {RightIconSize.Height}"
        info.HasNodeIcons = Me._hasNodeIcons
        info.CheckBoxes = Me._checkBoxes
        info.RootExpander = Me._RootExpander
        info.ShowRightIconOnHover_Global = Me._showRightIconOnHover
        info.ControlWidth = Me.ClientSize.Width
        info.ControlHeight = Me.ClientSize.Height
        info.ScrollBarVisible = _vScroll.Visible
        info.ScrollOffsetY = _vScroll.Value
        info.IsSelectedNode = (it Is pSelectedItem)
        info.IsHoveredNode = (it Is pHoveredItem)

        info.TreeControl_Bounds = Me.Bounds   ' relativ la container

        Dim parentFrm = Me.FindForm()
        If parentFrm IsNot Nothing Then
            info.ParentForm_Size = $"{parentFrm.Width} × {parentFrm.Height}"
            info.ParentForm_Bounds = $"X={parentFrm.Left}  Y={parentFrm.Top}  W={parentFrm.Width}  H={parentFrm.Height}"
        End If

        ' ──────────────────────────────────────────────────────────────────────
        ' Afisam form-ul
        ' ──────────────────────────────────────────────────────────────────────
        FrmNodeDebug.ShowForNode(info, Me.FindForm())
    End Sub

End Class