Imports System.Globalization
Imports System.Xml
Imports TREEVIEW_VBA.AdvancedTreeControl

Friend NotInheritable Class TreeXmlAppliers

    ' -------------------------------------------------------
    ' CULORI
    ' -------------------------------------------------------

    Friend Shared Sub Apply_BackColor(cfg As XmlNode,
                                      tree As AdvancedTreeControl,
                                      host As Control)
        If cfg.Attributes("BackColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("BackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.BackColor <> c Then tree.BackColor = c
            If host.BackColor <> c Then host.BackColor = c
            TreeLogger.Debug(Space(5) & $"BackColor xml='{xmlVal}' control='{tree.BackColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_BorderColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("BorderColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("BorderColor").Value
            If xmlVal.StartsWith("#"c) Then
                Dim c As Color = ColorTranslator.FromHtml(xmlVal)
                If tree.BorderColor <> c Then tree.BorderColor = c
                TreeLogger.Debug(Space(5) & $"BorderColor xml='{xmlVal}' control='{tree.BorderColor}'", "AplicareConfigurare")
            End If
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_ForeColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ForeColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("ForeColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.ForeColor <> c Then tree.ForeColor = c
            TreeLogger.Debug(Space(5) & $"ForeColor xml='{xmlVal}' control='{tree.ForeColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_HoverBackColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HoverBackColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("HoverBackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.HoverBackColor <> c Then tree.HoverBackColor = c
            TreeLogger.Debug(Space(5) & $"HoverBackColor xml='{xmlVal}' control='{tree.HoverBackColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_SelectedBackColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SelectedBackColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("SelectedBackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.SelectedBackColor <> c Then tree.SelectedBackColor = c
            TreeLogger.Debug(Space(5) & $"SelectedBackColor xml='{xmlVal}' control='{tree.SelectedBackColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_SelectedBorderColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SelectedBorderColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("SelectedBorderColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.SelectedBorderColor <> c Then tree.SelectedBorderColor = c
            TreeLogger.Debug(Space(5) & $"SelectedBorderColor xml='{xmlVal}' control='{tree.SelectedBorderColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Friend Shared Sub Apply_LineColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("LineColor") Is Nothing Then Exit Sub
        Try
            Dim xmlVal As String = cfg.Attributes("LineColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)
            If tree.LineColor <> c Then tree.LineColor = c
            TreeLogger.Debug(Space(5) & $"LineColor xml='{xmlVal}' control='{tree.LineColor}'", "AplicareConfigurare")
        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' -------------------------------------------------------
    ' CHECKBOX / RADIOBUTTON
    ' -------------------------------------------------------

    Friend Shared Sub Apply_RadioButtonLevel(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RadioButtonLevel") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RadioButtonLevel").Value
        Dim v As Integer = tree.RadioButtonLevel
        If Integer.TryParse(xmlVal, v) Then
            If tree.RadioButtonLevel <> v Then tree.RadioButtonLevel = v
            TreeLogger.Debug(Space(5) & $"RadioButtonLevel xml='{xmlVal}' control='{tree.RadioButtonLevel}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_CheckBoxes(cfg As XmlNode, tree As AdvancedTreeControl)
        If tree.RadioButtonLevel <> -1 Then Exit Sub
        If cfg.Attributes("CheckBoxes") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("CheckBoxes").Value
        Dim v As Integer = If(tree.CheckBoxes, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.CheckBoxes <> nv Then tree.CheckBoxes = nv
            TreeLogger.Debug(Space(5) & $"CheckBoxes xml='{xmlVal}' control='{tree.CheckBoxes}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_CheckboxSize(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("CheckboxSize") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("CheckboxSize").Value
        Dim v As Integer = If(tree.CheckBoxSize > 0, tree.CheckBoxSize, 16)
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.CheckBoxSize <> v Then tree.CheckBoxSize = v
            TreeLogger.Debug(Space(5) & $"CheckboxSize xml='{xmlVal}' control='{tree.CheckBoxSize}'", "AplicareConfigurare")
        End If
    End Sub

    ' -------------------------------------------------------
    ' FONT
    ' -------------------------------------------------------

    Friend Shared Sub Apply_Font(cfg As XmlNode,
                                 tree As AdvancedTreeControl,
                                 culture As CultureInfo)
        Dim curFontName As String = If(tree.Font IsNot Nothing, tree.Font.Name, "Segoe UI")
        Dim curFontSize As Single = If(tree.Font IsNot Nothing, tree.Font.Size, 9.0F)
        Dim fName As String = curFontName
        Dim fSize As Single = curFontSize
        Dim xmlFontName As String = Nothing
        Dim xmlFontSize As String = Nothing
        Dim hasFontChange As Boolean = False

        If cfg.Attributes("FontName") IsNot Nothing Then
            xmlFontName = cfg.Attributes("FontName").Value
            fName = xmlFontName
            hasFontChange = True
        End If

        If cfg.Attributes("FontSize") IsNot Nothing Then
            xmlFontSize = cfg.Attributes("FontSize").Value
            Dim tmp As Single = fSize
            If Single.TryParse(xmlFontSize, NumberStyles.Any, culture, tmp) Then
                fSize = tmp
                hasFontChange = True
            End If
        End If

        If hasFontChange Then
            Dim needSet As Boolean =
                (tree.Font Is Nothing) OrElse
                (tree.Font.Name <> fName) OrElse
                (Math.Abs(tree.Font.Size - fSize) > 0.001F)
            If needSet Then tree.Font = New Font(fName, fSize)
            TreeLogger.Debug(Space(5) &
                $"Font xmlName='{xmlFontName}' xmlSize='{xmlFontSize}' control='{tree.Font.Name} {tree.Font.Size}pt'",
                "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_FontName(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("FontName") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("FontName").Value
        If tree.FontName <> xmlVal Then tree.FontName = xmlVal
        TreeLogger.Debug(Space(5) & $"FontName xml='{xmlVal}' control='{tree.FontName}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_FontSize(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("FontSize") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("FontSize").Value
        Dim v As Integer = tree.FontSize
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.FontSize <> v Then tree.FontSize = v
            TreeLogger.Debug(Space(5) & $"FontSize xml='{xmlVal}' control='{tree.FontSize}'", "AplicareConfigurare")
        End If
    End Sub

    ' -------------------------------------------------------
    ' DIMENSIUNI / LAYOUT
    ' -------------------------------------------------------

    Friend Shared Sub Apply_ItemHeight(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ItemHeight") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ItemHeight").Value
        Dim v As Integer = tree.ItemHeight
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.ItemHeight <> v Then tree.ItemHeight = v
            TreeLogger.Debug(Space(5) & $"ItemHeight xml='{xmlVal}' control='{tree.ItemHeight}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_Indent(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("Indent") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("Indent").Value
        Dim v As Integer = tree.Indent
        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then
            If tree.Indent <> v Then tree.Indent = v
            TreeLogger.Debug(Space(5) & $"Indent xml='{xmlVal}' control='{tree.Indent}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_ExpanderSize(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ExpanderSize") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ExpanderSize").Value
        Dim v As Integer = tree.ExpanderSize
        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then
            If tree.ExpanderSize <> v Then tree.ExpanderSize = v
            TreeLogger.Debug(Space(5) & $"ExpanderSize xml='{xmlVal}' control='{tree.ExpanderSize}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_LeftIconHeight(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("LeftIconHeight") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("LeftIconHeight").Value
        Dim v As Integer = If(tree.LeftIconSize.Height > 0, tree.LeftIconSize.Height, 0)
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            Dim ns As New Size(v, v)
            If tree.LeftIconSize <> ns Then tree.LeftIconSize = ns
            TreeLogger.Debug(Space(5) & $"LeftIconHeight xml='{xmlVal}' control='{tree.LeftIconSize}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_RightIconHeight(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RightIconHeight") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RightIconHeight").Value
        Dim v As Integer = If(tree.RightIconSize.Height > 0, tree.RightIconSize.Height, 0)
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            Dim ns As New Size(v, v)
            If tree.RightIconSize <> ns Then tree.RightIconSize = ns
            TreeLogger.Debug(Space(5) & $"RightIconHeight xml='{xmlVal}' control='{tree.RightIconSize}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_LeftTextWidth(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("LeftTextWidth") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("LeftTextWidth").Value
        Dim v As Integer = tree.LeftTextWidth
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.LeftTextWidth <> v Then tree.LeftTextWidth = v
            TreeLogger.Debug(Space(5) & $"LeftTextWidth xml='{xmlVal}' control='{tree.LeftTextWidth}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_RightTextWidth(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RightTextWidth") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RightTextWidth").Value
        Dim v As Integer = tree.RightTextWidth
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.RightTextWidth <> v Then tree.RightTextWidth = v
            TreeLogger.Debug(Space(5) & $"RightTextWidth xml='{xmlVal}' control='{tree.RightTextWidth}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_RightIconRightPadding(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RightIconRightPadding") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RightIconRightPadding").Value
        Dim v As Integer = tree.RightIconRightPadding
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.RightIconRightPadding <> v Then tree.RightIconRightPadding = v
            TreeLogger.Debug(Space(5) & $"RightIconRightPadding xml='{xmlVal}' control='{tree.RightIconRightPadding}'", "AplicareConfigurare")
        End If
    End Sub
    ' -------------------------------------------------------
    ' COMPORTAMENT
    ' -------------------------------------------------------

    Friend Shared Sub Apply_HasNodeIcons(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HasNodeIcons") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HasNodeIcons").Value
        Dim v As Integer = If(tree.HasNodeIcons, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.HasNodeIcons <> nv Then tree.HasNodeIcons = nv
            TreeLogger.Debug(Space(5) & $"HasNodeIcons xml='{xmlVal}' control='{tree.HasNodeIcons}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_PopupTree(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("PopupTree") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("PopupTree").Value
        Dim v As Integer = If(tree.IsPopupTree, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.IsPopupTree <> nv Then tree.IsPopupTree = nv
            TreeLogger.Debug(Space(5) & $"PopupTree xml='{xmlVal}' control='{tree.IsPopupTree}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_PopupGraceMs(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("PopupGraceMs") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("PopupGraceMs").Value
        Dim v As Integer = tree.PopupGraceMs
        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then
            If tree.PopupGraceMs <> v Then tree.PopupGraceMs = v
            TreeLogger.Debug(Space(5) & $"PopupGraceMs xml='{xmlVal}' control='{tree.PopupGraceMs}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_ShowRightIconOnHover(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ShowRightIconOnHover") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ShowRightIconOnHover").Value
        Dim v As Integer = If(tree.ShowRightIconOnHover, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.ShowRightIconOnHover <> nv Then tree.ShowRightIconOnHover = nv
            TreeLogger.Debug(Space(5) & $"ShowRightIconOnHover xml='{xmlVal}' control='{tree.ShowRightIconOnHover}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_RightClickFunc(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RightClickFunc") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RightClickFunc").Value
        If tree.RightClickFunction <> xmlVal Then tree.RightClickFunction = xmlVal
        TreeLogger.Debug(Space(5) & $"RightClickFunc xml='{xmlVal}' control='{tree.RightClickFunction}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_RootExpander(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RootExpander") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RootExpander").Value
        Dim v As Integer = If(tree.RootExpander, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.RootExpander <> nv Then tree.RootExpander = nv
            TreeLogger.Debug(Space(5) & $"RootExpander xml='{xmlVal}' control='{tree.RootExpander}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_ReRaiseClickOnSameNode(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ReRaiseClickOnSameNode") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ReRaiseClickOnSameNode").Value
        Dim v As Integer = If(tree.ReRaiseClickOnSameNode, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.ReRaiseClickOnSameNode <> nv Then tree.ReRaiseClickOnSameNode = nv
            TreeLogger.Debug(Space(5) & $"ReRaiseClickOnSameNode xml='{xmlVal}' control='{tree.ReRaiseClickOnSameNode}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_RaiseLeftClickOnRightClick(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("RaiseLeftClickOnRightClick") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("RaiseLeftClickOnRightClick").Value
        Dim v As Integer = If(tree.RaiseLeftClickOnRightClick, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.RaiseLeftClickOnRightClick <> nv Then tree.RaiseLeftClickOnRightClick = nv
            TreeLogger.Debug(Space(5) & $"RaiseLeftClickOnRightClick xml='{xmlVal}' control='{tree.RaiseLeftClickOnRightClick}'", "AplicareConfigurare")
        End If
    End Sub

    ' -------------------------------------------------------
    ' TOOLTIP
    ' -------------------------------------------------------

    Friend Shared Sub Apply_ToolTipDelayMs(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ToolTipDelayMs") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ToolTipDelayMs").Value
        Dim v As Integer = tree.TooltipDelayMs
        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then
            If tree.TooltipDelayMs <> v Then tree.TooltipDelayMs = v
            TreeLogger.Debug(Space(5) & $"ToolTipDelayMs xml='{xmlVal}' control='{tree.TooltipDelayMs}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_TooltipAutoHideMs(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("TooltipAutoHideMs") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("TooltipAutoHideMs").Value
        Dim v As Integer = tree.AutoHideTooltipMs
        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then
            If tree.AutoHideTooltipMs <> v Then tree.AutoHideTooltipMs = v
            TreeLogger.Debug(Space(5) & $"TooltipAutoHideMs xml='{xmlVal}' control='{tree.AutoHideTooltipMs}'", "AplicareConfigurare")
        End If
    End Sub

    ' -------------------------------------------------------
    ' SELECTED NODE
    ' -------------------------------------------------------

    Friend Shared Sub Apply_SelectedNodeId(cfg As XmlNode,
                                           ByRef pendingSelectedNodeId As String)
        pendingSelectedNodeId = String.Empty
        If cfg.Attributes("SelectedNodeId") Is Nothing Then Exit Sub
        pendingSelectedNodeId = cfg.Attributes("SelectedNodeId").Value
        TreeLogger.Debug(Space(5) & $"SelectedNodeId xml='{pendingSelectedNodeId}'", "AplicareConfigurare")
    End Sub

    ' -------------------------------------------------------
    ' HEADER
    ' -------------------------------------------------------

    Friend Shared Sub Apply_HeaderVisible(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderVisible") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderVisible").Value
        Dim v As Integer = If(tree.HeaderVisible, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.HeaderVisible <> nv Then tree.HeaderVisible = nv
            TreeLogger.Debug(Space(5) & $"HeaderVisible xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_HeaderHeight(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderHeight") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderHeight").Value
        Dim v As Integer = tree.HeaderHeight
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            If tree.HeaderHeight <> v Then tree.HeaderHeight = v
            TreeLogger.Debug(Space(5) & $"HeaderHeight xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_HeaderCaption(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderCaption") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderCaption").Value
        If tree.HeaderCaption <> xmlVal Then tree.HeaderCaption = xmlVal
        TreeLogger.Debug(Space(5) & $"HeaderCaption xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_HeaderIconSize(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderIconSize") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderIconSize").Value
        Dim v As Integer = tree.HeaderIconSize.Width
        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then
            Dim ns As New Size(v, v)
            If tree.HeaderIconSize <> ns Then tree.HeaderIconSize = ns
            TreeLogger.Debug(Space(5) & $"HeaderIconSize xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_HeaderBackColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderBackColor") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderBackColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.HeaderBackColor)
        If tree.HeaderBackColor <> c Then
            tree.HeaderBackColor = c
            TreeLogger.Debug(Space(5) & $"HeaderBackColor xml='{xmlVal}' control='{tree.HeaderBackColor}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_HeaderForeColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderForeColor") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("HeaderForeColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.HeaderForeColor)
        If tree.HeaderForeColor <> c Then
            tree.HeaderForeColor = c
            TreeLogger.Debug(Space(5) & $"HeaderForeColor xml='{xmlVal}' control='{tree.HeaderForeColor}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_HeaderLeftIcon(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderLeftIcon") Is Nothing Then Exit Sub
        tree.HeaderLeftIconKey = cfg.Attributes("HeaderLeftIcon").Value
    End Sub

    Friend Shared Sub Apply_HeaderRightIcon(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderRightIcon") Is Nothing Then Exit Sub
        tree.HeaderRightIconKey = cfg.Attributes("HeaderRightIcon").Value
    End Sub

    Friend Shared Sub Apply_HeaderSearchIcon(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("HeaderSearchIcon") Is Nothing Then Exit Sub
        tree.HeaderSearchIconKey = cfg.Attributes("HeaderSearchIcon").Value
    End Sub

    ' -------------------------------------------------------
    ' SEARCH
    ' -------------------------------------------------------

    Friend Shared Sub Apply_SearchShow(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchShow") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchShow").Value
        Dim nv As Boolean = (xmlVal = "1")
        If tree.SearchShow <> nv Then tree.SearchShow = nv
        TreeLogger.Debug(Space(5) & $"SearchShow xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_SearchDefaultText(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchDefaultText") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchDefaultText").Value
        If xmlVal <> "" Then tree.SearchDefaultText = xmlVal
        TreeLogger.Debug(Space(5) & $"SearchDefaultText xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_SearchType(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchType") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchType").Value
        Dim v As Integer = CInt(tree.SearchType)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv = CType(v, AdvancedTreeControl.en_Tree_SearchType)
            If tree.SearchType <> nv Then tree.SearchType = nv
            TreeLogger.Debug(Space(5) & $"SearchType xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchIn(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchIn") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchIn").Value
        Dim v As Integer = CInt(tree.SearchIn)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv = CType(v, AdvancedTreeControl.En_Tree_SearchIn)
            If tree.SearchIn <> nv Then tree.SearchIn = nv
            TreeLogger.Debug(Space(5) & $"SearchIn xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchMode(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchMode") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchMode").Value
        Dim v As Integer = CInt(tree.SearchMode)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv = CType(v, AdvancedTreeControl.En_Tree_SearchMode)
            If tree.SearchMode <> nv Then tree.SearchMode = nv
            TreeLogger.Debug(Space(5) & $"SearchMode xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchBackColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBackColor") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("SearchBackColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.SearchBackColor)
        If tree.SearchBackColor <> c Then tree.SearchBackColor = c
    End Sub

    Friend Shared Sub Apply_SearchBarLabelText(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarLabelText") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarLabelText").Value
        If tree.SearchBarLabelText <> xmlVal Then tree.SearchBarLabelText = xmlVal
        TreeLogger.Debug(Space(5) & $"SearchBarLabelText xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_SearchBarLabelForeColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarLabelForeColor") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarLabelForeColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.SearchBarLabelForeColor)
        If tree.SearchBarLabelForeColor <> c Then tree.SearchBarLabelForeColor = c
        TreeLogger.Debug(Space(5) & $"SearchBarLabelForeColor xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_SearchBarLabelBold(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarLabelBold") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarLabelBold").Value
        Dim v As Integer = If(tree.SearchBarLabelBold, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.SearchBarLabelBold <> nv Then tree.SearchBarLabelBold = nv
            TreeLogger.Debug(Space(5) & $"SearchBarLabelBold xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchBarLabelItalic(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarLabelItalic") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarLabelItalic").Value
        Dim v As Integer = If(tree.SearchBarLabelItalic, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.SearchBarLabelItalic <> nv Then tree.SearchBarLabelItalic = nv
            TreeLogger.Debug(Space(5) & $"SearchBarLabelItalic xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchBarFontName(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarFontName") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarFontName").Value
        If tree.SearchBarFontName <> xmlVal Then tree.SearchBarFontName = xmlVal
        TreeLogger.Debug(Space(5) & $"SearchBarFontName xml='{xmlVal}'", "AplicareConfigurare")
    End Sub

    Friend Shared Sub Apply_SearchBarFontSize(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchBarFontSize") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchBarFontSize").Value
        Dim v As Single = tree.SearchBarFontSize
        If Single.TryParse(xmlVal, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, v) AndAlso v > 0 Then
            If tree.SearchBarFontSize <> v Then tree.SearchBarFontSize = v
            TreeLogger.Debug(Space(5) & $"SearchBarFontSize xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_SearchClearButton(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("SearchClearButton") Is Nothing Then Exit Sub
        tree.MarkSearchConfigured()
        Dim xmlVal As String = cfg.Attributes("SearchClearButton").Value
        Dim v As Integer = If(tree.SearchClearButton, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.SearchClearButton <> nv Then tree.SearchClearButton = nv
            TreeLogger.Debug(Space(5) & $"SearchClearButton xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_ScrollBarTheme(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("ScrollBarTheme") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("ScrollBarTheme").Value
        Dim v As Integer = 0
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As en_ScrollBarTheme = CType(v, en_ScrollBarTheme)
            If tree.ScrollBarTheme <> nv Then tree.ScrollBarTheme = nv
            TreeLogger.Debug(Space(5) & $"ScrollBarTheme xml='{xmlVal}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_TooltipShow(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("TooltipShow") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("TooltipShow").Value
        Dim v As Integer = If(tree.TooltipShow, 1, 0)
        If Integer.TryParse(xmlVal, v) Then
            Dim nv As Boolean = (v = 1)
            If tree.TooltipShow <> nv Then tree.TooltipShow = nv
            TreeLogger.Debug(Space(5) & $"TooltipShow xml='{xmlVal}' control='{tree.TooltipShow}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_TooltipBackColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("TooltipBackColor") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("TooltipBackColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.TooltipBackColor)
        If tree.TooltipBackColor <> c Then
            tree.TooltipBackColor = c
            TreeLogger.Debug(Space(5) & $"TooltipBackColor xml='{xmlVal}' control='{tree.TooltipBackColor}'", "AplicareConfigurare")
        End If
    End Sub

    Friend Shared Sub Apply_TooltipForeColor(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("TooltipForeColor") Is Nothing Then Exit Sub
        Dim xmlVal As String = cfg.Attributes("TooltipForeColor").Value
        Dim c As Color = AdvancedTreeControl.ParseColor(xmlVal, tree.TooltipForeColor)
        If tree.TooltipForeColor <> c Then
            tree.TooltipForeColor = c
            TreeLogger.Debug(Space(5) & $"TooltipForeColor xml='{xmlVal}' control='{tree.TooltipForeColor}'", "AplicareConfigurare")
        End If
    End Sub

    ''' <summary>
    ''' Citeste blocul &lt;Columns&gt; din XML si populeaza lista de definitii de coloane.
    ''' Seteaza treeListView = True daca exista cel putin o coloana.
    ''' </summary>
    Friend Shared Sub Apply_Columns(xmlDoc As XmlDocument,
                                 ByRef columns As List(Of ColumnDef),
                                 ByRef treeListView As Boolean)
        Try
            columns.Clear()
            treeListView = False

            Dim colNodes As XmlNodeList = xmlDoc.SelectNodes("//Columns/Column")
            If colNodes Is Nothing OrElse colNodes.Count = 0 Then Return

            For Each cn As XmlNode In colNodes
                Try
                    Dim cd As New ColumnDef With {
                        .Name = If(cn.Attributes("Name")?.Value, "")
                    }
                    cd.Header = If(cn.Attributes("Header")?.Value, cd.Name)

                    Dim wVal As Integer = 100
                    If Integer.TryParse(cn.Attributes("Width")?.Value, wVal) Then
                        cd.Width = wVal
                    End If

                    ' ── ColType (integer) ─────────────────────────────────────────
                    Dim tVal As Integer = 0
                    If Integer.TryParse(cn.Attributes("Type")?.Value, tVal) Then
                        cd.ColType = CType(tVal, En_ColType)
                    End If

                    ' ── Align (integer) ───────────────────────────────────────────
                    Dim aVal As Integer = 0
                    If Integer.TryParse(cn.Attributes("Align")?.Value, aVal) Then
                        cd.Align = CType(aVal, En_ColAlign)
                    End If

                    cd.Format = If(cn.Attributes("Format")?.Value, "")

                    ' ── header styling ────────────────────────────────────────────
                    Dim bgStr As String = If(cn.Attributes("HdrBackColor")?.Value, "")
                    If Not String.IsNullOrEmpty(bgStr) Then
                        cd.HeaderBackColor = AdvancedTreeControl.ParseColor(bgStr, Color.Empty)
                    End If

                    Dim fgStr As String = If(cn.Attributes("HdrForeColor")?.Value, "")
                    If Not String.IsNullOrEmpty(fgStr) Then
                        cd.HeaderForeColor = AdvancedTreeControl.ParseColor(fgStr, Color.Empty)
                    End If

                    Dim bv As Integer = 0
                    If Integer.TryParse(cn.Attributes("HdrBold")?.Value, bv) Then cd.HeaderBold = (bv = 1)
                    If Integer.TryParse(cn.Attributes("HdrItalic")?.Value, bv) Then cd.HeaderItalic = (bv = 1)
                    If Integer.TryParse(cn.Attributes("HdrUnderline")?.Value, bv) Then cd.HeaderUnderline = (bv = 1)

                    ' HeaderAlign: absent din XML → Inherit (-1)
                    Dim haVal As Integer = CInt(En_ColAlign.ColAlign_Inherit)
                    If Integer.TryParse(cn.Attributes("HdrAlign")?.Value, haVal) Then
                        cd.HeaderAlign = CType(haVal, En_ColAlign)
                    Else
                        cd.HeaderAlign = En_ColAlign.ColAlign_Inherit
                    End If

                    If Not String.IsNullOrEmpty(cd.Name) Then columns.Add(cd)

                    TreeLogger.Debug(Space(5) & $"Column: Name='{cd.Name}' Header='{cd.Header}' Width={cd.Width} Type={cd.ColType} Align={cd.Align} Format='{cd.Format}' BackColor={cd.HeaderBackColor} ForeColor={cd.HeaderForeColor} Bold={cd.HeaderBold} Italic={cd.HeaderItalic} Underline={cd.HeaderUnderline} HeaderAlign={cd.HeaderAlign}", "Apply_Columns/Column")

                Catch ex As Exception
                    TreeLogger.Ex(ex, "Apply_Columns/Column")
                End Try
            Next

            treeListView = (columns.Count > 0)
        Catch ex As Exception
            TreeLogger.Ex(ex, "Apply_Columns")
        End Try
    End Sub

    Friend Shared Sub Apply_TreeListViewEnabled(cfg As XmlNode, tree As AdvancedTreeControl)
        If cfg.Attributes("TreeListViewEnabled") Is Nothing Then
            TreeLogger.Debug(Space(5) & "TreeListViewEnabled attribute not found in XML", "AplicareConfigurare")
            Exit Sub
        End If
        Dim xmlVal As String = cfg.Attributes("TreeListViewEnabled").Value
        Dim v As Integer = xmlVal
        tree.TreeListView = v
        TreeLogger.Debug(Space(5) & $"TreeListViewEnabled={tree.TreeListView}", "AplicareConfigurare")
    End Sub
End Class