Imports System.Globalization
Imports System.Xml

Partial Public Class Tree

    Private Function AddXmlConfigToTree(cfg As XmlNode, Optional Reload As Boolean = False) As Boolean
        Try

            Dim sw As New Stopwatch()
            sw.Start()

            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            Apply_treeID(cfg, Reload)

            Apply_BackColor(cfg)
            Apply_BorderColor(cfg)
            Apply_ForeColor(cfg)
            Apply_HoverBackColor(cfg)
            Apply_SelectedBackColor(cfg)
            Apply_SelectedBorderColor(cfg)
            Apply_LineColor(cfg)

            Apply_RadioButtonLevel(cfg)
            Apply_CheckBoxes(cfg)

            Apply_Font(cfg, culture)
            Apply_FontName(cfg)
            Apply_FontSize(cfg)

            Apply_ItemHeight(cfg)

            Apply_HasNodeIcons(cfg)

            Apply_PopupTree(cfg)
            Apply_PopupGraceMs(cfg)

            Apply_LeftIconHeight(cfg)
            Apply_RightIconHeight(cfg)

            Apply_ShowRightIconOnHover(cfg)

            Apply_CheckboxSize(cfg)

            Apply_RightClickFunc(cfg)

            Apply_Indent(cfg)
            Apply_ExpanderSize(cfg)

            Apply_RootExpander(cfg)

            Apply_ReRaiseClickOnSameNode(cfg)
            Apply_RaiseLeftClickOnRightClick(cfg)

            Apply_ToolTipDelayMs(cfg)
            Apply_TooltipAutoHideMs(cfg)

            Apply_SelectedNodeId(cfg)

            Apply_LeftTextWidth(cfg)
            Apply_RightTextWidth(cfg)

            Apply_HeaderVisible(cfg)
            Apply_HeaderHeight(cfg)
            Apply_HeaderCaption(cfg)

            Apply_HeaderIconSize(cfg)

            Apply_HeaderBackColor(cfg)
            Apply_HeaderForeColor(cfg)

            Apply_HeaderLeftIcon(cfg)
            Apply_HeaderRightIcon(cfg)
            Apply_HeaderSearchIcon(cfg)

            Apply_SearchType(cfg)
            Apply_SearchIn(cfg)
            Apply_SearchMode(cfg)

            Apply_SearchDropdownHeight(cfg)

            TreeLogger.Info($"Configurare aplicată cu succes în {sw.ElapsedMilliseconds}ms", "AplicareConfigurare")

            Return True

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        End Try
    End Function

    Private Sub Apply_treeID(cfg As XmlNode, Reload As Boolean)

        If cfg.Attributes("treeID") Is Nothing Then
            TreeLogger.Err("EROARE: Atributul 'treeId' este obligatoriu în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If

        Dim tId As String = cfg.Attributes("treeID").Value

        If String.IsNullOrEmpty(tId) Then
            TreeLogger.Err("EROARE: Atributul 'treeId' nu poate fi gol în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End If

        If Reload Then

            If MyTree.treeID <> tId Then
                TreeLogger.Err("EROARE: La reîncărcare, atributul 'treeId' nu corespunde cu cel inițial.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If

        Else

            If MyTree.treeID <> tId Then
                MyTree.treeID = tId
            End If

        End If

        TreeLogger.Debug(Space(5) & $"treeID xml='{tId}' control='{MyTree.treeID}'", "AplicareConfigurare")

    End Sub

    Private Sub Apply_BackColor(cfg As XmlNode)

        If cfg.Attributes("BackColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("BackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.BackColor <> c Then MyTree.BackColor = c
            If Me.BackColor <> c Then Me.BackColor = c

            TreeLogger.Debug(Space(5) & $"BackColor xml='{xmlVal}' control='{MyTree.BackColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_BorderColor(cfg As XmlNode)

        If cfg.Attributes("BorderColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("BorderColor").Value

            If xmlVal.StartsWith("#"c) Then

                Dim c As Color = ColorTranslator.FromHtml(xmlVal)

                If MyTree.BorderColor <> c Then
                    MyTree.BorderColor = c
                End If

                TreeLogger.Debug(Space(5) & $"BorderColor xml='{xmlVal}' control='{MyTree.BorderColor}'", "AplicareConfigurare")

            End If

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_ForeColor(cfg As XmlNode)

        If cfg.Attributes("ForeColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("ForeColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.ForeColor <> c Then
                MyTree.ForeColor = c
            End If

            TreeLogger.Debug(Space(5) & $"ForeColor xml='{xmlVal}' control='{MyTree.ForeColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_HoverBackColor(cfg As XmlNode)

        If cfg.Attributes("HoverBackColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("HoverBackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.HoverBackColor <> c Then
                MyTree.HoverBackColor = c
            End If

            TreeLogger.Debug(Space(5) & $"HoverBackColor xml='{xmlVal}' control='{MyTree.HoverBackColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_SelectedBackColor(cfg As XmlNode)

        If cfg.Attributes("SelectedBackColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("SelectedBackColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.SelectedBackColor <> c Then
                MyTree.SelectedBackColor = c
            End If

            TreeLogger.Debug(Space(5) & $"SelectedBackColor xml='{xmlVal}' control='{MyTree.SelectedBackColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_SelectedBorderColor(cfg As XmlNode)

        If cfg.Attributes("SelectedBorderColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("SelectedBorderColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.SelectedBorderColor <> c Then
                MyTree.SelectedBorderColor = c
            End If

            TreeLogger.Debug(Space(5) & $"SelectedBorderColor xml='{xmlVal}' control='{MyTree.SelectedBorderColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_LineColor(cfg As XmlNode)

        If cfg.Attributes("LineColor") Is Nothing Then Exit Sub

        Try

            Dim xmlVal As String = cfg.Attributes("LineColor").Value
            Dim c As Color = ColorTranslator.FromHtml(xmlVal)

            If MyTree.LineColor <> c Then
                MyTree.LineColor = c
            End If

            TreeLogger.Debug(Space(5) & $"LineColor xml='{xmlVal}' control='{MyTree.LineColor}'", "AplicareConfigurare")

        Catch ex As Exception

            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub Apply_RadioButtonLevel(cfg As XmlNode)

        If cfg.Attributes("RadioButtonLevel") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RadioButtonLevel").Value
        Dim v As Integer = MyTree.RadioButtonLevel

        If Integer.TryParse(xmlVal, v) Then

            If MyTree.RadioButtonLevel <> v Then
                MyTree.RadioButtonLevel = v
            End If

            TreeLogger.Debug(Space(5) & $"RadioButtonLevel xml='{xmlVal}' control='{MyTree.RadioButtonLevel}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_CheckBoxes(cfg As XmlNode)

        If MyTree.RadioButtonLevel <> -1 Then Exit Sub
        If cfg.Attributes("CheckBoxes") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("CheckBoxes").Value
        Dim v As Integer = If(MyTree.CheckBoxes, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.CheckBoxes <> nv Then
                MyTree.CheckBoxes = nv
            End If

            TreeLogger.Debug(Space(5) & $"CheckBoxes xml='{xmlVal}' control='{MyTree.CheckBoxes}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_Font(cfg As XmlNode, culture As CultureInfo)

        Dim curFontName As String = If(MyTree.Font IsNot Nothing, MyTree.Font.Name, "Segoe UI")
        Dim curFontSize As Single = If(MyTree.Font IsNot Nothing, MyTree.Font.Size, 9.0F)

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
                (MyTree.Font Is Nothing) OrElse
                (MyTree.Font.Name <> fName) OrElse
                (Math.Abs(MyTree.Font.Size - fSize) > 0.001F)

            If needSet Then
                MyTree.Font = New Font(fName, fSize)
            End If

            TreeLogger.Debug(Space(5) &
                $"Font xmlName='{xmlFontName}' xmlSize='{xmlFontSize}' control='{MyTree.Font.Name} {MyTree.Font.Size}pt'",
                "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_FontName(cfg As XmlNode)

        If cfg.Attributes("FontName") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("FontName").Value

        If MyTree.FontName <> xmlVal Then
            MyTree.FontName = xmlVal
        End If

        TreeLogger.Debug(Space(5) & $"FontName xml='{xmlVal}' control='{MyTree.FontName}'", "AplicareConfigurare")

    End Sub

    Private Sub Apply_FontSize(cfg As XmlNode)

        If cfg.Attributes("FontSize") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("FontSize").Value
        Dim v As Integer = MyTree.FontSize

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.FontSize <> v Then
                MyTree.FontSize = v
            End If

            TreeLogger.Debug(Space(5) & $"FontSize xml='{xmlVal}' control='{MyTree.FontSize}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_ItemHeight(cfg As XmlNode)

        If cfg.Attributes("ItemHeight") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("ItemHeight").Value
        Dim v As Integer = MyTree.ItemHeight

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.ItemHeight <> v Then
                MyTree.ItemHeight = v
            End If

            TreeLogger.Debug(Space(5) & $"ItemHeight xml='{xmlVal}' control='{MyTree.ItemHeight}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_HasNodeIcons(cfg As XmlNode)

        If cfg.Attributes("HasNodeIcons") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HasNodeIcons").Value
        Dim v As Integer = If(MyTree.HasNodeIcons, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.HasNodeIcons <> nv Then
                MyTree.HasNodeIcons = nv
            End If

            TreeLogger.Debug(Space(5) & $"HasNodeIcons xml='{xmlVal}' control='{MyTree.HasNodeIcons}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_PopupTree(cfg As XmlNode)

        If cfg.Attributes("PopupTree") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("PopupTree").Value
        Dim v As Integer = If(MyTree.IsPopupTree, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.IsPopupTree <> nv Then
                MyTree.IsPopupTree = nv
            End If

            TreeLogger.Debug(Space(5) & $"PopupTree xml='{xmlVal}' control='{MyTree.IsPopupTree}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_PopupGraceMs(cfg As XmlNode)

        If cfg.Attributes("PopupGraceMs") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("PopupGraceMs").Value
        Dim v As Integer = MyTree.PopupGraceMs

        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then

            If MyTree.PopupGraceMs <> v Then
                MyTree.PopupGraceMs = v
            End If

            TreeLogger.Debug(Space(5) & $"PopupGraceMs xml='{xmlVal}' control='{MyTree.PopupGraceMs}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_LeftIconHeight(cfg As XmlNode)

        If cfg.Attributes("LeftIconHeight") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("LeftIconHeight").Value
        Dim v As Integer = If(MyTree.LeftIconSize.Height > 0, MyTree.LeftIconSize.Height, 16)

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            Dim ns As New Size(v, v)

            If MyTree.LeftIconSize <> ns Then
                MyTree.LeftIconSize = ns
            End If

            TreeLogger.Debug(Space(5) & $"LeftIconHeight xml='{xmlVal}' control='{MyTree.LeftIconSize}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_RightIconHeight(cfg As XmlNode)

        If cfg.Attributes("RightIconHeight") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RightIconHeight").Value
        Dim v As Integer = If(MyTree.RightIconSize.Height > 0, MyTree.RightIconSize.Height, 16)

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            Dim ns As New Size(v, v)

            If MyTree.RightIconSize <> ns Then
                MyTree.RightIconSize = ns
            End If

            TreeLogger.Debug(Space(5) & $"RightIconHeight xml='{xmlVal}' control='{MyTree.RightIconSize}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_ShowRightIconOnHover(cfg As XmlNode)

        If cfg.Attributes("ShowRightIconOnHover") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("ShowRightIconOnHover").Value
        Dim v As Integer = If(MyTree.ShowRightIconOnHover, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.ShowRightIconOnHover <> nv Then
                MyTree.ShowRightIconOnHover = nv
            End If

            TreeLogger.Debug(Space(5) & $"ShowRightIconOnHover xml='{xmlVal}' control='{MyTree.ShowRightIconOnHover}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_CheckboxSize(cfg As XmlNode)

        If cfg.Attributes("CheckboxSize") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("CheckboxSize").Value
        Dim v As Integer = MyTree.CheckBoxSize

        If v <= 0 Then v = 16

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.CheckBoxSize <> v Then
                MyTree.CheckBoxSize = v
            End If

            TreeLogger.Debug(Space(5) & $"CheckboxSize xml='{xmlVal}' control='{MyTree.CheckBoxSize}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_RightClickFunc(cfg As XmlNode)

        If cfg.Attributes("RightClickFunc") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RightClickFunc").Value

        If MyTree.RightClickFunction <> xmlVal Then
            MyTree.RightClickFunction = xmlVal
        End If

        TreeLogger.Debug(Space(5) & $"RightClickFunc xml='{xmlVal}' control='{MyTree.RightClickFunction}'", "AplicareConfigurare")

    End Sub

    Private Sub Apply_Indent(cfg As XmlNode)

        If cfg.Attributes("Indent") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("Indent").Value
        Dim v As Integer = MyTree.Indent

        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then

            If MyTree.Indent <> v Then
                MyTree.Indent = v
            End If

            TreeLogger.Debug(Space(5) & $"Indent xml='{xmlVal}' control='{MyTree.Indent}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_ExpanderSize(cfg As XmlNode)

        If cfg.Attributes("ExpanderSize") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("ExpanderSize").Value
        Dim v As Integer = MyTree.ExpanderSize

        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then

            If MyTree.ExpanderSize <> v Then
                MyTree.ExpanderSize = v
            End If

            TreeLogger.Debug(Space(5) & $"ExpanderSize xml='{xmlVal}' control='{MyTree.ExpanderSize}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_RootExpander(cfg As XmlNode)

        If cfg.Attributes("RootExpander") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RootExpander").Value
        Dim v As Integer = If(MyTree.RootExpander, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.RootExpander <> nv Then
                MyTree.RootExpander = nv
            End If

            TreeLogger.Debug(Space(5) & $"RootExpander xml='{xmlVal}' control='{MyTree.RootExpander}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_ReRaiseClickOnSameNode(cfg As XmlNode)

        If cfg.Attributes("ReRaiseClickOnSameNode") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("ReRaiseClickOnSameNode").Value
        Dim v As Integer = If(MyTree.ReRaiseClickOnSameNode, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.ReRaiseClickOnSameNode <> nv Then
                MyTree.ReRaiseClickOnSameNode = nv
            End If

            TreeLogger.Debug(Space(5) & $"ReRaiseClickOnSameNode xml='{xmlVal}' control='{MyTree.ReRaiseClickOnSameNode}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_RaiseLeftClickOnRightClick(cfg As XmlNode)

        If cfg.Attributes("RaiseLeftClickOnRightClick") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RaiseLeftClickOnRightClick").Value
        Dim v As Integer = If(MyTree.RaiseLeftClickOnRightClick, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.RaiseLeftClickOnRightClick <> nv Then
                MyTree.RaiseLeftClickOnRightClick = nv
            End If

            TreeLogger.Debug(Space(5) & $"RaiseLeftClickOnRightClick xml='{xmlVal}' control='{MyTree.RaiseLeftClickOnRightClick}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_ToolTipDelayMs(cfg As XmlNode)

        If cfg.Attributes("ToolTipDelayMs") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("ToolTipDelayMs").Value
        Dim v As Integer = MyTree.TooltipDelayMs

        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then

            If MyTree.TooltipDelayMs <> v Then
                MyTree.TooltipDelayMs = v
            End If

            TreeLogger.Debug(Space(5) & $"ToolTipDelayMs xml='{xmlVal}' control='{MyTree.TooltipDelayMs}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_TooltipAutoHideMs(cfg As XmlNode)

        If cfg.Attributes("TooltipAutoHideMs") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("TooltipAutoHideMs").Value
        Dim v As Integer = MyTree.AutoHideTooltipMs

        If Integer.TryParse(xmlVal, v) AndAlso v >= 0 Then

            If MyTree.AutoHideTooltipMs <> v Then
                MyTree.AutoHideTooltipMs = v
            End If

            TreeLogger.Debug(Space(5) & $"TooltipAutoHideMs xml='{xmlVal}' control='{MyTree.AutoHideTooltipMs}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_SelectedNodeId(cfg As XmlNode)

        _pendingSelectedNodeId = String.Empty

        If cfg.Attributes("SelectedNodeId") Is Nothing Then Exit Sub

        _pendingSelectedNodeId = cfg.Attributes("SelectedNodeId").Value

        TreeLogger.Debug(Space(5) & $"SelectedNodeId xml='{_pendingSelectedNodeId}'", "AplicareConfigurare")

    End Sub

    Private Sub Apply_LeftTextWidth(cfg As XmlNode)

        If cfg.Attributes("LeftTextWidth") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("LeftTextWidth").Value
        Dim v As Integer = MyTree.LeftTextWidth

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.LeftTextWidth <> v Then
                MyTree.LeftTextWidth = v
            End If

            TreeLogger.Debug(Space(5) & $"LeftTextWidth xml='{xmlVal}' control='{MyTree.LeftTextWidth}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_RightTextWidth(cfg As XmlNode)

        If cfg.Attributes("RightTextWidth") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("RightTextWidth").Value
        Dim v As Integer = MyTree.RightTextWidth

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.RightTextWidth <> v Then
                MyTree.RightTextWidth = v
            End If

            TreeLogger.Debug(Space(5) & $"RightTextWidth xml='{xmlVal}' control='{MyTree.RightTextWidth}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_HeaderVisible(cfg As XmlNode)

        If cfg.Attributes("HeaderVisible") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderVisible").Value
        Dim v As Integer = If(MyTree.HeaderVisible, 1, 0)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv As Boolean = (v = 1)

            If MyTree.HeaderVisible <> nv Then
                MyTree.HeaderVisible = nv
            End If

            TreeLogger.Debug(Space(5) & $"HeaderVisible xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_HeaderHeight(cfg As XmlNode)

        If cfg.Attributes("HeaderHeight") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderHeight").Value
        Dim v As Integer = MyTree.HeaderHeight

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.HeaderHeight <> v Then
                MyTree.HeaderHeight = v
            End If

            TreeLogger.Debug(Space(5) & $"HeaderHeight xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_HeaderCaption(cfg As XmlNode)

        If cfg.Attributes("HeaderCaption") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderCaption").Value

        If MyTree.HeaderCaption <> xmlVal Then
            MyTree.HeaderCaption = xmlVal
        End If

        TreeLogger.Debug(Space(5) & $"HeaderCaption xml='{xmlVal}'", "AplicareConfigurare")

    End Sub

    Private Sub Apply_HeaderIconSize(cfg As XmlNode)

        If cfg.Attributes("HeaderIconSize") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderIconSize").Value
        Dim v As Integer = MyTree.HeaderIconSize.Width

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            Dim ns As New Size(v, v)

            If MyTree.HeaderIconSize <> ns Then
                MyTree.HeaderIconSize = ns
            End If

            TreeLogger.Debug(Space(5) & $"HeaderIconSize xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_HeaderBackColor(cfg As XmlNode)

        If cfg.Attributes("HeaderBackColor") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderBackColor").Value
        Dim c = AdvancedTreeControl.ParseColor(xmlVal, MyTree.HeaderBackColor)

        If MyTree.HeaderBackColor <> c Then
            MyTree.HeaderBackColor = c
        End If

    End Sub

    Private Sub Apply_HeaderForeColor(cfg As XmlNode)

        If cfg.Attributes("HeaderForeColor") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("HeaderForeColor").Value
        Dim c = AdvancedTreeControl.ParseColor(xmlVal, MyTree.HeaderForeColor)

        If MyTree.HeaderForeColor <> c Then
            MyTree.HeaderForeColor = c
        End If

    End Sub

    Private Sub Apply_HeaderLeftIcon(cfg As XmlNode)

        If cfg.Attributes("HeaderLeftIcon") Is Nothing Then Exit Sub
        MyTree.HeaderLeftIconKey = cfg.Attributes("HeaderLeftIcon").Value

    End Sub

    Private Sub Apply_HeaderRightIcon(cfg As XmlNode)

        If cfg.Attributes("HeaderRightIcon") Is Nothing Then Exit Sub
        MyTree.HeaderRightIconKey = cfg.Attributes("HeaderRightIcon").Value

    End Sub

    Private Sub Apply_HeaderSearchIcon(cfg As XmlNode)

        If cfg.Attributes("HeaderSearchIcon") Is Nothing Then Exit Sub
        MyTree.HeaderSearchIconKey = cfg.Attributes("HeaderSearchIcon").Value

    End Sub

    Private Sub Apply_SearchType(cfg As XmlNode)

        If cfg.Attributes("SearchType") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("SearchType").Value
        Dim v As Integer = CInt(MyTree.SearchType)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv = CType(v, AdvancedTreeControl.en_Tree_SearchType)

            If MyTree.SearchType <> nv Then
                MyTree.SearchType = nv
            End If

            TreeLogger.Debug(Space(5) & $"SearchType xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_SearchIn(cfg As XmlNode)

        If cfg.Attributes("SearchIn") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("SearchIn").Value
        Dim v As Integer = CInt(MyTree.SearchIn)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv = CType(v, AdvancedTreeControl.en_Tree_SearchIn)

            If MyTree.SearchIn <> nv Then
                MyTree.SearchIn = nv
            End If

            TreeLogger.Debug(Space(5) & $"SearchIn xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_SearchMode(cfg As XmlNode)

        If cfg.Attributes("SearchMode") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("SearchMode").Value
        Dim v As Integer = CInt(MyTree.SearchMode)

        If Integer.TryParse(xmlVal, v) Then

            Dim nv = CType(v, AdvancedTreeControl.en_Tree_SearchMode)

            If MyTree.SearchMode <> nv Then
                MyTree.SearchMode = nv
            End If

            TreeLogger.Debug(Space(5) & $"SearchMode xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub

    Private Sub Apply_SearchDropdownHeight(cfg As XmlNode)

        If cfg.Attributes("SearchDropdownHeight") Is Nothing Then Exit Sub

        Dim xmlVal As String = cfg.Attributes("SearchDropdownHeight").Value
        Dim v As Integer = MyTree.SearchDropdownHeight

        If Integer.TryParse(xmlVal, v) AndAlso v > 0 Then

            If MyTree.SearchDropdownHeight <> v Then
                MyTree.SearchDropdownHeight = v
            End If

            TreeLogger.Debug(Space(5) & $"SearchDropdownHeight xml='{xmlVal}'", "AplicareConfigurare")

        End If

    End Sub
End Class