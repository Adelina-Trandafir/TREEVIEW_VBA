Imports System.Xml
Imports System.Drawing

Partial Public Class Tree

    ' =============================================================
    ' FUNCȚIA PRINCIPALĂ REFACTORIZATĂ
    ' =============================================================
    Private Sub AddXmlNodeToTree(xNode As XmlNode, parentItem As AdvancedTreeControl.TreeItem)
        Try
            ' --- Caption ---
            Dim nodeCaption As String = ""
            Apply_NodeCaption(xNode, nodeCaption)

            ' --- Key ---
            Dim nodeKey As String = ""
            Apply_NodeKey(xNode, nodeKey)

            ' --- Tag ---
            Dim nodeTag As String = ""
            Apply_NodeTag(xNode, nodeTag)

            ' --- IconClosed / IconOpen / IconRight ---
            Dim nodeIconNameClosed As String = ""
            Dim nodeIconNameOpen As String = ""
            Dim nodeIconRight As String = ""

            Apply_NodeIconClosed(xNode, nodeIconNameClosed)
            Apply_NodeIconOpen(xNode, nodeIconNameOpen)
            Apply_NodeIconRight(xNode, nodeIconRight)

            ' --- If only one of closed/open is specified, mirror it ---
            Apply_NodeMirrorIcons(nodeIconNameClosed, nodeIconNameOpen)

            ' --- Resolve images from cache ---
            Dim iconImgClosed As Image = Nothing
            Dim iconImgOpen As Image = Nothing
            Dim iconImgRight As Image = Nothing

            Apply_NodeClosedImage(nodeIconNameClosed, iconImgClosed)
            Apply_NodeOpenImage(nodeIconNameOpen, iconImgOpen)
            Apply_NodeRightImage(nodeIconRight, iconImgRight)

            ' --- Expanded ---
            Dim iconExpanded As Boolean = False ' default TreeItem
            Apply_NodeExpanded(xNode, iconExpanded)

            ' --- LazyNode ---
            Dim isLazy As Boolean = False ' default TreeItem
            Apply_NodeLazyNode(xNode, isLazy)

            ' --- AddItem ---
            Dim newItem As AdvancedTreeControl.TreeItem =
                MyTree.AddItem(nodeKey, nodeCaption, parentItem, iconImgClosed, iconImgOpen, iconImgRight, nodeTag, iconExpanded, isLazy)

            If newItem Is Nothing Then Exit Sub

            ' --- Key (redundant, but keep) ---
            If newItem.Key <> nodeKey Then newItem.Key = nodeKey

            ' --- Tooltip ---
            Apply_NodeTooltip(xNode, newItem)

            ' --- Bold ---
            Apply_NodeBold(xNode, newItem)

            ' --- Italic ---
            Apply_NodeItalic(xNode, newItem)

            ' --- HasCheckbox ---
            Apply_NodeHasCheckbox(xNode, newItem)

            ' --- ShowRightIconOnHover per nod ---
            Apply_NodeShowRightIconOnHover(xNode, newItem)

            ' --- CheckState ---
            Apply_NodeCheckState(xNode, newItem)

            ' --- ForeColor ---
            Apply_NodeForeColor(xNode, newItem)

            ' --- BackColor ---
            Apply_NodeBackColor(xNode, newItem)

            ' --- Children (recursive) ---
            For Each childNode As XmlNode In xNode.SelectNodes("Node")
                AddXmlNodeToTree(childNode, newItem)
            Next

        Catch ex As Exception
            TreeLogger.Ex(ex, "AddXmlNodeToTree", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =============================================================
    ' FUNCȚII INDIVIDUALE DE APLICARE A PROPRIETĂȚILOR
    ' =============================================================

    Private Sub Apply_NodeTooltip(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Tooltip") Is Nothing Then Exit Sub
        Dim xmlVal As String = xNode.Attributes("Tooltip").Value
        If newItem.Tooltip <> xmlVal Then
            newItem.Tooltip = xmlVal
        End If
    End Sub

    Private Sub Apply_NodeBold(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Bold") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Bold").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.Bold <> newVal Then
            newItem.Bold = newVal
        End If
    End Sub

    Private Sub Apply_NodeItalic(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Italic") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Italic").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.Italic <> newVal Then
            newItem.Italic = newVal
        End If
    End Sub

    Private Sub Apply_NodeHasCheckbox(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("HasCheckbox") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("HasCheckbox").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.HasCheckBox <> newVal Then
            newItem.HasCheckBox = newVal
        End If
    End Sub

    Private Sub Apply_NodeShowRightIconOnHover(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("ShowRightIconOnHover") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("ShowRightIconOnHover").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.ShowRightIconOnHover <> newVal Then
            newItem.ShowRightIconOnHover = newVal
        End If
    End Sub

    Private Sub Apply_NodeCheckState(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("CheckState") Is Nothing Then Exit Sub
        Dim xmlVal As String = xNode.Attributes("CheckState").Value.Trim()
        If newItem.CheckState <> xmlVal Then
            newItem.CheckState = xmlVal
        End If
    End Sub

    Private Sub Apply_NodeForeColor(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("ForeColor") Is Nothing Then Exit Sub
        Dim colorVal As String = xNode.Attributes("ForeColor").Value.Trim()
        If String.IsNullOrEmpty(colorVal) Then Exit Sub
        Try
            Dim newColor As Color
            If colorVal.StartsWith("#"c) Then
                newColor = ColorTranslator.FromHtml(colorVal)
            Else
                newColor = Color.FromName(colorVal)
            End If
            If newItem.NodeForeColor <> newColor Then
                newItem.NodeForeColor = newColor
            End If
        Catch
        End Try
    End Sub

    Private Sub Apply_NodeBackColor(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("BackColor") Is Nothing Then Exit Sub
        Dim colorVal As String = xNode.Attributes("BackColor").Value.Trim()
        If String.IsNullOrEmpty(colorVal) Then Exit Sub
        Try
            Dim newColor As Color
            If colorVal.StartsWith("#"c) Then
                newColor = ColorTranslator.FromHtml(colorVal)
            Else
                newColor = Color.FromName(colorVal)
            End If
            If newItem.NodeBackColor <> newColor Then
                newItem.NodeBackColor = newColor
            End If
        Catch
        End Try
    End Sub

    Private Sub Apply_NodeExpanded(xNode As XmlNode, ByRef iconExpanded As Boolean)
        If xNode.Attributes("Expanded") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Expanded").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If iconExpanded <> newVal Then
            iconExpanded = newVal
        End If
    End Sub

    Private Sub Apply_NodeLazyNode(xNode As XmlNode, ByRef isLazy As Boolean)
        If xNode.Attributes("LazyNode") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("LazyNode").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If isLazy <> newVal Then
            isLazy = newVal
        End If
    End Sub

    Private Sub Apply_NodeKey(xNode As XmlNode, ByRef nodeKey As String)
        If xNode.Attributes("Key") Is Nothing Then Exit Sub
        nodeKey = xNode.Attributes("Key").Value
    End Sub

    Private Sub Apply_NodeCaption(xNode As XmlNode, ByRef nodeCaption As String)
        If xNode.Attributes("Caption") Is Nothing Then Exit Sub
        nodeCaption = xNode.Attributes("Caption").Value
    End Sub

    Private Sub Apply_NodeTag(xNode As XmlNode, ByRef nodeTag As String)
        If xNode.Attributes("Tag") Is Nothing Then Exit Sub
        nodeTag = xNode.Attributes("Tag").Value
    End Sub

    Private Sub Apply_NodeIconClosed(xNode As XmlNode, ByRef nodeIconNameClosed As String)
        If xNode.Attributes("IconClosed") Is Nothing Then Exit Sub
        nodeIconNameClosed = xNode.Attributes("IconClosed").Value
    End Sub

    Private Sub Apply_NodeIconOpen(xNode As XmlNode, ByRef nodeIconNameOpen As String)
        If xNode.Attributes("IconOpen") Is Nothing Then Exit Sub
        nodeIconNameOpen = xNode.Attributes("IconOpen").Value
    End Sub

    Private Sub Apply_NodeIconRight(xNode As XmlNode, ByRef nodeIconRight As String)
        If xNode.Attributes("IconRight") Is Nothing Then Exit Sub
        nodeIconRight = xNode.Attributes("IconRight").Value
    End Sub

    Private Sub Apply_NodeMirrorIcons(ByRef nodeIconNameClosed As String, ByRef nodeIconNameOpen As String)
        If String.IsNullOrEmpty(nodeIconNameOpen) AndAlso nodeIconNameClosed <> "" Then
            nodeIconNameOpen = nodeIconNameClosed
        End If
        If nodeIconNameOpen <> "" AndAlso String.IsNullOrEmpty(nodeIconNameClosed) Then
            nodeIconNameClosed = nodeIconNameOpen
        End If
    End Sub

    Private Sub Apply_NodeClosedImage(ByVal nodeIconNameClosed As String, ByRef iconImgClosed As Image)
        If String.IsNullOrEmpty(nodeIconNameClosed) Then Exit Sub
        Dim value As Image = Nothing
        If _imageCache.TryGetValue(nodeIconNameClosed, value) Then
            iconImgClosed = value
        End If
    End Sub

    Private Sub Apply_NodeOpenImage(ByVal nodeIconNameOpen As String, ByRef iconImgOpen As Image)
        If String.IsNullOrEmpty(nodeIconNameOpen) Then Exit Sub
        Dim value As Image = Nothing
        If _imageCache.TryGetValue(nodeIconNameOpen, value) Then
            iconImgOpen = value
        End If
    End Sub

    Private Sub Apply_NodeRightImage(ByVal nodeIconRight As String, ByRef iconImgRight As Image)
        If String.IsNullOrEmpty(nodeIconRight) Then Exit Sub
        Dim value As Image = Nothing
        If _imageCache.TryGetValue(nodeIconRight, value) Then
            iconImgRight = value
        End If
    End Sub

End Class