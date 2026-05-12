Imports System.Globalization
Imports System.IO
Imports System.Xml

Partial Public Class Tree
    ' =============================================================
    ' ÎNCĂRCARE XML + CONFIGURARE (CORECTAT)
    ' =============================================================
    Private Function LoadXmlData(filePath As String) As Boolean
        If Not File.Exists(filePath) Then Return False
        Try
            Dim xDoc As New XmlDocument()
            xDoc.Load(filePath)

            TreeLogger.Info($"Încep încărcare XML din fișier: {filePath}", "LoadXmlData")
            MyTree.SuspendLayout()
            MyTree.Items.Clear()
            _imageCache.Clear()

            ' 1. CONFIGURARE 
            Dim configNode As XmlNode = xDoc.SelectSingleNode("/Tree/Config")
            If configNode IsNot Nothing Then
                AddXmlConfigToTree(configNode)
            End If

            ' 2. INCARCARE IMAGINI
            Dim imgListNode As XmlNode = xDoc.SelectSingleNode("/Tree/Images")
            If imgListNode IsNot Nothing Then
                LoadImagesToCache(imgListNode)
            End If

            MyTree.ResolveHeaderIcons(_imageCache)

            ' 3. POPULARE NODURI
            Dim nodesRoot = xDoc.SelectNodes("/Tree/Nodes/Node")

            For Each xNode As XmlNode In nodesRoot
                AddXmlNodeToTree(xNode, Nothing)
            Next

            ApplyPendingSelectedNode()

            MyTree.Invalidate()
            Return True

        Catch ex As Exception
            TreeLogger.Ex(ex, "LoadXmlDataFromString", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            MyTree.ResumeLayout()
        End Try
    End Function

    Private Function ReLoadXmlData(filePath As String) As Boolean
        If Not File.Exists(filePath) Then Return False

        Try
            Dim xDoc As New XmlDocument()
            Dim cNodeKey = MyTree.SelectedNode?.Key

            xDoc.Load(filePath)

            MyTree.SuspendLayout()

            ' 1. CONFIGURARE 
            Dim configNode As XmlNode = xDoc.SelectSingleNode("/Tree/Config")
            If configNode IsNot Nothing Then
                If Not AddXmlConfigToTree(configNode) Then
                    Application.Exit()
                End If
            End If

            ' 2. INCARCARE IMAGINI
            Dim imgListNode As XmlNode = xDoc.SelectSingleNode("/Tree/Images")
            If imgListNode IsNot Nothing Then
                LoadImagesToCache(imgListNode)
            End If

            MyTree.ResolveHeaderIcons(_imageCache)

            ' 3. POPULARE NODURI
            Dim nodesRoot = xDoc.SelectNodes("/Tree/Nodes/Node")
            If nodesRoot IsNot Nothing AndAlso nodesRoot.Count > 0 Then
                MyTree.Items.Clear()

                For Each xNode As XmlNode In nodesRoot
                    AddXmlNodeToTree(xNode, Nothing)
                Next
            Else
                MyTree.Refresh()
            End If

            'scroll la nodul activ dacă există
            ' SelectedNodeId din XML are prioritate; altfel restaurăm selecția anterioară
            If String.IsNullOrEmpty(_pendingSelectedNodeId) AndAlso Not String.IsNullOrEmpty(cNodeKey) Then
                _pendingSelectedNodeId = cNodeKey
            End If
            ApplyPendingSelectedNode()

            MyTree.Invalidate()
            Return True

        Catch ex As Exception
            TreeLogger.Ex(ex, "LoadXmlDataFromString", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            MyTree.ResumeLayout()
        End Try
    End Function

    Private Function AddXmlConfigToTree(cfg As XmlNode, Optional Reload As Boolean = False) As Boolean
        Try
            Dim sw As New Stopwatch()
            sw.Start()

            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            Apply_treeID(cfg, Reload) ' rămâne în Tree

            TreeXmlAppliers.Apply_BackColor(cfg, MyTree, Me)
            TreeXmlAppliers.Apply_BorderColor(cfg, MyTree)
            TreeXmlAppliers.Apply_ForeColor(cfg, MyTree)
            TreeXmlAppliers.Apply_HoverBackColor(cfg, MyTree)
            TreeXmlAppliers.Apply_SelectedBackColor(cfg, MyTree)
            TreeXmlAppliers.Apply_SelectedBorderColor(cfg, MyTree)
            TreeXmlAppliers.Apply_LineColor(cfg, MyTree)

            TreeXmlAppliers.Apply_RadioButtonLevel(cfg, MyTree)
            TreeXmlAppliers.Apply_CheckBoxes(cfg, MyTree)

            TreeXmlAppliers.Apply_Font(cfg, MyTree, culture)
            TreeXmlAppliers.Apply_FontName(cfg, MyTree)
            TreeXmlAppliers.Apply_FontSize(cfg, MyTree)

            TreeXmlAppliers.Apply_ItemHeight(cfg, MyTree)

            TreeXmlAppliers.Apply_HasNodeIcons(cfg, MyTree)

            TreeXmlAppliers.Apply_PopupTree(cfg, MyTree)
            TreeXmlAppliers.Apply_PopupGraceMs(cfg, MyTree)

            TreeXmlAppliers.Apply_LeftIconHeight(cfg, MyTree)
            TreeXmlAppliers.Apply_RightIconHeight(cfg, MyTree)

            TreeXmlAppliers.Apply_ShowRightIconOnHover(cfg, MyTree)

            TreeXmlAppliers.Apply_CheckboxSize(cfg, MyTree)

            TreeXmlAppliers.Apply_RightClickFunc(cfg, MyTree)

            TreeXmlAppliers.Apply_Indent(cfg, MyTree)
            TreeXmlAppliers.Apply_ExpanderSize(cfg, MyTree)

            TreeXmlAppliers.Apply_RootExpander(cfg, MyTree)

            TreeXmlAppliers.Apply_ReRaiseClickOnSameNode(cfg, MyTree)
            TreeXmlAppliers.Apply_RaiseLeftClickOnRightClick(cfg, MyTree)

            TreeXmlAppliers.Apply_ToolTipDelayMs(cfg, MyTree)
            TreeXmlAppliers.Apply_TooltipAutoHideMs(cfg, MyTree)

            TreeXmlAppliers.Apply_SelectedNodeId(cfg, _pendingSelectedNodeId)

            TreeXmlAppliers.Apply_LeftTextWidth(cfg, MyTree)
            TreeXmlAppliers.Apply_RightTextWidth(cfg, MyTree)

            TreeXmlAppliers.Apply_HeaderVisible(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderHeight(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderCaption(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderIconSize(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderBackColor(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderForeColor(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderLeftIcon(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderRightIcon(cfg, MyTree)
            TreeXmlAppliers.Apply_HeaderSearchIcon(cfg, MyTree)

            TreeXmlAppliers.Apply_SearchType(cfg, MyTree)
            TreeXmlAppliers.Apply_SearchIn(cfg, MyTree)
            TreeXmlAppliers.Apply_SearchMode(cfg, MyTree)
            TreeXmlAppliers.Apply_SearchDropdownHeight(cfg, MyTree)

            TreeLogger.Info($"Configurare aplicată cu succes în {sw.ElapsedMilliseconds}ms", "AplicareConfigurare")
            Return True

        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub AddXmlNodeToTree(xNode As XmlNode, parentItem As AdvancedTreeControl.TreeItem)
        Try
            Dim nodeCaption As String = ""
            NodeXmlAppliers.Apply_NodeCaption(xNode, nodeCaption)

            Dim nodeKey As String = ""
            NodeXmlAppliers.Apply_NodeKey(xNode, nodeKey)

            Dim nodeTag As String = ""
            NodeXmlAppliers.Apply_NodeTag(xNode, nodeTag)

            Dim nodeIconNameClosed As String = ""
            Dim nodeIconNameOpen As String = ""
            Dim nodeIconRight As String = ""
            NodeXmlAppliers.Apply_NodeIconClosed(xNode, nodeIconNameClosed)
            NodeXmlAppliers.Apply_NodeIconOpen(xNode, nodeIconNameOpen)
            NodeXmlAppliers.Apply_NodeIconRight(xNode, nodeIconRight)
            NodeXmlAppliers.Apply_NodeMirrorIcons(nodeIconNameClosed, nodeIconNameOpen)

            Dim iconImgClosed As Image = Nothing
            Dim iconImgOpen As Image = Nothing
            Dim iconImgRight As Image = Nothing
            NodeXmlAppliers.Apply_NodeClosedImage(nodeIconNameClosed, iconImgClosed, _imageCache)
            NodeXmlAppliers.Apply_NodeOpenImage(nodeIconNameOpen, iconImgOpen, _imageCache)
            NodeXmlAppliers.Apply_NodeRightImage(nodeIconRight, iconImgRight, _imageCache)

            Dim iconExpanded As Boolean = False
            NodeXmlAppliers.Apply_NodeExpanded(xNode, iconExpanded)

            Dim isLazy As Boolean = False
            NodeXmlAppliers.Apply_NodeLazyNode(xNode, isLazy)

            Dim newItem As AdvancedTreeControl.TreeItem =
                MyTree.AddItem(nodeKey, nodeCaption, parentItem, iconImgClosed, iconImgOpen, iconImgRight, nodeTag, iconExpanded, isLazy)

            If newItem Is Nothing Then Exit Sub

            If newItem.Key <> nodeKey Then newItem.Key = nodeKey

            NodeXmlAppliers.Apply_NodeTooltip(xNode, newItem)
            NodeXmlAppliers.Apply_NodeBold(xNode, newItem)
            NodeXmlAppliers.Apply_NodeItalic(xNode, newItem)
            NodeXmlAppliers.Apply_NodeHasCheckbox(xNode, newItem)
            NodeXmlAppliers.Apply_NodeShowRightIconOnHover(xNode, newItem)
            NodeXmlAppliers.Apply_NodeCheckState(xNode, newItem)
            NodeXmlAppliers.Apply_NodeForeColor(xNode, newItem)
            NodeXmlAppliers.Apply_NodeBackColor(xNode, newItem)

            For Each childNode As XmlNode In xNode.SelectNodes("Node")
                AddXmlNodeToTree(childNode, newItem)
            Next

        Catch ex As Exception
            TreeLogger.Ex(ex, "AddXmlNodeToTree", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadImagesToCache(imgRoot As XmlNode)
        Dim count As Integer = 0
        _imageCache.Clear()

        For Each imgNode As XmlNode In imgRoot.SelectNodes("Image")
            Try
                Dim key As String = imgNode.Attributes("Key").Value
                Dim b64 As String = imgNode.InnerText

                If Not String.IsNullOrEmpty(b64) Then
                    Dim bytes As Byte() = Convert.FromBase64String(b64)
                    Dim ms As New MemoryStream(bytes)
                    Dim bmp As Image = Image.FromStream(ms)

                    If _imageCache.TryAdd(key, bmp) Then
                        count += 1
                    End If
                End If
            Catch ex As Exception
                TreeLogger.Ex(ex, "LoadImagesToCache", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next

        TreeLogger.Debug($" Încărcat {count} imagini în cache.", "LoadImagesToCache")
    End Sub

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
            If MyTree.treeID <> tId Then MyTree.treeID = tId
        End If

        TreeLogger.Debug(Space(5) & $"treeID xml='{tId}' control='{MyTree.treeID}'", "AplicareConfigurare")

    End Sub

    Private Sub ApplyPendingSelectedNode()
        If String.IsNullOrEmpty(_pendingSelectedNodeId) Then Return

        Dim key As String = _pendingSelectedNodeId
        _pendingSelectedNodeId = String.Empty

        Dim foundNode As AdvancedTreeControl.TreeItem = Nothing
        For Each rootItem In MyTree.Items
            foundNode = FindNodeByIdRecursive(rootItem, key)
            If foundNode IsNot Nothing Then Exit For
        Next

        If foundNode Is Nothing Then
            TreeLogger.Debug($"SelectedNodeId '{key}' nu a fost găsit în arbore", "ApplyPendingSelectedNode")
            Return
        End If

        ' Expandăm părinții silențios
        Dim parent As AdvancedTreeControl.TreeItem = foundNode.Parent
        While parent IsNot Nothing
            parent.Expanded = True
            parent = parent.Parent
        End While

        MyTree.SelectedNode = foundNode
        ScrollToNode(foundNode)
        TreeLogger.Debug($"SelectedNodeId '{key}' aplicat silențios", "ApplyPendingSelectedNode")
    End Sub
End Class
