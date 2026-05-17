Imports System.Xml

Friend NotInheritable Class NodeXmlAppliers
    Private Sub New() ' Prevent instantiation - only Shared methods
    End Sub

    ' -------------------------------------------------------
    ' STRING / PRIMITIVE PROPERTIES
    ' -------------------------------------------------------

    Friend Shared Sub Apply_NodeCaption(xNode As XmlNode, ByRef nodeCaption As String)
        If xNode.Attributes("Caption") Is Nothing Then Exit Sub
        nodeCaption = xNode.Attributes("Caption").Value
    End Sub

    Friend Shared Sub Apply_NodeKey(xNode As XmlNode, ByRef nodeKey As String)
        If xNode.Attributes("Key") Is Nothing Then Exit Sub
        nodeKey = xNode.Attributes("Key").Value
    End Sub

    Friend Shared Sub Apply_NodeTag(xNode As XmlNode, ByRef nodeTag As String)
        If xNode.Attributes("Tag") Is Nothing Then Exit Sub
        nodeTag = xNode.Attributes("Tag").Value
    End Sub

    Friend Shared Sub Apply_NodeIconClosed(xNode As XmlNode, ByRef nodeIconNameClosed As String)
        If xNode.Attributes("IconClosed") Is Nothing Then Exit Sub
        nodeIconNameClosed = xNode.Attributes("IconClosed").Value
    End Sub

    Friend Shared Sub Apply_NodeIconOpen(xNode As XmlNode, ByRef nodeIconNameOpen As String)
        If xNode.Attributes("IconOpen") Is Nothing Then Exit Sub
        nodeIconNameOpen = xNode.Attributes("IconOpen").Value
    End Sub

    Friend Shared Sub Apply_NodeIconRight(xNode As XmlNode, ByRef nodeIconRight As String)
        If xNode.Attributes("IconRight") Is Nothing Then Exit Sub
        nodeIconRight = xNode.Attributes("IconRight").Value
    End Sub

    Friend Shared Sub Apply_NodeMirrorIcons(ByRef nodeIconNameClosed As String, ByRef nodeIconNameOpen As String)
        If String.IsNullOrEmpty(nodeIconNameOpen) AndAlso nodeIconNameClosed <> "" Then
            nodeIconNameOpen = nodeIconNameClosed
        End If
        If nodeIconNameOpen <> "" AndAlso String.IsNullOrEmpty(nodeIconNameClosed) Then
            nodeIconNameClosed = nodeIconNameOpen
        End If
    End Sub

    Friend Shared Sub Apply_NodeExpanded(xNode As XmlNode, ByRef iconExpanded As Boolean)
        If xNode.Attributes("Expanded") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Expanded").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If iconExpanded <> newVal Then iconExpanded = newVal
    End Sub

    Friend Shared Sub Apply_NodeLazyNode(xNode As XmlNode, ByRef isLazy As Boolean)
        If xNode.Attributes("LazyNode") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("LazyNode").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If isLazy <> newVal Then isLazy = newVal
    End Sub

    ' -------------------------------------------------------
    ' IMAGE RESOLUTION (necesită _imageCache)
    ' -------------------------------------------------------

    Friend Shared Sub Apply_NodeClosedImage(nodeIconNameClosed As String,
                                            ByRef iconImgClosed As Image,
                                            imageCache As Dictionary(Of String, Image))
        If String.IsNullOrEmpty(nodeIconNameClosed) Then Exit Sub
        Dim value As Image = Nothing
        If imageCache.TryGetValue(nodeIconNameClosed, value) Then iconImgClosed = value
    End Sub

    Friend Shared Sub Apply_NodeOpenImage(nodeIconNameOpen As String,
                                          ByRef iconImgOpen As Image,
                                          imageCache As Dictionary(Of String, Image))
        If String.IsNullOrEmpty(nodeIconNameOpen) Then Exit Sub
        Dim value As Image = Nothing
        If imageCache.TryGetValue(nodeIconNameOpen, value) Then iconImgOpen = value
    End Sub

    Friend Shared Sub Apply_NodeRightImage(nodeIconRight As String,
                                           ByRef iconImgRight As Image,
                                           imageCache As Dictionary(Of String, Image))
        If String.IsNullOrEmpty(nodeIconRight) Then Exit Sub
        Dim value As Image = Nothing
        If imageCache.TryGetValue(nodeIconRight, value) Then iconImgRight = value
    End Sub

    ' -------------------------------------------------------
    ' TREEITEM PROPERTIES
    ' -------------------------------------------------------

    Friend Shared Sub Apply_NodeTooltip(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Tooltip") Is Nothing Then Exit Sub
        Dim xmlVal As String = xNode.Attributes("Tooltip").Value
        If newItem.Tooltip <> xmlVal Then newItem.Tooltip = xmlVal
    End Sub

    Friend Shared Sub Apply_NodeColHeaderText(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("ColHeaderText") Is Nothing Then Exit Sub
        Dim xmlVal As String = xNode.Attributes("ColHeaderText").Value
        If newItem.ColHeaderText <> xmlVal Then newItem.ColHeaderText = xmlVal
    End Sub

    Friend Shared Sub Apply_NodeBold(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Bold") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Bold").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.Bold <> newVal Then newItem.Bold = newVal
    End Sub

    Friend Shared Sub Apply_NodeItalic(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("Italic") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("Italic").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.Italic <> newVal Then newItem.Italic = newVal
    End Sub

    Friend Shared Sub Apply_NodeHasCheckbox(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("HasCheckbox") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("HasCheckbox").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.HasCheckBox <> newVal Then newItem.HasCheckBox = newVal
    End Sub

    Friend Shared Sub Apply_NodeShowRightIconOnHover(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("ShowRightIconOnHover") Is Nothing Then Exit Sub
        Dim valStr As String = xNode.Attributes("ShowRightIconOnHover").Value.Trim().ToLower()
        Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
        If newItem.ShowRightIconOnHover <> newVal Then newItem.ShowRightIconOnHover = newVal
    End Sub

    Friend Shared Sub Apply_NodeCheckState(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("CheckState") Is Nothing Then Exit Sub
        Dim xmlVal As String = xNode.Attributes("CheckState").Value.Trim()
        If newItem.CheckState <> xmlVal Then newItem.CheckState = xmlVal
    End Sub

    Friend Shared Sub Apply_NodeForeColor(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("ForeColor") Is Nothing Then Exit Sub
        Dim colorVal As String = xNode.Attributes("ForeColor").Value.Trim()
        If String.IsNullOrEmpty(colorVal) Then Exit Sub
        Try
            Dim newColor As Color = If(colorVal.StartsWith("#"c),
                                       ColorTranslator.FromHtml(colorVal),
                                       Color.FromName(colorVal))
            If newItem.NodeForeColor <> newColor Then newItem.NodeForeColor = newColor
        Catch
        End Try
    End Sub

    Friend Shared Sub Apply_NodeBackColor(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        If xNode.Attributes("BackColor") Is Nothing Then Exit Sub
        Dim colorVal As String = xNode.Attributes("BackColor").Value.Trim()
        If String.IsNullOrEmpty(colorVal) Then Exit Sub
        Try
            Dim newColor As Color = If(colorVal.StartsWith("#"c),
                                       ColorTranslator.FromHtml(colorVal),
                                       Color.FromName(colorVal))
            If newItem.NodeBackColor <> newColor Then newItem.NodeBackColor = newColor
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Citeste sub-elementele &lt;Cells&gt;&lt;Cell Col="..." Val="..." /&gt; ale unui nod XML
    ''' si le aplica pe TreeItem.Cells. Suporta BackColor si ForeColor per celula.
    ''' </summary>
    Friend Shared Sub Apply_NodeCells(xNode As XmlNode, newItem As AdvancedTreeControl.TreeItem)
        Try
            Dim cellsNode As XmlNode = xNode.SelectSingleNode("Cells")
            If cellsNode Is Nothing Then Return
            For Each cellEl As XmlNode In cellsNode.SelectNodes("Cell")
                Try
                    Dim colName As String = cellEl.Attributes("Col")?.Value
                    If String.IsNullOrEmpty(colName) Then Continue For
                    Dim cd As New AdvancedTreeControl.TreeItem.CellData()
                    cd.Value = If(cellEl.Attributes("Val")?.Value, "")
                    Dim bgStr As String = If(cellEl.Attributes("BackColor")?.Value, "")
                    If Not String.IsNullOrEmpty(bgStr) Then
                        cd.BackColor = AdvancedTreeControl.ParseColor(bgStr, Color.Empty)
                    End If
                    Dim fgStr As String = If(cellEl.Attributes("ForeColor")?.Value, "")
                    If Not String.IsNullOrEmpty(fgStr) Then
                        cd.ForeColor = AdvancedTreeControl.ParseColor(fgStr, Color.Empty)
                    End If
                    newItem.Cells(colName) = cd
                Catch ex As Exception
                    TreeLogger.Ex(ex, "Apply_NodeCells/Cell")
                End Try
            Next
        Catch ex As Exception
            TreeLogger.Ex(ex, "Apply_NodeCells")
        End Try
    End Sub

End Class