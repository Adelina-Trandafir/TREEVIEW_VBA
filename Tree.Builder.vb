Imports System.Globalization
Imports System.IO
Imports System.Text.Json
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

            ' 3. POPULARE NODURI
            Dim nodesRoot = xDoc.SelectNodes("/Tree/Nodes/Node")

            For Each xNode As XmlNode In nodesRoot
                AddXmlNodeToTree(xNode, Nothing)
            Next

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
            If cNodeKey <> "" Then
                Dim foundNode As AdvancedTreeControl.TreeItem = Nothing

                ' Iterăm prin rădăcini pentru a găsi nodul (fix pentru eroarea cu 'root')
                For Each rootItem In MyTree.Items
                    foundNode = FindNodeByIdRecursive(rootItem, cNodeKey)
                    If foundNode IsNot Nothing Then Exit For
                Next

                If foundNode IsNot Nothing Then
                    MyTree.SelectedNode = foundNode
                    foundNode.SetExpanded(True, True)
                    ScrollToNode(foundNode)
                End If
            End If

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

            '========================
            ' treeID (obligatoriu)
            '========================
            If cfg.Attributes("treeID") IsNot Nothing Then
                Dim tId As String = cfg.Attributes("treeID").Value
                If Not String.IsNullOrEmpty(tId) Then
                    If Reload Then
                        If MyTree.treeID <> tId Then
                            TreeLogger.Err("EROARE: La reîncărcare, atributul 'treeId' nu corespunde cu cel inițial.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Application.Exit()
                        End If
                    Else
                        If MyTree.treeID <> tId Then MyTree.treeID = tId
                    End If

                    TreeLogger.Debug(Space(5) & $"treeID xml='{tId}' control='{MyTree.treeID}'", "AplicareConfigurare")
                Else
                    TreeLogger.Err("EROARE: Atributul 'treeId' nu poate fi gol în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Application.Exit()
                End If
            Else
                TreeLogger.Err("EROARE: Atributul 'treeId' este obligatoriu în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If

            '========================
            ' BackColor
            '========================
            If cfg.Attributes("BackColor") IsNot Nothing Then
                Try
                    Dim xmlVal As String = cfg.Attributes("BackColor").Value
                    Dim c As Color = ColorTranslator.FromHtml(xmlVal)

                    If MyTree.BackColor <> c Then
                        MyTree.BackColor = c
                    End If
                    If Me.BackColor <> c Then
                        Me.BackColor = c
                    End If

                    TreeLogger.Debug(Space(5) & $"BackColor xml='{xmlVal}' control='{MyTree.BackColor}'", "AplicareConfigurare")
                Catch ex As Exception
                    TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            '========================
            ' BorderColor
            '========================
            If cfg.Attributes("BorderColor") IsNot Nothing Then
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
            End If

            '========================
            ' ForeColor
            '========================
            If cfg.Attributes("ForeColor") IsNot Nothing Then
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
            End If

            '========================
            ' RadioButtonLevel
            '========================
            If cfg.Attributes("RadioButtonLevel") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RadioButtonLevel").Value
                Dim v As Integer = MyTree.RadioButtonLevel
                If Integer.TryParse(xmlVal, v) Then
                    If MyTree.RadioButtonLevel <> v Then
                        MyTree.RadioButtonLevel = v
                    End If
                    TreeLogger.Debug(Space(5) & $"RadioButtonLevel xml='{xmlVal}' control='{MyTree.RadioButtonLevel}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' CheckBoxes (doar daca RadioButtonLevel = -1)
            '========================
            If MyTree.RadioButtonLevel = -1 AndAlso cfg.Attributes("CheckBoxes") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("CheckBoxes").Value
                Dim v As Integer = If(MyTree.CheckBoxes, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.CheckBoxes <> newVal Then
                        MyTree.CheckBoxes = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"CheckBoxes xml='{xmlVal}' control='{MyTree.CheckBoxes}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' Font
            ' - daca XML nu are nimic -> NU schimba
            ' - daca XML are, dar rezultatul e identic cu cel curent -> NU schimba
            '========================
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
                Dim needSet As Boolean = (MyTree.Font Is Nothing) OrElse (MyTree.Font.Name <> fName) OrElse (Math.Abs(MyTree.Font.Size - fSize) > 0.001F)
                If needSet Then
                    MyTree.Font = New Font(fName, fSize)
                End If

                TreeLogger.Debug(Space(5) &
                $"Font xmlName='{xmlFontName}' xmlSize='{xmlFontSize}' control='{MyTree.Font.Name} {MyTree.Font.Size}pt'",
                "AplicareConfigurare")
            End If

            '========================
            ' FontName
            '========================
            If cfg.Attributes("FontName") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("FontName").Value
                If MyTree.FontName <> xmlVal Then
                    MyTree.FontName = xmlVal
                End If
                TreeLogger.Debug(Space(5) & $"FontName xml='{xmlVal}' control='{MyTree.FontName}'", "AplicareConfigurare")
            End If

            '========================
            ' FontSize
            '========================
            If cfg.Attributes("FontSize") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("FontSize").Value
                Dim ih As Integer = MyTree.FontSize
                If Integer.TryParse(xmlVal, ih) AndAlso ih > 0 Then
                    If MyTree.FontSize <> ih Then
                        MyTree.FontSize = ih
                    End If
                    TreeLogger.Debug(Space(5) & $"FontSize xml='{xmlVal}' control='{MyTree.FontSize}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' ItemHeight
            '========================
            If cfg.Attributes("ItemHeight") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("ItemHeight").Value
                Dim ih As Integer = MyTree.ItemHeight
                If Integer.TryParse(xmlVal, ih) AndAlso ih > 0 Then
                    If MyTree.ItemHeight <> ih Then
                        MyTree.ItemHeight = ih
                    End If
                    TreeLogger.Debug(Space(5) & $"ItemHeight xml='{xmlVal}' control='{MyTree.ItemHeight}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' HasNodeIcons
            '========================
            If cfg.Attributes("HasNodeIcons") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("HasNodeIcons").Value
                Dim v As Integer = If(MyTree.HasNodeIcons, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.HasNodeIcons <> newVal Then
                        MyTree.HasNodeIcons = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"HasNodeIcons xml='{xmlVal}' control='{MyTree.HasNodeIcons}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' PopupTree
            '========================
            If cfg.Attributes("PopupTree") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("PopupTree").Value
                Dim v As Integer = If(MyTree.IsPopupTree, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.IsPopupTree <> newVal Then
                        MyTree.IsPopupTree = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"PopupTree xml='{xmlVal}' control='{MyTree.IsPopupTree}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' LeftIconHeight -> LeftIconSize (patrat)
            '========================
            If cfg.Attributes("LeftIconHeight") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("LeftIconHeight").Value
                Dim lih As Integer = If(MyTree.LeftIconSize.Height > 0, MyTree.LeftIconSize.Height, MyTree.LeftIconSize.Width)
                If lih <= 0 Then lih = 16

                If Integer.TryParse(xmlVal, lih) AndAlso lih > 0 Then
                    Dim newSize As New Size(lih, lih)
                    If MyTree.LeftIconSize <> newSize Then
                        MyTree.LeftIconSize = newSize
                    End If
                    TreeLogger.Debug(Space(5) & $"LeftIconHeight xml='{xmlVal}' control='{MyTree.LeftIconSize}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' RightIconHeight -> RightIconSize (patrat)
            '========================
            If cfg.Attributes("RightIconHeight") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RightIconHeight").Value
                Dim rih As Integer = If(MyTree.RightIconSize.Height > 0, MyTree.RightIconSize.Height, MyTree.RightIconSize.Width)
                If rih <= 0 Then rih = 16

                If Integer.TryParse(xmlVal, rih) AndAlso rih > 0 Then
                    Dim newSize As New Size(rih, rih)
                    If MyTree.RightIconSize <> newSize Then
                        MyTree.RightIconSize = newSize
                    End If
                    TreeLogger.Debug(Space(5) & $"RightIconHeight xml='{xmlVal}' control='{MyTree.RightIconSize}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' CheckboxSize
            '========================
            If cfg.Attributes("CheckboxSize") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("CheckboxSize").Value
                Dim cs As Integer = MyTree.CheckBoxSize
                If cs <= 0 Then cs = 16

                If Integer.TryParse(xmlVal, cs) AndAlso cs > 0 Then
                    If MyTree.CheckBoxSize <> cs Then
                        MyTree.CheckBoxSize = cs
                    End If
                    TreeLogger.Debug(Space(5) & $"CheckboxSize xml='{xmlVal}' control='{MyTree.CheckBoxSize}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' RightClickFunc
            '========================
            If cfg.Attributes("RightClickFunc") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RightClickFunc").Value
                If MyTree.RightClickFunction <> xmlVal Then
                    MyTree.RightClickFunction = xmlVal
                End If
                TreeLogger.Debug(Space(5) & $"RightClickFunc xml='{xmlVal}' control='{MyTree.RightClickFunction}'", "AplicareConfigurare")
            End If

            '========================
            ' Indent
            '========================
            If cfg.Attributes("Indent") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("Indent").Value
                Dim indentVal As Integer = MyTree.Indent
                If Integer.TryParse(xmlVal, indentVal) AndAlso indentVal >= 0 Then
                    If MyTree.Indent <> indentVal Then
                        MyTree.Indent = indentVal
                    End If
                    TreeLogger.Debug(Space(5) & $"Indent xml='{xmlVal}' control='{MyTree.Indent}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' ExpanderSize
            '========================
            If cfg.Attributes("ExpanderSize") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("ExpanderSize").Value
                Dim expSize As Integer = MyTree.ExpanderSize
                If expSize < 0 Then expSize = 0

                If Integer.TryParse(xmlVal, expSize) AndAlso expSize >= 0 Then
                    If MyTree.ExpanderSize <> expSize Then
                        MyTree.ExpanderSize = expSize
                    End If
                    TreeLogger.Debug(Space(5) & $"ExpanderSize xml='{xmlVal}' control='{MyTree.ExpanderSize}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' RootButton
            '========================
            If cfg.Attributes("RootButton") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RootButton").Value
                Dim v As Integer = If(MyTree.RootButton, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.RootButton <> newVal Then
                        MyTree.RootButton = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"RootButton xml='{xmlVal}' control='{MyTree.RootButton}'", "AplicareConfigurare")
                End If
            End If

            TreeLogger.Info($"Configurare aplicată cu succes în {sw.ElapsedMilliseconds}ms", "AplicareConfigurare")

            '========================
            ' ReRaiseClickOnSameNode
            '========================
            If cfg.Attributes("ReRaiseClickOnSameNode") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("ReRaiseClickOnSameNode").Value
                Dim v As Integer = If(MyTree.ReRaiseClickOnSameNode, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.ReRaiseClickOnSameNode <> newVal Then
                        MyTree.ReRaiseClickOnSameNode = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"ReRaiseClickOnSameNode xml='{xmlVal}' control='{MyTree.ReRaiseClickOnSameNode}'", "AplicareConfigurare")
                End If
            End If

            '========================
            ' RaiseLeftClickOnRightClick
            '========================
            If cfg.Attributes("RaiseLeftClickOnRightClick") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RaiseLeftClickOnRightClick").Value
                Dim v As Integer = If(MyTree.RaiseLeftClickOnRightClick, 1, 0)
                If Integer.TryParse(xmlVal, v) Then
                    Dim newVal As Boolean = (v = 1)
                    If MyTree.RaiseLeftClickOnRightClick <> newVal Then
                        MyTree.RaiseLeftClickOnRightClick = newVal
                    End If
                    TreeLogger.Debug(Space(5) & $"RaiseLeftClickOnRightClick xml='{xmlVal}' control='{MyTree.RaiseLeftClickOnRightClick}'", "AplicareConfigurare")
                End If
            End If

            TreeLogger.Info($"Configurare aplicată cu succes în {sw.ElapsedMilliseconds}ms", "AplicareConfigurare")

            Return True

        Catch ex As Exception
            TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ' =============================================================
    ' LOGICA RECURSIVĂ DE ADĂUGARE NODURI
    ' =============================================================
    Private Sub AddXmlNodeToTree(xNode As XmlNode, parentItem As AdvancedTreeControl.TreeItem)
        Try
            ' --- Caption ---
            Dim nodeCaption As String = ""
            If xNode.Attributes("Caption") IsNot Nothing Then nodeCaption = xNode.Attributes("Caption").Value

            ' --- Key ---
            Dim nodeKey As String = ""
            If xNode.Attributes("Key") IsNot Nothing Then nodeKey = xNode.Attributes("Key").Value

            ' --- Tag ---
            Dim nodeTag As String = ""
            If xNode.Attributes("Tag") IsNot Nothing Then nodeTag = xNode.Attributes("Tag").Value

            ' --- IconClosed / IconOpen / IconRight ---
            Dim nodeIconNameClosed As String = ""
            Dim nodeIconNameOpen As String = ""
            Dim nodeIconRight As String = ""

            If xNode.Attributes("IconClosed") IsNot Nothing Then nodeIconNameClosed = xNode.Attributes("IconClosed").Value
            If xNode.Attributes("IconOpen") IsNot Nothing Then nodeIconNameOpen = xNode.Attributes("IconOpen").Value
            If xNode.Attributes("IconRight") IsNot Nothing Then nodeIconRight = xNode.Attributes("IconRight").Value

            ' --- If only one of closed/open is specified, mirror it ---
            If String.IsNullOrEmpty(nodeIconNameOpen) AndAlso nodeIconNameClosed <> "" Then nodeIconNameOpen = nodeIconNameClosed
            If nodeIconNameOpen <> "" AndAlso String.IsNullOrEmpty(nodeIconNameClosed) Then nodeIconNameClosed = nodeIconNameOpen

            Dim iconImgClosed As Image = Nothing
            Dim iconImgOpen As Image = Nothing
            Dim iconImgRight As Image = Nothing

            If Not String.IsNullOrEmpty(nodeIconNameClosed) Then
                Dim value As Image = Nothing
                If _imageCache.TryGetValue(nodeIconNameClosed, value) Then iconImgClosed = value
            End If

            If Not String.IsNullOrEmpty(nodeIconNameOpen) Then
                Dim value As Image = Nothing
                If _imageCache.TryGetValue(nodeIconNameOpen, value) Then iconImgOpen = value
            End If

            If Not String.IsNullOrEmpty(nodeIconRight) Then
                Dim value As Image = Nothing
                If _imageCache.TryGetValue(nodeIconRight, value) Then iconImgRight = value
            End If

            ' --- Expanded ---
            Dim iconExpanded As Boolean = False ' default TreeItem
            If xNode.Attributes("Expanded") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Expanded").Value.Trim().ToLower()
                Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
                If iconExpanded <> newVal Then iconExpanded = newVal
            End If

            ' --- LazyNode ---
            Dim isLazy As Boolean = False ' default TreeItem
            If xNode.Attributes("LazyNode") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("LazyNode").Value.Trim().ToLower()
                Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
                If isLazy <> newVal Then isLazy = newVal
            End If

            ' --- AddItem ---
            Dim newItem As AdvancedTreeControl.TreeItem =
            MyTree.AddItem(nodeKey, nodeCaption, parentItem, iconImgClosed, iconImgOpen, iconImgRight, nodeTag, iconExpanded, isLazy)

            If newItem Is Nothing Then Exit Sub

            ' --- Key (redundant, but keep) ---
            If newItem.Key <> nodeKey Then newItem.Key = nodeKey

            ' --- Tooltip ---
            If xNode.Attributes("Tooltip") IsNot Nothing Then
                Dim xmlVal As String = xNode.Attributes("Tooltip").Value
                If newItem.Tooltip <> xmlVal Then newItem.Tooltip = xmlVal
            End If

            ' --- Bold ---
            If xNode.Attributes("Bold") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Bold").Value.Trim().ToLower()
                Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
                If newItem.Bold <> newVal Then newItem.Bold = newVal
            End If

            ' --- Italic ---
            If xNode.Attributes("Italic") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Italic").Value.Trim().ToLower()
                Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
                If newItem.Italic <> newVal Then newItem.Italic = newVal
            End If

            ' --- HasCheckbox ---
            If xNode.Attributes("HasCheckbox") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("HasCheckbox").Value.Trim().ToLower()
                Dim newVal As Boolean = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
                If newItem.HasCheckBox <> newVal Then newItem.HasCheckBox = newVal
            End If

            ' --- CheckState ---
            If xNode.Attributes("CheckState") IsNot Nothing Then
                Dim xmlVal As String = xNode.Attributes("CheckState").Value.Trim() ' "Checked"/"Unchecked"/"Indeterminate"
                If newItem.CheckState <> xmlVal Then newItem.CheckState = xmlVal
            End If

            ' --- ForeColor ---
            If xNode.Attributes("ForeColor") IsNot Nothing Then
                Dim colorVal As String = xNode.Attributes("ForeColor").Value.Trim()
                If Not String.IsNullOrEmpty(colorVal) Then
                    Try
                        Dim newColor As Color
                        If colorVal.StartsWith("#"c) Then
                            newColor = ColorTranslator.FromHtml(colorVal)
                        Else
                            newColor = Color.FromName(colorVal)
                        End If

                        If newItem.NodeForeColor <> newColor Then newItem.NodeForeColor = newColor
                    Catch
                        ' Ignoram culori invalide
                    End Try
                End If
            End If

            ' --- BackColor ---
            If xNode.Attributes("BackColor") IsNot Nothing Then
                Dim colorVal As String = xNode.Attributes("BackColor").Value.Trim()
                If Not String.IsNullOrEmpty(colorVal) Then
                    Try
                        Dim newColor As Color
                        If colorVal.StartsWith("#"c) Then
                            newColor = ColorTranslator.FromHtml(colorVal)
                        Else
                            newColor = Color.FromName(colorVal)
                        End If

                        If newItem.NodeBackColor <> newColor Then newItem.NodeBackColor = newColor
                    Catch
                        ' Ignoram culori invalide
                    End Try
                End If
            End If

            ' --- Children (recursive) ---
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

End Class
