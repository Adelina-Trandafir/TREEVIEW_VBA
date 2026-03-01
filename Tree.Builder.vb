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
                AplicareConfigurare(configNode)
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
                If Not AplicareConfigurare(configNode) Then
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

    Private Function AplicareConfigurare(cfg As XmlNode, Optional Reload As Boolean = False) As Boolean
        Try
            Dim sw As New Stopwatch()
            sw.Start()

            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            If cfg.Attributes("treeID") IsNot Nothing Then
                Dim tId As String = cfg.Attributes("treeID").Value
                If Not String.IsNullOrEmpty(tId) Then
                    If Reload Then
                        If MyTree.treeID <> tId Then
                            TreeLogger.Err("EROARE: La reîncărcare, atributul 'treeId' nu corespunde cu cel inițial.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Application.Exit()
                        End If
                    Else
                        MyTree.treeID = tId
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

            ' --- BackColor ---
            If cfg.Attributes("BackColor") IsNot Nothing Then
                Try
                    Dim xmlVal = cfg.Attributes("BackColor").Value
                    Dim c As Color = ColorTranslator.FromHtml(xmlVal)
                    MyTree.BackColor = c
                    Me.BackColor = c

                    TreeLogger.Debug(Space(5) & $"BackColor xml='{xmlVal}' control='{MyTree.BackColor}'", "AplicareConfigurare")
                Catch ex As Exception
                    TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ' --- BorderColor ---
            If cfg.Attributes("BorderColor") IsNot Nothing Then
                Try
                    Dim v As String = cfg.Attributes("BorderColor").Value
                    If v.StartsWith("#"c) Then
                        Dim c As Color = ColorTranslator.FromHtml(v)
                        MyTree.BorderColor = c
                        TreeLogger.Debug(Space(5) & $"BorderColor xml='{v}' control='{MyTree.BorderColor}'", "AplicareConfigurare")
                    End If
                Catch ex As Exception
                    TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ' --- ForeColor ---
            If cfg.Attributes("ForeColor") IsNot Nothing Then
                Try
                    Dim xmlVal = cfg.Attributes("ForeColor").Value
                    Dim c As Color = ColorTranslator.FromHtml(xmlVal)
                    MyTree.ForeColor = c
                    TreeLogger.Debug(Space(5) & $"ForeColor xml='{xmlVal}' control='{MyTree.ForeColor}'", "AplicareConfigurare")
                Catch ex As Exception
                    TreeLogger.Ex(ex, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ' --- RadioButtonLevel ---
            If cfg.Attributes("RadioButtonLevel") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("RadioButtonLevel").Value
                Dim v As Integer = -1
                Dim parsed = Integer.TryParse(xmlVal, v)
                If parsed Then
                    MyTree.RadioButtonLevel = v
                    TreeLogger.Debug(Space(5) & $"RadioButtonLevel xml='{xmlVal}' control='{MyTree.RadioButtonLevel}'", "AplicareConfigurare")
                End If
            End If

            ' --- Checkboxes ---
            If MyTree.RadioButtonLevel = -1 AndAlso cfg.Attributes("CheckBoxes") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("CheckBoxes").Value
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(xmlVal, v)
                If parsed Then
                    MyTree.CheckBoxes = v = 1
                    TreeLogger.Debug(Space(5) & $"CheckBoxes xml='{xmlVal}' control='{MyTree.CheckBoxes}'", "AplicareConfigurare")
                End If
            End If

            ' --- Font ---
            Dim fName As String = "Segoe UI"
            Dim fSize As Single = 9.0F
            If MyTree.Font IsNot Nothing Then
                fName = MyTree.Font.Name
                fSize = MyTree.Font.Size
            End If

            Dim xmlFontName As String = Nothing
            Dim xmlFontSize As String = Nothing

            If cfg.Attributes("FontName") IsNot Nothing Then
                xmlFontName = cfg.Attributes("FontName").Value
                fName = xmlFontName
            End If

            If cfg.Attributes("FontSize") IsNot Nothing Then
                xmlFontSize = cfg.Attributes("FontSize").Value
                Single.TryParse(xmlFontSize, NumberStyles.Any, culture, fSize)
            End If

            MyTree.Font = New Font(fName, fSize)

            TreeLogger.Debug(Space(5) &
            $"Font xmlName='{xmlFontName}' xmlSize='{xmlFontSize}' control='{MyTree.Font.Name} {MyTree.Font.Size}pt'",
            "AplicareConfigurare")

            ' --- ItemHeight ---
            If cfg.Attributes("ItemHeight") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("ItemHeight").Value
                Dim ih As Integer = 22
                Dim v = Integer.TryParse(xmlVal, ih)
                If ih > 0 Then
                    MyTree.ItemHeight = ih
                    TreeLogger.Debug(Space(5) & $"ItemHeight xml='{xmlVal}' control='{MyTree.ItemHeight}'", "AplicareConfigurare")
                End If
            End If

            ' --- HasNodeIcons ---
            If cfg.Attributes("HasNodeIcons") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("HasNodeIcons").Value
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(xmlVal, v)
                If parsed Then
                    MyTree.HasNodeIcons = v = 1
                    TreeLogger.Debug(Space(5) & $"HasNodeIcons xml='{xmlVal}' control='{MyTree.HasNodeIcons}'", "AplicareConfigurare")
                End If
            End If

            ' --- PopupTree ---
            If cfg.Attributes("PopupTree") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("PopupTree").Value
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(xmlVal, v)
                If parsed Then
                    MyTree.IsPopupTree = v = 1
                    TreeLogger.Debug(Space(5) & $"PopupTree xml='{xmlVal}' control='{MyTree.IsPopupTree}'", "AplicareConfigurare")
                End If
            End If

            ' --- LeftIconHeight ---
            If cfg.Attributes("LeftIconHeight") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("LeftIconHeight").Value
                Dim lih As Integer = 16
                Dim v = Integer.TryParse(xmlVal, lih)
                If lih > 0 Then
                    MyTree.LeftIconSize = New Size(lih, lih)
                    TreeLogger.Debug(Space(5) & $"LeftIconHeight xml='{xmlVal}' control='{MyTree.LeftIconSize}'", "AplicareConfigurare")
                End If
            End If

            ' --- RightIconHeight ---
            If cfg.Attributes("RightIconHeight") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("RightIconHeight").Value
                Dim rih As Integer = 16
                Dim v = Integer.TryParse(xmlVal, rih)
                If rih > 0 Then
                    MyTree.RightIconSize = New Size(rih, rih)
                    TreeLogger.Debug(Space(5) & $"RightIconHeight xml='{xmlVal}' control='{MyTree.RightIconSize}'", "AplicareConfigurare")
                End If
            End If

            ' --- CheckboxSize ---
            If cfg.Attributes("CheckboxSize") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("CheckboxSize").Value
                Dim cs As Integer = 16
                Dim v = Integer.TryParse(xmlVal, cs)
                If cs > 0 Then
                    MyTree.CheckBoxSize = cs
                    TreeLogger.Debug(Space(5) & $"CheckboxSize xml='{xmlVal}' control='{MyTree.CheckBoxSize}'", "AplicareConfigurare")
                End If
            End If

            ' --- RightClickFunc ---
            If cfg.Attributes("RightClickFunc") IsNot Nothing Then
                Dim xmlVal As String = cfg.Attributes("RightClickFunc").Value
                MyTree.RightClickFunction = xmlVal
                TreeLogger.Debug(Space(5) & $"RightClickFunc xml='{xmlVal}' control='{MyTree.RightClickFunction}'", "AplicareConfigurare")
            End If

            ' --- Indent ---
            If cfg.Attributes("Indent") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("Indent").Value
                Dim indentVal As Integer = 20
                Dim v = Integer.TryParse(xmlVal, indentVal)
                If indentVal >= 0 Then
                    MyTree.Indent = indentVal
                    TreeLogger.Debug(Space(5) & $"Indent xml='{xmlVal}' control='{MyTree.Indent}'", "AplicareConfigurare")
                End If
            End If

            ' --- ExpanderSize ---
            If cfg.Attributes("ExpanderSize") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("ExpanderSize").Value
                Dim expSize As Integer = 12
                Dim v = Integer.TryParse(xmlVal, expSize)
                If expSize >= 0 Then
                    MyTree.ExpanderSize = expSize
                    TreeLogger.Debug(Space(5) & $"ExpanderSize xml='{xmlVal}' control='{MyTree.ExpanderSize}'", "AplicareConfigurare")
                End If
            End If

            ' --- RootButton ---
            If cfg.Attributes("RootButton") IsNot Nothing Then
                Dim xmlVal = cfg.Attributes("RootButton").Value
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(xmlVal, v)
                If parsed Then
                    MyTree.RootButton = v = 1
                    TreeLogger.Debug(Space(5) & $"RootButton xml='{xmlVal}' control='{MyTree.RootButton}'", "AplicareConfigurare")
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
            ' 1. Citim atributele
            Dim nodeCaption As String = ""
            If xNode.Attributes("Caption") IsNot Nothing Then nodeCaption = xNode.Attributes("Caption").Value

            Dim nodeKey As String = ""
            If xNode.Attributes("Key") IsNot Nothing Then nodeKey = xNode.Attributes("Key").Value

            Dim nodeTag As String = ""
            If xNode.Attributes("Tag") IsNot Nothing Then nodeTag = xNode.Attributes("Tag").Value

            ' 2. Gestionare Iconițe
            Dim nodeIconNameClosed As String = ""
            Dim nodeIconNameOpen As String = ""
            Dim nodeIconRight As String = ""

            If xNode.Attributes("IconClosed") IsNot Nothing Then nodeIconNameClosed = xNode.Attributes("IconClosed").Value
            If xNode.Attributes("IconOpen") IsNot Nothing Then nodeIconNameOpen = xNode.Attributes("IconOpen").Value
            If xNode.Attributes("IconRight") IsNot Nothing Then nodeIconRight = xNode.Attributes("IconRight").Value

            ' 2.1 Dacă nu s-au specificat ambele, le setăm la fel
            If String.IsNullOrEmpty(nodeIconNameOpen) AndAlso nodeIconNameClosed <> "" Then
                nodeIconNameOpen = nodeIconNameClosed
            End If

            If nodeIconNameOpen <> "" AndAlso String.IsNullOrEmpty(nodeIconNameClosed) Then
                nodeIconNameClosed = nodeIconNameOpen
            End If

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

            ' 3. Setări Stare (Expanded)
            Dim iconExpanded As Boolean = False
            If xNode.Attributes("Expanded") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Expanded").Value.Trim().ToLower()
                ' Verificăm manual cazurile comune: "1", "true", "-1"
                If valStr = "1" OrElse valStr = "-1" OrElse valStr = "true" Then
                    iconExpanded = True
                Else
                    iconExpanded = False
                End If
            End If

            ' 4. Setări Stare (LazyNode)
            Dim isLazy As Boolean = False
            If xNode.Attributes("LazyNode") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("LazyNode").Value.Trim().ToLower()
                isLazy = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
            End If

            ' 5. Adăugăm Itemul
            Dim newItem As AdvancedTreeControl.TreeItem = MyTree.AddItem(nodeKey, nodeCaption, parentItem, iconImgClosed, iconImgOpen, iconImgRight, nodeTag, iconExpanded, isLazy)
            newItem.Key = nodeKey

            If xNode.Attributes("Tooltip") IsNot Nothing Then newItem.Tooltip = xNode.Attributes("Tooltip").Value

            ' 6. Atribute vizuale per nod (Bold, Italic, ForeColor, BackColor)
            If xNode.Attributes("Bold") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Bold").Value.Trim().ToLower()
                newItem.Bold = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
            End If

            If xNode.Attributes("Italic") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("Italic").Value.Trim().ToLower()
                newItem.Italic = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
            End If

            If xNode.Attributes("HasCheckbox") IsNot Nothing Then
                Dim valStr As String = xNode.Attributes("HasCheckbox").Value.Trim().ToLower()
                newItem.HasCheckBox = (valStr = "1" OrElse valStr = "-1" OrElse valStr = "true")
            End If

            If xNode.Attributes("ForeColor") IsNot Nothing Then
                Dim colorVal As String = xNode.Attributes("ForeColor").Value.Trim()
                If Not String.IsNullOrEmpty(colorVal) Then
                    Try
                        If colorVal.StartsWith("#"c) Then
                            newItem.NodeForeColor = ColorTranslator.FromHtml(colorVal)
                        Else
                            newItem.NodeForeColor = Color.FromName(colorVal)
                        End If
                    Catch
                        ' Ignorăm culori invalide, rămâne Color.Empty
                    End Try
                End If
            End If

            If xNode.Attributes("BackColor") IsNot Nothing Then
                Dim colorVal As String = xNode.Attributes("BackColor").Value.Trim()
                If Not String.IsNullOrEmpty(colorVal) Then
                    Try
                        If colorVal.StartsWith("#"c) Then
                            newItem.NodeBackColor = ColorTranslator.FromHtml(colorVal)
                        Else
                            newItem.NodeBackColor = Color.FromName(colorVal)
                        End If
                    Catch
                    End Try
                End If
            End If
            ' 5. Recursivitate
            For Each childNode As XmlNode In xNode.SelectNodes("Node")
                AddXmlNodeToTree(childNode, newItem)
            Next

            'TreeLogger.Debug($"AddXmlNodeToTree - Adăugat nod '{nodeCaption}' cu key='{nodeKey}' sub parent='{If(parentItem IsNot Nothing, parentItem.Caption, "ROOT")}'", "AddXmlNodeToTree")
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
