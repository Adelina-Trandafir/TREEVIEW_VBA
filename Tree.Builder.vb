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
            MessageBox.Show("EROARE: " & ex.Message, "LoadXmlDataFromString", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            MyTree.ResumeLayout()
        End Try
    End Function

    Private Function ReLoadXmlData(filePath As String) As Boolean
        If Not File.Exists(filePath) Then Return False

        Try
            Dim xDoc As New XmlDocument()
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

            MyTree.Invalidate()
            Return True

        Catch ex As Exception
            MessageBox.Show("EROARE: " & ex.Message, "LoadXmlDataFromString", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False

        Finally
            MyTree.ResumeLayout()
        End Try
    End Function

    Private Function AplicareConfigurare(cfg As XmlNode, Optional Reload As Boolean = False) As Boolean
        Try
            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            If cfg.Attributes("treeID") IsNot Nothing Then
                Dim tId As String = cfg.Attributes("treeID").Value
                If Not String.IsNullOrEmpty(tId) Then
                    If Reload Then
                        ' La reload verific daca e acelasi ID. Daca nu, eroare.
                        If MyTree.treeID <> tId Then
                            MessageBox.Show("EROARE: La reîncărcare, atributul 'treeId' nu corespunde cu cel inițial.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Application.Exit()
                        End If
                    Else
                        MyTree.treeID = tId
                    End If
                Else
                    MessageBox.Show("EROARE: Atributul 'treeId' nu poate fi gol în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Application.Exit()
                End If
            Else
                MessageBox.Show("EROARE: Atributul 'treeId' este obligatoriu în configurație.", "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If

            ' --- BackColor ---
            If cfg.Attributes("BackColor") IsNot Nothing Then
                Try
                    Dim c As Color = ColorTranslator.FromHtml(cfg.Attributes("BackColor").Value)
                    MyTree.BackColor = c
                    Me.BackColor = c
                Catch ex As Exception
                    MessageBox.Show("EROARE: " & ex.Message, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ' --- ForeColor ---
            If cfg.Attributes("ForeColor") IsNot Nothing Then
                Try
                    Dim c As Color = ColorTranslator.FromHtml(cfg.Attributes("ForeColor").Value)
                    MyTree.ForeColor = c
                Catch ex As Exception
                    MessageBox.Show("EROARE: " & ex.Message, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ' --- Checkboxex ---
            If cfg.Attributes("CheckBoxes") IsNot Nothing Then
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(cfg.Attributes("CheckBoxes").Value, v)
                If parsed Then MyTree.CheckBoxes = v = 1
            End If

            ' --- Font ---
            Dim fName As String = "Segoe UI"
            Dim fSize As Single = 9.0F

            If MyTree.Font IsNot Nothing Then
                fName = MyTree.Font.Name
                fSize = MyTree.Font.Size
            End If

            If cfg.Attributes("FontName") IsNot Nothing Then fName = cfg.Attributes("FontName").Value
            If cfg.Attributes("FontSize") IsNot Nothing Then
                Single.TryParse(cfg.Attributes("FontSize").Value, NumberStyles.Any, culture, fSize)
            End If

            MyTree.Font = New Font(fName, fSize)

            ' --- ItemHeight ---
            If cfg.Attributes("ItemHeight") IsNot Nothing Then
                Dim ih As Integer = 22
                Dim v = Integer.TryParse(cfg.Attributes("ItemHeight").Value, ih)
                If ih > 0 Then MyTree.ItemHeight = ih
            End If

            ' --- NodeIcons ---
            If cfg.Attributes("HasNodeIcons") IsNot Nothing Then
                Dim v As Integer = 0
                Dim parsed = Integer.TryParse(cfg.Attributes("HasNodeIcons").Value, v)
                If parsed Then MyTree.HasNodeIcons = v = 1
            End If

            ' === LeftIconHeight ===
            If cfg.Attributes("LeftIconHeight") IsNot Nothing Then
                Dim lih As Integer = 16
                Dim v = Integer.TryParse(cfg.Attributes("LeftIconHeight").Value, lih)
                If lih > 0 Then MyTree.LeftIconSize = New Size(lih, lih)
            End If

            ' === RightIconHeight ===
            If cfg.Attributes("RightIconHeight") IsNot Nothing Then
                Dim rih As Integer = 16
                Dim v = Integer.TryParse(cfg.Attributes("RightIconHeight").Value, rih)
                If rih > 0 Then MyTree.RightIconSize = New Size(rih, rih)
            End If

            ' --- CheckboxSize ---
            If cfg.Attributes("CheckboxSize") IsNot Nothing Then
                Dim cs As Integer = 16
                Dim v = Integer.TryParse(cfg.Attributes("CheckboxSize").Value, cs)
                If cs > 0 Then MyTree.CheckBoxSize = cs
            End If

            ' --- RightClickFunc ---
            If cfg.Attributes("RightClickFunc") IsNot Nothing Then
                Dim rcFunc As String = cfg.Attributes("RightClickFunc").Value
                MyTree.RightClickFunction = rcFunc
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("EROARE: " & ex.Message, "AplicareConfigurare", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

            ' 5. Recursivitate
            For Each childNode As XmlNode In xNode.SelectNodes("Node")
                AddXmlNodeToTree(childNode, newItem)
            Next
        Catch ex As Exception
            MessageBox.Show("EROARE: " & ex.Message, "AddXmlNodeToTree", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                MessageBox.Show("EROARE: " & ex.Message, "LoadImagesToCache", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Next
    End Sub

End Class
