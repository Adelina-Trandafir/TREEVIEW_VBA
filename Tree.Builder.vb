Imports System.Globalization
Imports System.IO
Imports System.Xml

Partial Public Class Tree
    ' =============================================================
    ' ÎNCĂRCARE XML + CONFIGURARE
    ' =============================================================
    Private Sub LoadXmlDataFromString(xmlContent As String)
        Try
            Dim xDoc As New XmlDocument()
            xDoc.LoadXml(xmlContent)

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

        Catch ex As Exception
            MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "LoadXmlDataFromString")
        Finally
            MyTree.ResumeLayout()
        End Try
    End Sub

    Private Sub AplicareConfigurare(cfg As XmlNode)
        Try
            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            ' --- BackColor ---
            If cfg.Attributes("BackColor") IsNot Nothing Then
                Try
                    Dim c As Color = ColorTranslator.FromHtml(cfg.Attributes("BackColor").Value)
                    MyTree.BackColor = c
                    Me.BackColor = c
                Catch ex As Exception
                    MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "AplicareConfigurare")
                End Try
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

        Catch ex As Exception
            MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "AplicareConfigurare")
        End Try
    End Sub

    ' =============================================================
    ' LOGICA RECURSIVĂ DE ADĂUGARE NODURI
    ' =============================================================
    Private Sub AddXmlNodeToTree(xNode As XmlNode, parentItem As AdvancedTreeControl.TreeItem)
        Try
            ' 1. Citim atributele
            Dim text As String = ""
            If xNode.Attributes("Text") IsNot Nothing Then text = xNode.Attributes("Text").Value

            Dim nodeId As String = ""
            If xNode.Attributes("ID") IsNot Nothing Then nodeId = xNode.Attributes("ID").Value

            ' 2. Gestionare Iconiță
            Dim iconName As String = ""
            If xNode.Attributes("IconClosed") IsNot Nothing Then iconName = xNode.Attributes("IconClosed").Value
            If String.IsNullOrEmpty(iconName) AndAlso xNode.Attributes("IconOpen") IsNot Nothing Then
                iconName = xNode.Attributes("IconOpen").Value
            End If

            Dim iconImg As Image = Nothing
            If Not String.IsNullOrEmpty(iconName) Then
                Dim value As Image = Nothing
                If _imageCache.TryGetValue(iconName, value) Then iconImg = value
            End If

            ' 3. Adăugăm Itemul
            Dim newItem As AdvancedTreeControl.TreeItem = MyTree.AddItem(text, parentItem, iconImg)
            newItem.Tag = nodeId

            ' 4. Setări Stare (Expanded)
            Dim iExpanded As Boolean = False
            If xNode.Attributes("Expanded") IsNot Nothing Then
                Dim v = Boolean.TryParse(xNode.Attributes("Expanded").Value, iExpanded)
                If Not v Then iExpanded = False
            End If
            newItem.Expanded = iExpanded

            ' 5. Recursivitate
            For Each childNode As XmlNode In xNode.SelectNodes("Node")
                AddXmlNodeToTree(childNode, newItem)
            Next
        Catch ex As Exception
            MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "AddXmlNodeToTree")
        End Try
    End Sub

    Private Sub LoadImagesToCache(imgRoot As XmlNode)
        Dim count As Integer = 0
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
                MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "LoadImagesToCache")
            End Try
        Next
    End Sub

End Class
