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
