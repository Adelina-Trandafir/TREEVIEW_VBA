Imports System.ComponentModel
Imports System.Globalization ' IMPORT NECESAR PENTRU PUNCT ZECIMAL
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Xml

Partial Public Class Tree
    Public Sub New()
        InitializeComponent()

        Me.FormBorderStyle = FormBorderStyle.None
        Me.ShowInTaskbar = False
        Me.TopLevel = False
        Me.DoubleBuffered = True

        ' --- INIT CONTROL CUSTOM ---
        MyTree = New AdvancedTreeControl With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.White,
            .ItemHeight = 20,
            .Indent = 19,
            .RightIconHeight = 16
        }

        Me.Controls.Add(MyTree)

        ' --- EVENIMENTE CLICK ---
        AddHandler MyTree.NodeMouseUp, AddressOf MyTree_NodeMouseUp
        AddHandler MyTree.NodeMouseDown, AddressOf MyTree_NodeMouseDown

        _MonitorTimer = New Timer()
        _MonitorTimer.Interval = 100
        AddHandler _MonitorTimer.Tick, AddressOf AliniazaLaParinte
    End Sub

    Private Sub Tree_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CurataResurseSiIesi()
    End Sub

    Private Sub Tree_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim args As String() = Environment.GetCommandLineArgs()

        For Each arg As String In args
            Dim lowerArg As String = arg.ToLower()

            If lowerArg.StartsWith("/frm:") Then
                _formHwnd = New IntPtr(Long.Parse(arg.Substring(5)))

            ElseIf lowerArg.StartsWith("/acc:") Then
                _mainAccessHwnd = New IntPtr(Long.Parse(arg.Substring(5)))

            ElseIf lowerArg.StartsWith("/idt:") Then
                _idTree = arg.Substring(5)

            ElseIf lowerArg.StartsWith("/fis:") Then
                _fisier = arg.Substring(5)
            End If
        Next

#If DEBUG Then
        ' Valori Hardcoded pentru debug
        _formHwnd = New IntPtr(5637362)

        _mainAccessHwnd = New IntPtr(4064482)
        _idTree = "Clasificatii"
        _fisier = "C:\Avacont\tree_Clasificatii.xml"
#End If

        ' Conectare COM
        If _mainAccessHwnd <> IntPtr.Zero Then
            ConecteazaLaAccess(_mainAccessHwnd)
        End If

        If _formHwnd <> IntPtr.Zero Then
            Me.TopLevel = False
            Me.FormBorderStyle = FormBorderStyle.None
            Me.ShowInTaskbar = False

            SetParent(Me.Handle, _formHwnd)

            _MonitorTimer.Start()
            AliniazaLaParinte()
        Else
#If DEBUG Then
#Else
            MessageBox.Show("Acest program trebuie deschis din Access!")
            Application.Exit()
#End If
        End If

        'If Not String.IsNullOrEmpty(_fisier) Then
        'LoadXmlData(_fisier)
        Dim xmlContent As String = GetValoareLocala("TreeXML")
        If xmlContent <> "" Then
            LoadXmlDataFromString(xmlContent)
        Else
            MsgBox("Nu s-a putut încărca structura arborelui din Access.", MsgBoxStyle.Exclamation)
            Environment.Exit(0)
        End If
        'End If
    End Sub

    Private Sub AddXmlNodeToTree(xNode As XmlNode, parentCol As TreeNodeCollection)
        ' 1. Instantiem clasa noastra extinsa
        Dim newNode As New ExtendedTreeNode(xNode.Attributes("Text").Value)

        If xNode.Attributes("ID") IsNot Nothing Then newNode.Tag = xNode.Attributes("ID").Value

        ' 2. Citim atributele din XML
        Dim iClosed As String = ""
        Dim iOpen As String = ""
        Dim iExpanded As Boolean

        If xNode.Attributes("IconClosed") IsNot Nothing Then iClosed = xNode.Attributes("IconClosed").Value
        If xNode.Attributes("IconOpen") IsNot Nothing Then iOpen = xNode.Attributes("IconOpen").Value
        If xNode.Attributes("Expanded") IsNot Nothing Then
            Boolean.TryParse(xNode.Attributes("Expanded").Value, iExpanded)
        End If

        ' 3. Setam proprietatile custom
        newNode.IconClosed = iClosed
        ' Daca nu are iconita de Open, o folosim pe cea de Closed (fallback)
        newNode.IconOpen = If(String.IsNullOrEmpty(iOpen), iClosed, iOpen)

        If iExpanded Then
            newNode.ImageKey = newNode.IconOpen
            newNode.SelectedImageKey = newNode.IconOpen
            newNode.Expand()
        End If

        ' 4. Setam starea initiala (Closed)
        If Not String.IsNullOrEmpty(newNode.IconClosed) Then
            newNode.ImageKey = newNode.IconClosed
            newNode.SelectedImageKey = newNode.IconClosed
        End If

        ' 5. Adaugam in colectie
        parentCol.Add(newNode)

        ' 6. Recursivitate
        For Each childNode As XmlNode In xNode.SelectNodes("Node")
            AddXmlNodeToTree(childNode, newNode.Nodes)
        Next
    End Sub

    ' =============================================================
    ' ÎNCĂRCARE XML + CONFIGURARE (CORECTAT)
    ' =============================================================
    Private Overloads Sub LoadXmlData(filePath As String)
        If Not File.Exists(filePath) Then Return

        Try
            Dim xDoc As New XmlDocument()
            xDoc.Load(filePath)

            MyTree.BeginUpdate()
            MyTree.Nodes.Clear()
            MyTree.ImageList = Nothing

            ' 1. CONFIGURARE (Inclusiv RightClickFunc)
            Dim configNode As XmlNode = xDoc.SelectSingleNode("/Tree/Config")
            If configNode IsNot Nothing Then
                AplicareConfigurare(configNode)
            End If

            ' 2. INCARCARE IMAGINI (Base64 -> ImageList)
            Dim imgListNode As XmlNode = xDoc.SelectSingleNode("/Tree/Images")
            If imgListNode IsNot Nothing Then
                LoadImagesFromXml(imgListNode)
            End If

            ' 3. POPULARE NODURI
            For Each xNode As XmlNode In xDoc.SelectNodes("/Tree/Nodes/Node")
                AddXmlNodeToTree(xNode, MyTree.Nodes)
            Next

        Catch ex As Exception
            MessageBox.Show("Err XML: " & ex.Message)
        Finally
            MyTree.EndUpdate()
            MyTree.Refresh()
        End Try
    End Sub

    Private Sub LoadXmlDataFromString(xmlContent As String)
        Try
            Dim xDoc As New XmlDocument()
            xDoc.LoadXml(xmlContent)

            MyTree.BeginUpdate()
            MyTree.Nodes.Clear()
            MyTree.ImageList = Nothing

            ' 1. CONFIGURARE (Inclusiv RightClickFunc)
            Dim configNode As XmlNode = xDoc.SelectSingleNode("/Tree/Config")
            If configNode IsNot Nothing Then
                AplicareConfigurare(configNode)
            End If

            ' 2. INCARCARE IMAGINI (Base64 -> ImageList)
            Dim imgListNode As XmlNode = xDoc.SelectSingleNode("/Tree/Images")
            If imgListNode IsNot Nothing Then
                LoadImagesFromXml(imgListNode)
            End If

            ' 3. POPULARE NODURI
            For Each xNode As XmlNode In xDoc.SelectNodes("/Tree/Nodes/Node")
                AddXmlNodeToTree(xNode, MyTree.Nodes)
            Next

        Catch ex As Exception
            MessageBox.Show("Err XML: " & ex.Message)
        Finally
            MyTree.EndUpdate()
            MyTree.Refresh()
        End Try
    End Sub

    Private Sub LoadImagesFromXml(imgRoot As XmlNode)
        Dim imgList As New ImageList()
        imgList.ColorDepth = ColorDepth.Depth32Bit
        imgList.ImageSize = New Size(16, 16) ' Sau cat vrei tu

        For Each imgNode As XmlNode In imgRoot.SelectNodes("Image")
            Try
                Dim key As String = imgNode.Attributes("Key").Value
                Dim b64 As String = imgNode.InnerText

                If Not String.IsNullOrEmpty(b64) Then
                    Dim bytes As Byte() = Convert.FromBase64String(b64)
                    Using ms As New MemoryStream(bytes)
                        Dim bmp As Image = Image.FromStream(ms)
                        imgList.Images.Add(key, bmp)
                    End Using
                End If
            Catch
            End Try
        Next

        If imgList.Images.Count > 0 Then
            MyTree.ImageList = imgList
        End If
    End Sub

    Private Sub AplicareConfigurare(cfg As XmlNode)
        Try
            ' Folosim InvariantCulture pentru a citi corect "9.5" indiferent de setările PC-ului
            Dim culture As CultureInfo = CultureInfo.InvariantCulture

            ' --- BackColor ---
            If cfg.Attributes("BackColor") IsNot Nothing Then
                Try
                    Dim c As Color = ColorTranslator.FromHtml(cfg.Attributes("BackColor").Value)
                    MyTree.BackColor = c
                    Me.BackColor = c
                Catch
                End Try
            End If

            ' --- ForeColor ---
            If cfg.Attributes("ForeColor") IsNot Nothing Then
                Try
                    MyTree.ForeColor = ColorTranslator.FromHtml(cfg.Attributes("ForeColor").Value)
                Catch
                End Try
            End If

            ' --- Font ---
            Dim fName As String = "Segoe UI"
            Dim fSize As Single = 9.0F

            ' Păstrăm valorile curente dacă există
            If MyTree.Font IsNot Nothing Then
                fName = MyTree.Font.Name
                fSize = MyTree.Font.Size
            End If

            If cfg.Attributes("FontName") IsNot Nothing Then fName = cfg.Attributes("FontName").Value

            If cfg.Attributes("FontSize") IsNot Nothing Then
                ' Folosim TryParse cu InvariantCulture pentru siguranță
                Single.TryParse(cfg.Attributes("FontSize").Value, NumberStyles.Any, culture, fSize)
            End If

            MyTree.Font = New Font(fName, fSize)

            ' --- Indent ---
            If cfg.Attributes("Indent") IsNot Nothing Then
                Dim ind As Integer = 19
                Dim unused = Integer.TryParse(cfg.Attributes("Indent").Value, ind)
                MyTree.Indent = ind
            End If

            ' --- ItemHeight ---
            If cfg.Attributes("ItemHeight") IsNot Nothing Then
                Dim ih As Integer = 20
                Dim unused1 = Integer.TryParse(cfg.Attributes("ItemHeight").Value, ih)
                If ih > 0 Then MyTree.ItemHeight = ih
            End If

        Catch ex As Exception
            ' Ignorăm erorile de configurare minore
        End Try
    End Sub

    ' =============================================================
    ' LOGICA MOUSE SIMPLIFICATĂ (Fără Timer)
    ' =============================================================

    Private Sub MyTree_NodeMouseUp(pNode As TreeNode, pE As MouseEventArgs) Handles MyTree.NodeMouseUp
        ' A. CLICK STANGA -> Trimitem INSTANT la Access
        If pE.Button = MouseButtons.Left Then
            TrimiteMesajAccess(pNode)
        End If

        ' B. CLICK DREAPTA -> Execuție Funcție Custom (Context Menu)
        If pE.Button = MouseButtons.Right Then
            ' Selectăm vizual nodul (fiind click dreapta, treeview nu-l selectează automat întotdeauna)
            MyTree.SelectedNode = pNode

            If Not String.IsNullOrEmpty(_RightClickFunc) AndAlso _accessApp IsNot Nothing Then
                Try
                    Dim nodeId As String = If(pNode.Tag IsNot Nothing, pNode.Tag.ToString(), "")
                    ' Apelăm funcția VBA: NumeFunctie(TreeID, NodeID)
                    _accessApp.Run(_RightClickFunc, _idTree, nodeId)
                Catch
                    ' Erorile de VBA nu trebuie să crape .NET-ul
                End Try
            End If
        End If
    End Sub

    Private Sub MyTree_DoubleClick(sender As Object, e As EventArgs) Handles MyTree.DoubleClick
        ' Nu mai facem nimic aici. 
        ' TreeView-ul se va expanda/strânge automat (comportament nativ).
        ' Access-ul a primit deja notificarea de Click la primul MouseUp.
    End Sub

    ' Helper pentru a nu duplica codul de comunicare
    Private Sub TrimiteMesajAccess(pNode As TreeNode)
        If _accessApp IsNot Nothing Then
            Try
                Dim nodeId As String = If(pNode.Tag IsNot Nothing, pNode.Tag.ToString(), "")

                ' Trimitem ID-ul și Textul la Access
                _accessApp.Run("OnTreeEvent", _idTree, nodeId, pNode.Text)
            Catch
            End Try
        End If
    End Sub
    Private Sub MyTree_AfterExpand(sender As Object, e As TreeViewEventArgs) Handles MyTree.AfterExpand
        ' Facem cast la tipul nostru
        Dim myNode As ExtendedTreeNode = TryCast(e.Node, ExtendedTreeNode)

        If myNode IsNot Nothing AndAlso Not String.IsNullOrEmpty(myNode.IconOpen) Then
            myNode.ImageKey = myNode.IconOpen
            myNode.SelectedImageKey = myNode.IconOpen
        End If
    End Sub

    Private Sub MyTree_AfterCollapse(sender As Object, e As TreeViewEventArgs) Handles MyTree.AfterCollapse
        Dim myNode As ExtendedTreeNode = TryCast(e.Node, ExtendedTreeNode)

        If myNode IsNot Nothing AndAlso Not String.IsNullOrEmpty(myNode.IconClosed) Then
            myNode.ImageKey = myNode.IconClosed
            myNode.SelectedImageKey = myNode.IconClosed
        End If
    End Sub

    Private Sub AddXmlNodeToTree_Custom(xNode As XmlNode, pParent As AdvancedTreeControl.TreeItem)
        Dim text As String = xNode.Attributes("Text").Value
        Dim idVal As String = If(xNode.Attributes("ID") IsNot Nothing, xNode.Attributes("ID").Value, "")

        Dim it As AdvancedTreeControl.TreeItem =
        MyTree.AddItem(text, pParent)

        it.Tag = idVal   ' EXTINDE TreeItem cu Tag (vezi NOTĂ jos)

        If xNode.Attributes("Expanded") IsNot Nothing Then
            Dim exp As Boolean
            If Boolean.TryParse(xNode.Attributes("Expanded").Value, exp) Then
                it.Expanded = exp
            End If
        End If

        For Each child As XmlNode In xNode.SelectNodes("Node")
            AddXmlNodeToTree_Custom(child, it)
        Next
    End Sub

End Class