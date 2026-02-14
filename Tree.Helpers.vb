Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text.Json
Imports System.Diagnostics

Partial Public Class Tree
    Private Sub PositioneazaInParent()
        Dim rParent As RECT
        GetClientRect(_formHwnd, rParent)
        If Me.Width <> (rParent.Right - rParent.Left) OrElse Me.Height <> (rParent.Bottom - rParent.Top) Then
            MoveWindow(Me.Handle, 0, 0, rParent.Right - rParent.Left, rParent.Bottom - rParent.Top, True)
        End If
    End Sub

    Private Sub ConecteazaLaAccess(hwndAccess As IntPtr)
        Dim guidIDispatch As New Guid("{00020400-0000-0000-C000-000000000046}") ' IID_IDispatch
        Dim obj As Object = Nothing

        ' Această funcție returnează obiectul "Window" din modelul de obiecte Access
        Dim hr As Integer = AccessibleObjectFromWindow(hwndAccess, OBJID_NATIVEOM, guidIDispatch, obj)

        If hr >= 0 AndAlso obj IsNot Nothing Then
            Try
                ' Din obiectul Window, urcăm la Application
                Dim windowObj As Object = obj
                _accessApp = windowObj.Application
                TreeLogger.Debug("Conectat la Access. Versiune Access: " & _accessApp.Version, "ConecteazaLaAccess", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                TreeLogger.Ex(ex, "ConecteazaLaAccess", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End Try
        Else
            TreeLogger.Err("Nu s-a putut obține obiectul COM din HWND.", "ConecteazaLaAccess", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        End If
    End Sub

    Private Function GetValoareLocala(numeControl As String) As String
        ' 1. Găsim formularul de care suntem lipiți (ParentHwnd)
        Dim targetForm As Object = GetFormObjectFromHwnd(_formHwnd)

        If targetForm Is Nothing Then
            TreeLogger.Err($"Form object not found for HWND:{_formHwnd}", "GetValoareLocala", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return "Form not found"
        End If

        ' 2. Citim controlul direct din acest formular
        Try
            Dim ctl As Object = targetForm.Controls(numeControl)
            If ctl Is Nothing Then Return "Control missing"

            Dim val As Object = ctl.Value
            If val Is Nothing Then Return ""

            TreeLogger.Debug($"Valoare control '{numeControl}' este: {val}", "GetValoareLocala", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Return val.ToString()
        Catch ex As Exception
            TreeLogger.Ex(ex, "GetValoareLocala", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return ""
        End Try
    End Function

    Private Sub CurataResurseSiIesi()
        If _cleaningDone Then Return
        _cleaningDone = True

        ' 1. Oprire Timer
        _MonitorTimer?.Stop()

        ' 2. Eliberare Access COM (foarte important cu Try/Catch)
        If _accessApp IsNot Nothing Then
            Try
                ' Eliberam referinta COM
                Marshal.ReleaseComObject(_accessApp)
                TreeLogger.Debug("COM object Access eliberat cu succes.", "CurataResurseSiIesi", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                ' Aceasta eroare e normala daca Access s-a inchis deja (RPC unavailable)
                TreeLogger.Ex(ex, "CurataResurseSiIesi", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            _accessApp = Nothing
        End If

        GC.Collect()
        GC.Collect()

        TreeLogger.Debug("Curățare resurse completă. Iesire din aplicatie.", "CurataResurseSiIesi", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub TrimiteMesajAccess(Action As String, pItem As AdvancedTreeControl.TreeItem, Optional ExtraInfo As String = "")
        TreeLogger.Debug($"TrimiteMesajAccess: Action='{Action}', ItemKey='{If(pItem IsNot Nothing, pItem.Key.ToString(), "null")}', ExtraInfo='{ExtraInfo}'", "TrimiteMesajAccess", MessageBoxButtons.OK, MessageBoxIcon.Information)

        If pItem Is Nothing Then
            If _accessApp IsNot Nothing Then
                Try
                    Me.BeginInvoke(Sub()
                                       If _formHwnd <> IntPtr.Zero Then
                                           _accessApp.Run("OnTreeEvent", _idTree, Action, "", "", ExtraInfo)
                                       End If
                                   End Sub)
                Catch ex As Exception
                    TreeLogger.Ex(ex, "TrimiteMesajAccess", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        Else
            Dim nodeKey As String = If(pItem IsNot Nothing, pItem.Key.ToString(), "")
            Dim nodeCaption As String = If(pItem IsNot Nothing, pItem.Caption, "")

            If _accessApp IsNot Nothing Then
                Try
                    Me.BeginInvoke(Sub()
                                       Try
                                           If _formHwnd <> IntPtr.Zero Then
                                               _accessApp.Run("OnTreeEvent", _idTree, Action, nodeKey, nodeCaption, ExtraInfo)
                                           End If
                                       Catch ex As Exception
                                           TreeLogger.Debug("Err TrimiteMesajAccess Inner: " & ex.Message, "TrimiteMesajAccess")
                                       End Try
                                   End Sub)
                Catch ex As Exception
                    TreeLogger.Ex(ex, "TrimiteMesajAccess", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End If
    End Sub

    ' =============================================================
    ' PROCESARE COMENZI EXTERNE
    ' =============================================================
    Private Sub ProcesareComandaAccess(cmd As String)
        Try
            ' Format protocol: ACTION||Param1||Param2||...
            Dim parts() As String = cmd.Split(inCommandSeparator, StringSplitOptions.None)

            If parts.Length < 1 Then Return

            TreeLogger.Debug(cmd, "ProcesareComandaAccess") ' Logăm comanda primită pentru debugging

            Select Case parts(0).ToUpper()
                Case "FORCE_EXPAND"
                    ' Format: FORCE_EXPAND||NodeID
                    If parts.Length >= 2 Then
                        Dim nodeId As String = parts(1)
                        Dim foundNode As AdvancedTreeControl.TreeItem = Nothing

                        ' Iterăm prin rădăcini pentru a găsi nodul (fix pentru eroarea cu 'root')
                        For Each rootItem In MyTree.Items
                            foundNode = FindNodeByIdRecursive(rootItem, nodeId)
                            If foundNode IsNot Nothing Then Exit For
                        Next

                        If foundNode IsNot Nothing Then
                            foundNode.Expanded = True
                            MyTree.Invalidate()
                        End If
                    End If

                Case "ADD_BATCH_JSON"
                    ' Format: ADD_BATCH_JSON||ParentID||JsonString
                    If parts.Length >= 3 Then
                        ExecuteAddBatchJson(parts(1), parts(2))
                    End If

                Case "ADD_BATCH_FILE"
                    ' Format: ADD_BATCH_FILE||ParentID||FilePath
                    If parts.Length >= 3 Then
                        ExecuteAddBatchFile(parts(1), parts(2))
                    End If

                Case "ENABLE"
                    ' Format: "ENABLE||1/0"
                    If parts.Length >= 2 Then
                        Dim enable As Boolean = (parts(1) = "1")
                        MyTree.Enabled = enable
                        TrimiteMesajAccess("Enabled", Nothing)
                    End If

                Case "REFRESH"
                    ' Format: "REFRES||xml_to_use_in_refresh_path"
                    If parts.Length >= 2 Then
                        'MyTree.Clear()
                        If ReLoadXmlData(parts(1)) Then
                            TrimiteMesajAccess("Refreshed", Nothing)
                        End If
                    End If

                Case "CLEAR_NODES"
                    ' Format: CLEAR_NODES
                    MyTree.Clear()
                    MyTree.Refresh()

                Case "FIND_NODE"
                    ' Format: FIND_NODE||Caption||MatchExact(1/0)||Scroll(1/0)||Click(1/0)
                    If parts.Length >= 5 Then
                        Dim textToFind As String = parts(1)
                        Dim matchExact As Boolean = (parts(2) = "1")
                        Dim scrollToView As Boolean = (parts(3) = "1")
                        Dim raiseClick As Boolean = (parts(4) = "1")

                        FindAndSelectNode(textToFind, matchExact, scrollToView, raiseClick)
                    End If
                Case "ADD_NODE"
                    ' Format: ADD_NODE||ParentID||NewNodeID||Caption||IconKey
                    Dim iconKey As String = ""
                    If parts.Length > 4 Then iconKey = parts(4)
                    ExecuteAddNode(parts(1), parts(2), parts(3), iconKey)

                Case "REMOVE_NODE"
                    ' Format: REMOVE_NODE||NodeID
                    ExecuteRemoveNode(parts(1))

                Case "GET_PROPERTY"
                    ' Format: GET_PROPERTY||PropertyName||[Optional:NodeID]
                    Dim propValue As String = MyTree.ProcessPropertyRequest(cmd)
                    ' Trimitem inapoi la Access prin Run (sau altă metodă, depinde de implementare)
                    If _accessApp IsNot Nothing Then
                        Me.BeginInvoke(Sub()
                                           If _formHwnd <> IntPtr.Zero Then
                                               TrimiteMesajAccess("PropertyValue", Nothing, propValue)
                                               '_accessApp.Run("OnTreeEvent", _idTree, "PropertyValue", "", "", propValue)
                                           End If
                                       End Sub)
                    End If

                Case "SELECT_NODE"
                    ' Format: SELECT_NODE||NodeID
                    ' Selectează vizual nodul pe baza cheii unice (Key)
                    If parts.Length >= 2 Then
                        Dim nodeId As String = parts(1)
                        Dim foundNode As AdvancedTreeControl.TreeItem = Nothing

                        ' Cautam nodul in toata ierarhia
                        For Each root In MyTree.Items
                            foundNode = FindNodeByIdRecursive(root, nodeId)
                            If foundNode IsNot Nothing Then Exit For
                        Next

                        If foundNode IsNot Nothing Then
                            ' Selectam
                            MyTree.SelectedNode = foundNode
                            ' Expandam parintii sa se vada
                            Dim parent As AdvancedTreeControl.TreeItem = foundNode.Parent
                            While parent IsNot Nothing
                                parent.Expanded = True
                                parent = parent.Parent
                            End While
                            ' Scroll la el
                            ScrollToNode(foundNode)
                            MyTree.Invalidate()

                            'TrimiteMesajAccess("Click", foundNode)
                        End If
                    End If

                Case "SET_CHECKBOX"
                    ' Format: SET_CHECKBOX||NodeID||State(0/1)
                    If parts.Length >= 3 Then
                        Dim nodeId As String = parts(1)
                        Dim stateInt As Integer = 0
                        Dim v = Integer.TryParse(parts(2), stateInt) ' 0 = Unchecked, 1 = Checked

                        Dim foundNode As AdvancedTreeControl.TreeItem = Nothing
                        For Each root In MyTree.Items
                            foundNode = FindNodeByIdRecursive(root, nodeId)
                            If foundNode IsNot Nothing Then Exit For
                        Next

                        If foundNode IsNot Nothing Then
                            Dim newState As AdvancedTreeControl.TreeCheckState
                            If stateInt = 1 Then
                                newState = AdvancedTreeControl.TreeCheckState.Checked
                            Else
                                newState = AdvancedTreeControl.TreeCheckState.Unchecked
                            End If

                            MyTree.SetItemCheckState(foundNode, newState)
                        End If
                    End If
            End Select

        Catch ex As Exception
            ' Ignorăm erorile de parsare silențios sau le logăm
            TreeLogger.Ex(ex, "ProcesareComandaAccess", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =============================================================
    ' 1. LOGICA ADĂUGARE NOD (Root sau Child)
    ' =============================================================
    Private Sub ExecuteAddNode(parentId As String, newId As String, text As String, iconKey As String)
        Dim parentNode As AdvancedTreeControl.TreeItem = Nothing
        Dim iconImg As Image = Nothing

        Try
            ' 1. Gestionare Imagine (Robustness: Fallback dacă nu există cheia)
            If Not String.IsNullOrEmpty(iconKey) Then
                If Not _imageCache.TryGetValue(iconKey, iconImg) Then
                    ' Opțional: Logare eroare sau folosire imagine default
                    ' iconImg = _defaultImage 
                End If
            End If

            ' 2. Determinăm Părintele
            If Not String.IsNullOrEmpty(parentId) Then
                ' Folosim funcția de căutare după ID, nu Text
                For Each root In MyTree.Items
                    parentNode = FindNodeByIdRecursive(root, parentId)
                    If parentNode IsNot Nothing Then Exit For
                Next

                If parentNode Is Nothing Then
                    If DEBUG_MODE Then MessageBox.Show($"Parent ID '{parentId}' not found. Cannot add child '{newId}'.", "EROARE", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            ' 3. Creăm nodul
            Dim newItem As New AdvancedTreeControl.TreeItem With {
                .Caption = text,
                .Key = newId,
                .LeftIconClosed = iconImg,
                .LeftIconOpen = iconImg, ' Asigurăm și iconița de Open
                .Expanded = False ' Implicit collapse la nodurile noi dinamice
            }

            If parentNode Is Nothing Then
                ' -- ROOT NODE --
                newItem.Level = 0
                MyTree.Items.Add(newItem)
            Else
                ' -- CHILD NODE --
                newItem.Level = parentNode.Level + 1
                newItem.Parent = parentNode
                parentNode.Children.Add(newItem)

                ' FOARTE IMPORTANT: Expandăm părintele ca să vedem ce am adăugat
                parentNode.Expanded = True
            End If

            ' 4. ACTUALIZARE VIZUALĂ COMPLETĂ
            ' Recalculăm scroll-ul și forțăm redesenarea
            MyTree.SetAutoHeight() ' Sau logica ta de recalculare înălțime totală
            MyTree.Invalidate()

            ' Opțional: Scroll până la noul element creat (User Experience)
            ScrollToNode(newItem)

        Catch ex As Exception
            TreeLogger.Ex(ex, "ExecuteAddNode", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' =============================================================
    ' 2. LOGICA ȘTERGERE NOD
    ' =============================================================
    Private Sub ExecuteRemoveNode(nodeId As String)
        Dim removed As Boolean

        ' Încercăm să ștergem din rădăcini
        removed = RemoveNodeFromList(MyTree.Items, nodeId)

        ' Actualizare vizuală
        If removed Then
            MyTree.Refresh()
        Else
            TreeLogger.Debug("Node ID to remove not found: " & nodeId, "ExecuteRemoveNode", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    ' Helper recursiv pentru ștergere
    Private Function RemoveNodeFromList(list As List(Of AdvancedTreeControl.TreeItem), targetId As String) As Boolean
        ' Iterăm invers pentru a putea șterge sigur
        For i As Integer = list.Count - 1 To 0 Step -1
            Dim it = list(i)

            ' Verificăm ID-ul (_tag)
            Dim currentId As String = If(it.Key IsNot Nothing, it.Key.ToString(), "")

            If currentId = targetId Then
                ' Am găsit nodul, îl ștergem (dispare el și toți copiii)
                list.RemoveAt(i)
                Return True
            End If

            ' Dacă nu e el, căutăm în copiii lui
            If it.Children.Count > 0 Then
                If RemoveNodeFromList(it.Children, targetId) Then
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Function FindNodeByIdRecursive(item As AdvancedTreeControl.TreeItem, targetId As String) As AdvancedTreeControl.TreeItem
        Dim currentId As String = If(item.Key IsNot Nothing, item.Key.ToString(), "")
        If currentId = targetId Then Return item

        For Each child In item.Children
            Dim res = FindNodeByIdRecursive(child, targetId)
            If res IsNot Nothing Then Return res
        Next
        Return Nothing
    End Function

    '============================================================
    ' 3. LOGICA GĂSIRE ȘI SELECTARE NOD
    '============================================================
    Private Sub FindAndSelectNode(text As String, matchExact As Boolean, scroll As Boolean, doClick As Boolean)
        If MyTree.Items.Count = 0 Then Return

        Dim foundNode As AdvancedTreeControl.TreeItem = Nothing

        ' 1. Căutare Recursivă
        For Each rootItem In MyTree.Items
            foundNode = SearchNodeRecursive(rootItem, text, matchExact)
            If foundNode IsNot Nothing Then Exit For
        Next

        If foundNode IsNot Nothing Then
            ' 2. Expandăm părinții ca să fie vizibil
            Dim parent As AdvancedTreeControl.TreeItem = foundNode.Parent
            While parent IsNot Nothing
                parent.Expanded = True
                parent = parent.Parent
            End While

            ' Fortam redesenarea pentru a actualiza lista vizibila dupa expandare
            MyTree.SelectedNode = foundNode
            MyTree.Invalidate()
            Application.DoEvents()

            ' 4. Scroll to View (Calcul manual al poziției Y)
            If scroll Then
                ScrollToNode(foundNode)
            End If

            ' 5. Trimite eveniment inapoi la Access (Simulare Click)
            If doClick Then
                TrimiteMesajAccess("MouseUp", foundNode)
            End If
        Else
            TreeLogger.Debug($"Node with text '{text}' not found.", "FindAndSelectNode", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Function SearchNodeRecursive(item As AdvancedTreeControl.TreeItem, text As String, exact As Boolean) As AdvancedTreeControl.TreeItem
        ' Verificare curentă
        Dim isMatch As Boolean
        If exact Then
            isMatch = (String.Compare(item.Caption, text, True) = 0)
        Else
            isMatch = (item.Caption.Contains(text, StringComparison.OrdinalIgnoreCase))
        End If

        If isMatch Then Return item

        ' Căutare în copii
        For Each child In item.Children
            Dim result = SearchNodeRecursive(child, text, exact)
            If result IsNot Nothing Then Return result
        Next

        Return Nothing
    End Function

    ' Înlocuiește funcția ScrollToNode cu aceasta:
    Private Sub ScrollToNode(node As AdvancedTreeControl.TreeItem)
        ' 1. Calculăm lista vizibilă actualizată
        Dim list As New List(Of AdvancedTreeControl.TreeItem)
        For Each it In MyTree.Items
            GetVisibleListRecursive(it, list)
        Next

        Dim visibleIndex As Integer = list.IndexOf(node)

        If visibleIndex >= 0 Then
            ' 2. Forțăm actualizarea înălțimii totale înainte de a muta scroll-ul
            Dim totalHeight As Integer = list.Count * MyTree.ItemHeight

            If MyTree.AutoScrollMinSize.Height <> totalHeight Then
                MyTree.AutoScrollMinSize = New Size(0, totalHeight)
            End If

            ' 3. Calculăm poziția Y și centrăm (opțional)
            Dim yPos As Integer = visibleIndex * MyTree.ItemHeight

            ' Centrare vizuală (ca să nu fie lipit de marginea de sus)
            yPos = Math.Max(0, yPos - (MyTree.Height \ 2) + (MyTree.ItemHeight \ 2))

            ' 4. Aplicăm scroll-ul și redesenăm
            MyTree.AutoScrollPosition = New Point(0, yPos)
            MyTree.Invalidate()
        End If
    End Sub

    Private Sub GetVisibleListRecursive(item As AdvancedTreeControl.TreeItem, list As List(Of AdvancedTreeControl.TreeItem))
        list.Add(item)
        If item.Expanded Then
            For Each child In item.Children
                GetVisibleListRecursive(child, list)
            Next
        End If
    End Sub

    ' =============================================================
    ' 4. LOGICA BATCH PROCESSING (NOU)
    ' =============================================================

    Private Sub ExecuteAddBatchFile(parentID As String, filePath As String)
        If File.Exists(filePath) Then
            Try
                Dim jsonContent As String = File.ReadAllText(filePath)
                ExecuteAddBatchJson(parentID, jsonContent)

                ' Opțional: curățăm fișierul temporar creat de VBA
                ' File.Delete(filePath) 
            Catch ex As Exception
                TreeLogger.Ex(ex, "EROARE", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub ExecuteAddBatchJson(parentID As String, jsonString As String)
        Try
            ' 1. Găsim părintele
            Dim parentNode As AdvancedTreeControl.TreeItem = Nothing

            If Not String.IsNullOrEmpty(parentID) Then
                For Each root In MyTree.Items
                    parentNode = FindNodeByIdRecursive(root, parentID)
                    If parentNode IsNot Nothing Then Exit For
                Next

                ' Dacă am cerut un părinte și nu există, ieșim (safety)
                If parentNode Is Nothing Then
                    TreeLogger.Err(parentID & " not found for batch add.", "ExecuteAddBatchJson", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If
            End If

            Try
                ' 2. Deserializare
                ' Opțiuni pentru a fi permisivi cu JSON-ul (Case Insensitive)
                Dim options As New JsonSerializerOptions With {
                .PropertyNameCaseInsensitive = True
            }
                Dim newNodes As List(Of NodeDto) = JsonSerializer.Deserialize(Of List(Of NodeDto))(jsonString, options)

                If newNodes Is Nothing OrElse newNodes.Count = 0 Then Return

                ' 3. OPRIM DESENAREA (PERFORMANȚĂ CRITICĂ)
                MyTree.SuspendLayout()

                ' Dacă părintele are copii și primul este un Loader, îi ștergem pe toți
                If parentNode IsNot Nothing AndAlso parentNode.Children.Count > 0 Then
                    If parentNode.Children(0).IsLoader Then
                        parentNode.Children.Clear()
                    End If
                End If

                ' 4. Adăugare recursivă
                For Each nodeDto In newNodes
                    AddNodeDtoToTree(nodeDto, parentNode)
                Next

                ' Dacă am adăugat la un nod existent, îl expandăm
                If parentNode IsNot Nothing Then parentNode.Expanded = True

            Catch ex As Exception
                TreeLogger.Ex(ex, "ExecuteAddBatchJson", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        Catch ex As Exception
            TreeLogger.Ex(ex, "ExecuteAddBatchJson", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ' 5. REPORNIM DESENAREA
            MyTree.SetAutoHeight()
            MyTree.ResumeLayout()
            MyTree.Invalidate()
        End Try
    End Sub

    Private Sub AddNodeDtoToTree(dto As NodeDto, parentItem As AdvancedTreeControl.TreeItem)
        ' Rezolvare imagine
        Dim iconImgClosed As Image = Nothing
        If Not String.IsNullOrEmpty(dto.IconClosed) Then
            _imageCache.TryGetValue(dto.IconClosed, iconImgClosed)
        End If

        Dim iconImgOpen As Image = Nothing
        If Not String.IsNullOrEmpty(dto.IconOpen) Then
            _imageCache.TryGetValue(dto.IconOpen, iconImgOpen)
        End If

        Dim iconImgRight As Image = Nothing
        If Not String.IsNullOrEmpty(dto.IconRight) Then
            _imageCache.TryGetValue(dto.IconRight, iconImgRight)
        End If

        ' --- CONVERSIE SIGURĂ BOOLEAN ---
        ' Verificăm ce am primit în Object și transformăm în Boolean curat
        Dim isExpanded As Boolean = False
        If dto.Expanded IsNot Nothing Then
            Dim s As String = dto.Expanded.ToString().ToLower()
            ' Acceptăm 1, -1, "1", "true" ca fiind TRUE. Restul e False.
            isExpanded = (s = "1" OrElse s = "-1" OrElse s = "true")
        End If

        Dim isLazyNode As Boolean = False
        If dto.LazyNode IsNot Nothing Then
            Dim s As String = dto.LazyNode.ToString().ToLower()
            ' Acceptăm 1, -1, "1", "true" ca fiind TRUE. Restul e False.
            isLazyNode = (s = "1" OrElse s = "-1" OrElse s = "true")
        End If
        ' --------------------------------

        ' Creare Nod UI
        Dim newItem As New AdvancedTreeControl.TreeItem With {
            .Key = dto.Key,
            .Caption = dto.Caption,
            .LeftIconClosed = iconImgClosed,
            .LeftIconOpen = iconImgClosed,
            .RightIcon = iconImgRight,
            .Tag = dto.Tag,
            .Expanded = isExpanded,
            .LazyNode = isLazyNode
        }

        ' Atribute vizuale
        If dto.Bold IsNot Nothing Then
            Dim s As String = dto.Bold.ToString().ToLower()
            newItem.Bold = (s = "1" OrElse s = "-1" OrElse s = "true")
        End If
        If dto.Italic IsNot Nothing Then
            Dim s As String = dto.Italic.ToString().ToLower()
            newItem.Italic = (s = "1" OrElse s = "-1" OrElse s = "true")
        End If
        If Not String.IsNullOrEmpty(dto.ForeColor) Then
            Try
                If dto.ForeColor.StartsWith("#"c) Then
                    newItem.NodeForeColor = ColorTranslator.FromHtml(dto.ForeColor)
                Else
                    newItem.NodeForeColor = Color.FromName(dto.ForeColor)
                End If
            Catch : End Try
        End If

        ' Linkare la părinte sau root
        If parentItem Is Nothing Then
            newItem.Level = 0
            MyTree.Items.Add(newItem)
        Else
            newItem.Level = parentItem.Level + 1
            newItem.Parent = parentItem
            parentItem.Children.Add(newItem)
        End If

        ' Procesare copii (Recursivitate)
        If dto.Children IsNot Nothing AndAlso dto.Children.Count > 0 Then
            For Each childDto In dto.Children
                AddNodeDtoToTree(childDto, newItem)
            Next
        End If
    End Sub

    ' Aceasta înlocuiește SetupEvents și OnRequestLazyLoad-ul anterior
    Private Sub MyTree_RequestLazyLoad(sender As Object, item As AdvancedTreeControl.TreeItem) Handles MyTree.RequestLazyLoad
        ' Trimitem cererea la VBA: "BEFORE_EXPAND||NodeID"
        TrimiteMesajAccess("BEFORE_EXPAND", item)
    End Sub

    ' =============================================================
    ' WNDPROC - INTERCEPTARE DISTRUGERE FORTATA
    ' =============================================================
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_SETTEXT As Integer = &HC

        If m.Msg = WM_SETTEXT Then
            ' 1. Citim mesajul venit din Access
            Dim messageFromAccess As String = Marshal.PtrToStringAuto(m.LParam) ' Folosim Ansi pt ca VBA SendMessageA trimite Ansi

            ' 2. Procesăm comanda
            If Not String.IsNullOrEmpty(messageFromAccess) Then
                ' Verificăm dacă este o comandă complexă cu delimitator
                If messageFromAccess.Contains("||") Then
                    ProcesareComandaAccess(messageFromAccess)
                End If
            End If
        End If

        ' Interceptare distrugere
        If m.Msg = WM_DESTROY Then
            Dim st As New StackTrace(True)
            Dim stackInfo As String = st.ToString()
            TreeLogger.Debug("WM_DESTROY received.", "WndProc", MessageBoxButtons.OK, MessageBoxIcon.Information)
            If Not _cleaningDone Then
                CurataResurseSiIesi()
            End If
        End If
        'TreeLogger.Debug($"WndProc received. Msg: {m.Msg}, LParam: {m.LParam}", "WndProc", MessageBoxButtons.OK, MessageBoxIcon.Information)
        MyBase.WndProc(m)
    End Sub
End Class

