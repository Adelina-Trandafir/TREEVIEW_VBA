Imports System.Runtime.InteropServices

Partial Public Class Tree
    Private Sub MonitorTimerHandle()
        If _formHwnd = IntPtr.Zero Then Return

        ' Daca fereastra parinte (Access) NU mai exista
        If Not IsWindow(_formHwnd) Then
            _MonitorTimer.Stop()
            CurataResurseSiIesi()
            Application.Exit()
            Return
        End If

        PositioneazaInParent()
    End Sub

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
                'txtLog.AppendText("Conexiune COM reușită la instanța Access specifică!" & vbCrLf)
            Catch ex As Exception
                MsgBox("Eroare la obținerea Application din Window: " & ex.Message)
                Application.Exit()
            End Try
        Else
            MsgBox("Nu s-a putut obține obiectul COM din HWND.")
            Application.Exit()
        End If
    End Sub

    Private Function GetValoareLocala(numeControl As String) As String
        ' 1. Găsim formularul de care suntem lipiți (ParentHwnd)
        Dim targetForm As Object = GetFormObjectFromHwnd(_formHwnd)

        If targetForm Is Nothing Then
            Return "Form not found"
        End If

        ' 2. Citim controlul direct din acest formular
        Try
            Dim ctl As Object = targetForm.Controls(numeControl)
            If ctl Is Nothing Then Return "Control missing"

            Dim val As Object = ctl.Value
            If val Is Nothing Then Return ""

            Return val.ToString()
        Catch ex As Exception
            Return "Err: " & ex.Message
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
            Catch ex As Exception
                ' Aceasta eroare e normala daca Access s-a inchis deja (RPC unavailable)
                MsgBox("COM Cleanup Info (Access probabil inchis deja): " & ex.Message)
            End Try
            _accessApp = Nothing
        End If

        GC.Collect()
        GC.Collect()
    End Sub

    Private Sub TrimiteMesajAccess(Action As String, pItem As AdvancedTreeControl.TreeItem, Optional ExtraInfo As String = "")
        Dim nodeId As String = If(pItem.Tag IsNot Nothing, pItem.Tag.ToString(), "")
        If _accessApp IsNot Nothing Then
            Try
                Me.BeginInvoke(Sub()
                                   If _formHwnd <> IntPtr.Zero Then
                                       _accessApp.Run("OnTreeEvent", _idTree, "MouseUp", nodeId, pItem.Text, ExtraInfo)
                                   End If
                               End Sub)
            Catch ex As Exception
                MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "TrimiteMesajAccess")
            End Try
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

            Select Case parts(0).ToUpper()
                Case "FIND_NODE"
                    ' Format: FIND_NODE||Text||MatchExact(1/0)||Scroll(1/0)||Click(1/0)
                    If parts.Length >= 5 Then
                        Dim textToFind As String = parts(1)
                        Dim matchExact As Boolean = (parts(2) = "1")
                        Dim scrollToView As Boolean = (parts(3) = "1")
                        Dim raiseClick As Boolean = (parts(4) = "1")

                        FindAndSelectNode(textToFind, matchExact, scrollToView, raiseClick)
                    End If
                Case "ADD_NODE"
                    ' Format: ADD_NODE||ParentID||NewNodeID||Text||IconKey
                    Dim iconKey As String = ""
                    If parts.Length > 4 Then iconKey = parts(4)
                    ExecuteAddNode(parts(1), parts(2), parts(3), iconKey)

                Case "REMOVE_NODE"
                    ' Format: REMOVE_NODE||NodeID
                    ExecuteRemoveNode(parts(1))
            End Select

        Catch ex As Exception
            ' Ignorăm erorile de parsare silențios sau le logăm
            If DEBUG_MODE Then MsgBox("Err ProcessCmd: " & ex.Message)
        End Try
    End Sub

    ' =============================================================
    ' 1. LOGICA ADĂUGARE NOD (Root sau Child)
    ' =============================================================
    Private Sub ExecuteAddNode(parentId As String, newId As String, text As String, iconKey As String)
        Dim parentNode As AdvancedTreeControl.TreeItem = Nothing
        Dim iconImg As Image = Nothing

        ' 1. Găsim imaginea (dacă există cheie)
        If Not String.IsNullOrEmpty(iconKey) AndAlso _imageCache.ContainsKey(iconKey) Then
            iconImg = _imageCache(iconKey)
        End If

        ' 2. Determinăm Părintele
        If Not String.IsNullOrEmpty(parentId) Then
            ' Căutăm recursiv părintele
            For Each root In MyTree.Items
                parentNode = SearchNodeRecursive(root, parentId, True) ' Căutare după ID (presupunem că SearchNodeRecursive caută după text sau ID? Vezi nota de jos*)
                ' *NOTA: Funcția ta SearchNodeRecursive caută după TEXT. 
                ' Trebuie o funcție mică de căutare după ID (Tag). O scriu mai jos (FindNodeById).

                parentNode = FindNodeByIdRecursive(root, parentId)
                If parentNode IsNot Nothing Then Exit For
            Next

            If parentNode Is Nothing Then
                ' Dacă s-a cerut părinte dar nu există, ieșim (sau adăugăm ca root, depinde de logică)
                If DEBUG_MODE Then MsgBox("Parent ID not found: " & parentId)
                Return
            End If
        End If

        ' 3. Creăm și adăugăm nodul
        Dim newItem As New AdvancedTreeControl.TreeItem()
        newItem.Text = text
        newItem.Tag = newId
        newItem.LeftIcon = iconImg
        newItem.Expanded = True ' Implicit expandat

        If parentNode Is Nothing Then
            ' Adăugare ca ROOT
            newItem.Level = 0
            MyTree.Items.Add(newItem)
        Else
            ' Adăugare ca CHILD
            newItem.Level = parentNode.Level + 1
            newItem.Parent = parentNode
            parentNode.Children.Add(newItem)
            parentNode.Expanded = True ' Expandăm părintele să se vadă copilul
        End If

        ' 4. ACTUALIZARE VIZUALĂ
        MyTree.Refresh()
    End Sub

    ' =============================================================
    ' 2. LOGICA ȘTERGERE NOD
    ' =============================================================
    Private Sub ExecuteRemoveNode(nodeId As String)
        Dim removed As Boolean = False

        ' Încercăm să ștergem din rădăcini
        removed = RemoveNodeFromList(MyTree.Items, nodeId)

        ' Actualizare vizuală
        If removed Then
            MyTree.Refresh()
        Else
            If DEBUG_MODE Then MsgBox("Node ID to remove not found: " & nodeId)
        End If
    End Sub

    ' Helper recursiv pentru ștergere
    Private Function RemoveNodeFromList(list As List(Of AdvancedTreeControl.TreeItem), targetId As String) As Boolean
        ' Iterăm invers pentru a putea șterge sigur
        For i As Integer = list.Count - 1 To 0 Step -1
            Dim it = list(i)

            ' Verificăm ID-ul (Tag)
            Dim currentId As String = If(it.Tag IsNot Nothing, it.Tag.ToString(), "")

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

    ' Helper recursiv pentru găsire după ID (Tag)
    Private Function FindNodeByIdRecursive(item As AdvancedTreeControl.TreeItem, targetId As String) As AdvancedTreeControl.TreeItem
        Dim currentId As String = If(item.Tag IsNot Nothing, item.Tag.ToString(), "")
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
            If DEBUG_MODE Then MsgBox("Nodul nu a fost găsit: " & text)
        End If
    End Sub

    Private Function SearchNodeRecursive(item As AdvancedTreeControl.TreeItem, text As String, exact As Boolean) As AdvancedTreeControl.TreeItem
        ' Verificare curentă
        Dim isMatch As Boolean = False
        If exact Then
            isMatch = (String.Compare(item.Text, text, True) = 0)
        Else
            isMatch = (item.Text.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
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
            If Not _cleaningDone Then
                CurataResurseSiIesi()
            End If
        End If

        MyBase.WndProc(m)
    End Sub
End Class

