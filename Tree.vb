Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

Partial Public Class Tree
    Private ReadOnly version As Single = CSng(Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major)
    ' =============================================================
    ' INIT
    ' =============================================================
    Public Sub New()
#If DEBUG Then
        DEBUG_MODE = True
#End If
        InitializeComponent()
        Try
            ' Configurare Formă Gazdă
            Me.FormBorderStyle = FormBorderStyle.None
            Me.ShowInTaskbar = False
            Me.TopLevel = False
            Me.DoubleBuffered = True

            ' ======================================================
            ' INIȚIALIZARE AdvancedTreeControl
            ' ======================================================
            MyTree = New AdvancedTreeControl With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White,
                .ItemHeight = 22,
                .Indent = 20,
                .HoverBackColor = Color.FromArgb(230, 240, 255),
                .SelectedBackColor = Color.FromArgb(200, 220, 255)
            }

            Me.Controls.Add(MyTree)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "NEW_TREE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Tree_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        CurataResurseSiIesi()
    End Sub

    Private Sub Tree_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Dim args As String() = Environment.GetCommandLineArgs()

            If args.Length <= 1 And Not DEBUG_MODE Then
                MessageBox.Show("EROARE: Aplicatia poate fi pornita DOAR din AVACONT (/frm:? /acc:? /idt:!? /log:!?)", $"TreeView v{version}", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If

            Dim earlyTreeId As String = "startup"
            Dim debugSwitch As String = Nothing   ' Nothing = logging dezactivat
            Dim earlyLogPath As String = ""

            For Each a As String In args
                Dim aLow As String = a.ToLower()
                If aLow.StartsWith("/idt:") Then
                    earlyTreeId = a.Substring(5)
                ElseIf aLow.StartsWith("/log:") Then
                    earlyLogPath = a.Substring(5)
                ElseIf aLow = "/d" OrElse aLow = "/d2" Then
                    debugSwitch = "D2"
                ElseIf aLow = "/d1" Then
                    debugSwitch = "D1"
                ElseIf aLow = "/d3" Then
                    debugSwitch = "D3"
                End If
            Next

            If debugSwitch IsNot Nothing Then
                Dim level As TreeLogger.LogLevel = TreeLogger.LogLevel.INFO  ' /D sau /D2
                If debugSwitch = "D1" Then level = TreeLogger.LogLevel.WARN
                If debugSwitch = "D3" Then level = TreeLogger.LogLevel.DEBUG_

                TreeLogger.Init(earlyTreeId, level, earlyLogPath)
                TreeLogger.Info($"=== Aplicația pornește (v{version}) [logging={debugSwitch}] ===", "Tree_Load")
                TreeLogger.Debug($"Args: {String.Join(" ", args)}", "Tree_Load")
            End If

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
            If _formHwnd = IntPtr.Zero Or _mainAccessHwnd = IntPtr.Zero Then
                _manual_params = True
                '################################################
                _formHwnd = New IntPtr(592672) '################
                '################################################
                _mainAccessHwnd = New IntPtr(591810)
                _idTree = "Clasificatii" '"EFACTURA_2025" '"Clasificatii" '"frmFX_MAIN"
                _fisier = "C:\AVACONT\RES\tree_Clasificatii.xml" 'tree_EFACTURA_2025.xml" 'tree_Clasificatii.xml" 'tree_frmFX_MAIN.xml"
            End If
#Else
            If _formHwnd = IntPtr.Zero Or _mainAccessHwnd = IntPtr.Zero Then
                MessageBox.Show("EROARE: Parametrii de lansare invalizi!", $"TreeView v{version}", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If
#End If

            If Not String.IsNullOrEmpty(_fisier) Then
                If LoadXmlData(_fisier) Then
#If DEBUG Then
#Else
                    File.Delete(_fisier)
#End If
                End If
            Else
                TreeLogger.Err("ERROR: Nu s-a putut încărca structura arborelui din Access.", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(0)
            End If

            ' Conectare COM
            If Not IsWindow(_mainAccessHwnd) Then
                TreeLogger.Err("EROARE: Fereastra Access invalida in DEBUG MODE!", $"Tree_Load ({version})", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If

            'TreeLogger.Debug(">>> Înainte de ConecteazaLaAccess", "PERF")
            ConecteazaLaAccess(_mainAccessHwnd)
            'TreeLogger.Debug(">>> După ConecteazaLaAccess", "PERF")

            ' === GĂSIRE FORMULAR PĂRINTE ACCESS (NOU) ===
            'TreeLogger.Debug(">>> Înainte de SetParent", "PERF")
            'Debug.WriteLine("Căutare formular părinte Access:")
            _formParentHwnd = GetAccessFormParent(_formHwnd)

            If _formParentHwnd = IntPtr.Zero Then
                TreeLogger.Debug("AVERTISMENT: Nu s-a găsit formular părinte!")
            Else
                TreeLogger.Info($"Formular părinte găsit: {GetWindowInfo(_formParentHwnd)}")
            End If
            ' ============================================

            Dim spHwnd As IntPtr = SetParent(Me.Handle, _formHwnd)

            'SetParent returneaza HWND-ul anterior al ferestrei copil daca reuseste, sau NULL daca esueaza
            If spHwnd = IntPtr.Zero Then
                Marshal.GetLastWin32Error()
                Dim dllErrInt As Integer = Marshal.GetLastWin32Error()
                Dim dllErr As String = New Win32Exception(dllErrInt).Message
                TreeLogger.Err("EROARE: Formularul ACCESS nu este valid!" & ControlChars.CrLf & dllErr & ControlChars.CrLf & $"Form Handle:{_formHwnd}", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If

            PositioneazaInParent()
            'TreeLogger.Debug(">>> După SetParent + Poziționare", "PERF")

            'TreeLogger.Debug(">>> Înainte de TrimiteMesajAccess HWND", "PERF")
            'TreeLogger.Debug(">>> După TrimiteMesajAccess HWND", "PERF")

            ' === PORNIRE MONITORIZARE RESIZE ===
            Dim rParent As RECT
            GetClientRect(_formHwnd, rParent)
            _lastParentSize = New Size(rParent.Right - rParent.Left, rParent.Bottom - rParent.Top)

            ' Mesajul se pune automat în coadă (VBA nu e ready încă)
            TrimiteMesajAccess("HWND", Nothing, CStr(Me.Handle))

            ' Inițializare Timer monitorizare
            MonitorTimer = New Timer With {.Interval = 100, .Enabled = False}
            MonitorTimer.Start()

#If DEBUG Then
            Dim prop As IntPtr = GetProp(_formHwnd, "VBA_READY_" & _idTree)
            OnVbaReady(prop)
#Else
            ' Poll timer: verifică dacă VBA a pus SetProp("VBA_READY") pe _formHwnd
            _readyPollTimer = New Timer With {.Interval = 30}
            AddHandler _readyPollTimer.Tick, Sub(s, ev)
                                                 Debug.Print(">>> Poll VBA_READY...")
                                                 ' 1. Mai există fereastra?
                                                 If Not IsWindow(_formHwnd) Then
                                                     If _formParentHwnd = IntPtr.Zero OrElse Not IsWindow(_formParentHwnd) Then
                                                         _readyPollTimer.Stop()
                                                         _readyPollTimer.Dispose()
                                                         _readyPollTimer = Nothing
                                                         TreeLogger.Warn("Fereastra dispărută în timpul handshake — ies", "ReadyPoll")
                                                         CurataResurseSiIesi()
                                                         Application.Exit()
                                                         Return
                                                     End If
                                                     TreeLogger.Warn(">>> Fereastra dispărută în timpul handshake — întreb VBA dacă a recreat ceva")
                                                     Return ' părintele există, poate Access recreează — așteptăm
                                                 End If

                                                 ' 2. Biletul e acolo?
                                                 Dim prop As IntPtr = GetProp(_formHwnd, "VBA_READY_" & _idTree)
                                                 'Debug.Print($">>> GetProp VBA_READY: {_formHwnd.ToInt64}:{ _idTree}")
                                                 If prop <> IntPtr.Zero Then
                                                     _readyPollTimer.Stop()
                                                     _readyPollTimer.Dispose()
                                                     _readyPollTimer = Nothing
                                                     TreeLogger.Info($">>> VBA a confirmat că e ready {_idTree}!", "ReadyPoll")
                                                     'RemoveProp(_formHwnd, "VBA_READY_" & _idTree)
                                                     OnVbaReady(prop)
                                                 End If
                                             End Sub
            _readyPollTimer.Start()

#End If

        Catch ex As Exception
            TreeLogger.Ex(ex, "Tree_Load")
        End Try
    End Sub

    ' =============================================================
    ' MOUSE EVENTS
    ' =============================================================
    Private Sub MyTree_NodeDoubleClicked(pNode As AdvancedTreeControl.TreeItem, e As MouseEventArgs) Handles MyTree.NodeDoubleClicked
        If Not MyTree.IsPopupTree Then TrimiteMesajAccess("DblClick", pNode)
    End Sub

    Private Sub MyTree_NodeMouseUp(pItem As AdvancedTreeControl.TreeItem, e As MouseEventArgs) Handles MyTree.NodeMouseUp
        If e.Button = MouseButtons.Left Then
            TrimiteMesajAccess("Click", pItem)
            ' ── TreeListView: ridica si CELLCLICK daca s-a dat click pe o coloana ──
            Try
                If _treeListView AndAlso pItem.LastClickedColumnIndex >= 0 Then
                    Dim extraInfo As String = $"{pItem.LastClickedColumnIndex}||{pItem.LastClickedColumnName}"
                    TrimiteMesajAccess("CellClick", pItem, extraInfo)
                End If
            Catch ex As Exception
                TreeLogger.Ex(ex, "MyTree_NodeMouseUp/CellClick")
            End Try
            ' ──────────────────────────────────────────────────────────────────────

        ElseIf e.Button = MouseButtons.Right Then
            If MyTree.RaiseLeftClickOnRightClick OrElse pItem IsNot MyTree.OldSelectedNode Then
                TrimiteMesajAccess("Click", pItem)
            End If

            If Not String.IsNullOrEmpty(MyTree.RightClickFunction) AndAlso _accessApp IsNot Nothing Then
                TrimiteMesajAccess("RightClickFunction", pItem)
            Else
                TrimiteMesajAccess("RightClick", pItem, String.Join(",", e.Location.X.ToString(), e.Location.Y.ToString()))
            End If
        End If
    End Sub

    Private Sub MyTree_NodeChecked(pNode As AdvancedTreeControl.TreeItem) Handles MyTree.NodeChecked
        TrimiteMesajAccess("NodeChecked", pNode, If(pNode.CheckState = CheckState.Checked, 1, 0))
    End Sub

    Private Sub MyTree_RightIconClicked(pNode As AdvancedTreeControl.TreeItem, e As MouseEventArgs) Handles MyTree.RightIconClicked
        TrimiteMesajAccess("RightIconClicked", pNode, String.Join(",", e.Location.X.ToString(), e.Location.Y.ToString()))
    End Sub

    Private Sub MyTree_NodeRadioSelected(nodeOn As AdvancedTreeControl.TreeItem, nodeOff As AdvancedTreeControl.TreeItem) Handles MyTree.NodeRadioSelected
        Dim nodeOffKey As String = If(nodeOff IsNot Nothing, nodeOff.Key, "")
        TrimiteMesajAccess("NodeRadio", nodeOn, nodeOffKey)
    End Sub

    Private Sub MyTree_SearchFinished(matchingItems As List(Of AdvancedTreeControl.TreeItem), searchText As String) Handles MyTree.SearchFinished
        Dim sb As New System.Text.StringBuilder()
        sb.Append("["c)
        For i As Integer = 0 To matchingItems.Count - 1
            Dim n = matchingItems(i)
            Dim parentKey As String = If(n.Parent IsNot Nothing, n.Parent.Key, "")
            Dim rootKey As String = GetNodeRootKey(n)

            sb.Append("{"c)
            sb.Append($"""Key"":""{EscapeJson(n.Key)}"",")
            sb.Append($"""ParentKey"":""{EscapeJson(parentKey)}"",")
            sb.Append($"""RootKey"":""{EscapeJson(rootKey)}""")
            sb.Append("}"c)
            If i < matchingItems.Count - 1 Then sb.Append(","c)
        Next
        sb.Append("]"c)

        TrimiteMesajAccess("SearchFinished", Nothing, searchText & "||" & sb.ToString())
    End Sub

    Private Sub MyTree_HeaderRightIconClicked(e As MouseEventArgs) Handles MyTree.HeaderRightIconClicked
        TrimiteMesajAccess("HeaderRightIconClicked", Nothing)
    End Sub

    ' =============================================================
    ' TIMER MONITORIZARE RESIZE & FOCUS
    ' =============================================================
    Private Sub MonitorTimer_Tick(sender As Object, e As EventArgs) Handles MonitorTimer.Tick
        If _formHwnd = IntPtr.Zero Then Return

        ' === VERIFICARE VALIDITATE _formHwnd ===
        If Not IsWindow(_formHwnd) Then
            ' 1. Părintele mai există?
            If _formParentHwnd = IntPtr.Zero OrElse Not IsWindow(_formParentHwnd) Then
                TreeLogger.Warn("Nici _formParentHwnd nu mai e valid, închid aplicația", "MonitorTimer_Tick")
                MonitorTimer.Stop()
                CurataResurseSiIesi()
                Application.Exit()
                Return
            End If

            ' 2. Părintele există — întrebăm VBA pentru HWND nou
            TreeLogger.Warn("_formHwnd invalidat, întreb VBA...", "MonitorTimer_Tick")
            Dim newHwnd As IntPtr = IntPtr.Zero
            Try
                Dim result As Object = _accessApp.Run("GetTreeFormHwnd", _idTree)
                If result IsNot Nothing Then
                    newHwnd = New IntPtr(CLng(result))
                End If
            Catch ex As Exception
                TreeLogger.Ex(ex, "MonitorTimer_Tick.Recovery")
            End Try

            ' 3. Valid?
            If newHwnd <> IntPtr.Zero AndAlso IsWindow(newHwnd) Then
                ReattachToNewHwnd(newHwnd)
            Else
                TreeLogger.Info("VBA a confirmat: formularul nu mai există", "MonitorTimer_Tick")
                MonitorTimer.Stop()
                CurataResurseSiIesi()
                Application.Exit()
            End If
            Return
        End If

        ' === RESIZE MONITORING ===
        Dim rParent As RECT
        GetClientRect(_formHwnd, rParent)
        Dim currentSize As New Size(rParent.Right - rParent.Left, rParent.Bottom - rParent.Top)

        If currentSize <> _lastParentSize Then
            _lastParentSize = currentSize
            PositioneazaInParent()
        End If

        ' === POPUP FOCUS MONITORING ===
        If MyTree.IsPopupTree Then
            If _popupGraceActive Then Return

            Dim foregroundWnd As IntPtr = GetForegroundWindow()
            Dim tooltipHwnd As IntPtr = MyTree.TooltipPopupHandle

            ' === FIX: dacă foreground aparține propriului proces, NU e focus pierdut ===
            Dim fgPid As Integer = 0
            Dim v1 = GetWindowThreadProcessId(foregroundWnd, fgPid)

            If fgPid = Environment.ProcessId Then Return
            ' =========================================================================
            If foregroundWnd <> _formHwnd AndAlso foregroundWnd <> _formParentHwnd AndAlso foregroundWnd <> tooltipHwnd Then
                TreeLogger.Debug(GetWindowDump(foregroundWnd, "FG-LOST"), "MonitorTimer_Tick")
                _MonitorTimer.Stop()
                PostMessage(_formParentHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
            End If
        End If
    End Sub

    Private Function GetWindowDump(hWnd As IntPtr, Optional label As String = "") As String
        If hWnd = IntPtr.Zero Then Return $"{label}HWND: NULL"

        Dim sb As New System.Text.StringBuilder()
        Dim prefix = If(label <> "", $"[{label}] ", "")

        ' --- titlu + clasă ---
        Dim titleLen = GetWindowTextLength(hWnd)
        Dim titleSb As New System.Text.StringBuilder(titleLen + 1)
        Dim v = GetWindowText(hWnd, titleSb, titleSb.Capacity)
        Dim title = If(titleLen = 0, "(fără titlu)", titleSb.ToString())

        Dim classSb As New System.Text.StringBuilder(256)
        Dim v4 = GetClassName(hWnd, classSb, 256)

        ' --- PID + nume proces ---
        Dim pid As Integer = 0
        Dim v1 = GetWindowThreadProcessId(hWnd, pid)
        Dim procName As String = "(unknown)"
        Try
            procName = System.Diagnostics.Process.GetProcessById(pid).ProcessName
        Catch : End Try

        sb.AppendLine($"{prefix}HWND:{hWnd:X}  Title:[{title}]  Class:[{classSb}]  PID:{pid} ({procName})")

        ' --- parent + owner ---
        Dim parentHwnd = GetParent(hWnd)
        Dim ownerHwnd = GetWindow(hWnd, GW_OWNER)

        If parentHwnd <> IntPtr.Zero Then
            Dim pClass As New System.Text.StringBuilder(256)
            Dim v5 = GetClassName(parentHwnd, pClass, 256)
            Dim pTitleLen = GetWindowTextLength(parentHwnd)
            Dim pTitle As New System.Text.StringBuilder(pTitleLen + 1)
            Dim v2 = GetWindowText(parentHwnd, pTitle, pTitle.Capacity)
            sb.AppendLine($"  Parent : HWND:{parentHwnd:X}  Class:[{pClass}]  Title:[{If(pTitleLen = 0, "(fără titlu)", pTitle.ToString())}]")
        Else
            sb.AppendLine($"  Parent : (none)")
        End If

        If ownerHwnd <> IntPtr.Zero Then
            Dim oClass As New System.Text.StringBuilder(256)
            Dim v6 = GetClassName(ownerHwnd, oClass, 256)
            sb.AppendLine($"  Owner  : HWND:{ownerHwnd:X}  Class:[{oClass}]")
        End If

        ' --- root ancestor (top-level) ---
        Dim rootHwnd = GetAncestor(hWnd, GA_ROOT)
        If rootHwnd <> hWnd Then
            Dim rClass As New System.Text.StringBuilder(256)
            Dim v7 = GetClassName(rootHwnd, rClass, 256)
            Dim rTitleLen = GetWindowTextLength(rootHwnd)
            Dim rTitle As New System.Text.StringBuilder(rTitleLen + 1)
            Dim v3 = GetWindowText(rootHwnd, rTitle, rTitle.Capacity)
            sb.AppendLine($"  Root   : HWND:{rootHwnd:X}  Class:[{rClass}]  Title:[{If(rTitleLen = 0, "(fără titlu)", rTitle.ToString())}]")
        End If

        ' --- relație cu HWND-urile cunoscute ---
        sb.AppendLine($"  == față de known HWNDs ==")
        sb.AppendLine($"  _formHwnd       ({_formHwnd:X}) : {If(hWnd = _formHwnd, "MATCH", "diferit")}")
        sb.AppendLine($"  _formParentHwnd ({_formParentHwnd:X}) : {If(hWnd = _formParentHwnd, "MATCH", "diferit")}")
        sb.AppendLine($"  IsChild(_formParentHwnd, fg) : {IsChild(_formParentHwnd, hWnd)}")
        sb.AppendLine($"  IsChild(_formHwnd, fg)       : {IsChild(_formHwnd, hWnd)}")

        ' --- lanț complet de parents ---
        sb.Append("  ParentChain: ")
        Dim chain As New List(Of String)
        Dim cur = GetParent(hWnd)
        Dim depth = 0
        While cur <> IntPtr.Zero AndAlso depth < 15
            Dim cClass As New System.Text.StringBuilder(256)
            Dim v8 = GetClassName(cur, cClass, 256)
            chain.Add($"HWND:{cur:X}[{cClass}]")
            cur = GetParent(cur)
            depth += 1
        End While
        sb.AppendLine(If(chain.Count = 0, "(none)", String.Join(" → ", chain)))

        Return sb.ToString().TrimEnd()
    End Function

    Private Function GetAccessFormParent(childHwnd As IntPtr) As IntPtr
        Dim currentHwnd As IntPtr = GetParent(childHwnd)
        Dim maxLevels As Integer = 10
        Dim className As New StringBuilder(256)

        For level As Integer = 0 To maxLevels - 1
            If currentHwnd = IntPtr.Zero Then Exit For

            Dim v = GetClassName(currentHwnd, className, 256)

            Dim cls As String = className.ToString()
            TreeLogger.Info($"  Nivel {level}: class='{cls}' {GetWindowInfo(currentHwnd)}", "GetAccessFormParent")

            If cls = "OForm" OrElse cls = "OFormPopup" OrElse cls = "OFormPopupNC" OrElse cls = "OFormNoClose" Then
                TreeLogger.Info($"  → Găsit formular Access la nivel {level} ({cls})", "GetAccessFormParent")
                Return currentHwnd
            End If

            currentHwnd = GetParent(currentHwnd)
        Next

        TreeLogger.Err("  → NU s-a găsit OForm/OFormPopup, returnez IntPtr.Zero", "GetAccessFormParent")
        Return IntPtr.Zero
    End Function

    Private Function GetWindowInfo(hWnd As IntPtr) As String
        If hWnd = IntPtr.Zero Then Return "NULL"

        ' Obține titlul ferestrei
        Dim length As Integer = GetWindowTextLength(hWnd)
        If length = 0 Then Return $"HWND:{hWnd:X} (fără titlu)"

        Dim sb As New System.Text.StringBuilder(length + 1)
        Dim v = GetWindowText(hWnd, sb, sb.Capacity)

        ' Obține ProcessID
        Dim processId As Integer = 0
        Dim v2 = GetWindowThreadProcessId(hWnd, processId)

        Return $"HWND:{hWnd:X} | PID:{processId} | Title:[{sb}]"
    End Function

    Private Sub FlushPendingMessages()
        While _pendingMessages.Count > 0
            Dim act As Action = _pendingMessages.Peek()
            Try
                act.Invoke()
                _pendingMessages.Dequeue()
            Catch ex As Runtime.InteropServices.COMException
                ' VBA încă ocupată — reîncercăm silențios peste 50ms
                Dim retryTimer As New Timer With {.Interval = 50}
                AddHandler retryTimer.Tick, Sub(s, ev)
                                                retryTimer.Stop()
                                                retryTimer.Dispose()
                                                FlushPendingMessages()
                                            End Sub
                retryTimer.Start()
                Return
            Catch ex As Exception
                TreeLogger.Ex(ex, "FlushPending")
                _pendingMessages.Dequeue()
            End Try
        End While

        ' Totul trimis
        Dim elapsed As TimeSpan = DateTime.Now - _handshakeStart
        TreeLogger.Info($"Flush complet — {elapsed.TotalMilliseconds:F0}ms de la prima punere în coadă", "FlushPending")
    End Sub

    Private Sub RunOnVBA(action As String, nodeKey As String, nodeCaption As String, extraInfo As String)
        If _vbaBusy Then
            ' Capturam parametrii, nu referinta la variabile locale
            Dim a = action, k = nodeKey, c = nodeCaption, x = extraInfo
            _eventQueue.Enqueue(Sub() RunOnVBA(a, k, c, x))
            Return
        End If

        _vbaBusy = True
        Dim succeeded As Boolean = False

        Try
            _accessApp.Run("OnTreeEvent", _idTree, action, nodeKey, nodeCaption, extraInfo)
            succeeded = True
        Catch comEx As Runtime.InteropServices.COMException
            ' VBA ocupata neasteptat (race extern) — retry dupa 100ms
            Dim a = action, k = nodeKey, c = nodeCaption, x = extraInfo
            Dim retryTimer As New Timer With {.Interval = 100}
            AddHandler retryTimer.Tick, Sub(s, ev)
                                            retryTimer.Stop()
                                            retryTimer.Dispose()
                                            _vbaBusy = False
                                            RunOnVBA(a, k, c, x)
                                        End Sub
            retryTimer.Start()
            Return  ' _vbaBusy ramane True pana la retry
        Catch ex As Exception
            TreeLogger.Debug("RunOnVBA Err: " & ex.Message, "RunOnVBA")
            succeeded = True  ' eroare non-COM, eliberam oricum
        End Try

        If succeeded Then
            _vbaBusy = False
            If _eventQueue.Count > 0 Then
                _eventQueue.Dequeue().Invoke()
            End If
        End If
    End Sub

    Private Sub OnVbaReady(newFormHwnd As IntPtr)
        If _vbaReady Then Return
        _vbaReady = True

        _readyPollTimer?.Stop()
        _readyPollTimer?.Dispose()
        _readyPollTimer = Nothing

        If newFormHwnd <> IntPtr.Zero AndAlso newFormHwnd <> _formHwnd Then
            TreeLogger.Info($"formHwnd schimbat: {_formHwnd:X} → {newFormHwnd:X}", "OnVbaReady")
            ReattachToNewHwnd(newFormHwnd)
        End If

        TreeLogger.Info($"VBA READY — flush {_pendingMessages.Count} mesaje pending", "OnVbaReady")
        FlushPendingMessages()

        ' === POPUP: forțare focus + perioadă de grație ===
        If MyTree.IsPopupTree Then
            Dim focused = SetForegroundWindow(_formHwnd)
            TreeLogger.Info($"Popup ready — SetForegroundWindow({_formHwnd:X}) = {focused}, grație {MyTree.PopupGraceMs}ms", "OnVbaReady")

            _popupGraceActive = True
            _popupGraceTimer = New Timer With {.Interval = MyTree.PopupGraceMs}
            AddHandler _popupGraceTimer.Tick, Sub(s, ev)
                                                  _popupGraceTimer.Stop()
                                                  _popupGraceTimer.Dispose()
                                                  _popupGraceTimer = Nothing
                                                  _popupGraceActive = False
                                                  TreeLogger.Info($"Perioadă de grație expirată ({MyTree.PopupGraceMs}ms) — pornesc monitorizarea focus", "PopupGrace")
                                              End Sub
            _popupGraceTimer.Start()
        End If
    End Sub

    Private Function GetNodeRootKey(item As AdvancedTreeControl.TreeItem) As String
        Dim current As AdvancedTreeControl.TreeItem = item
        While current.Parent IsNot Nothing
            current = current.Parent
        End While
        Return current.Key
    End Function

    Private Function EscapeJson(s As String) As String
        Return s.Replace("\", "\\").Replace("""", "\""")
    End Function
End Class