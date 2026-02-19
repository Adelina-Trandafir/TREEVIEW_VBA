Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

' V.5.0 - 28.01.2026
' Adugat LAZY LOADING pentru noduri
' Adaugat Eveniment RightIconClicked
' V.5.0 - 01.02.2026
' Modificat checkbox si adaugat HasNodeIcons 
' Activat timer monitorizare redimensionare
' V.6.0 - 10.02.2026
' Adaugat click in zona moarta a copacului => deschide / inchide nodul
' V.7.0 - 12.02.2026
' Adaugat functia de inchidere automata a popup-ului la click pe leaf

Partial Public Class Tree
    Private version As String = "7.0"
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
                MessageBox.Show("EROARE: Aplicatia poate fi pornita DOAR din AVACONT (/frm:? /acc:? /idt:?!", $"Tree_Load ({version})", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If

            ' Parsăm treeId din argumente ÎNAINTE de Init logger
            Dim earlyTreeId As String = "startup"
            For Each a As String In args
                If a.StartsWith("/idt:", StringComparison.CurrentCultureIgnoreCase) Then
                    earlyTreeId = a.Substring(5)
                    Exit For
                End If
            Next
            TreeLogger.Init(earlyTreeId)
            TreeLogger.Info($"=== Aplicația pornește (v{version}) ===", "Tree_Load")
            TreeLogger.Debug($"Args: {String.Join(" ", args)}", "Tree_Load")

            ' Diagnostic: background thread care loghează la fiecare 200ms
            'Dim diagThread As New System.Threading.Thread(Sub()
            '                                                  For i As Integer = 1 To 20  ' 6 secunde de monitorizare
            '                                                      TreeLogger.Debug($">>> HEARTBEAT #{i}", "BG_THREAD")
            '                                                      System.Threading.Thread.Sleep(200)
            '                                                  Next
            '                                              End Sub)
            'diagThread.IsBackground = True
            'diagThread.Start()



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
                _formHwnd = New IntPtr(3803798) '################
                '################################################
                _mainAccessHwnd = New IntPtr(395282)
                _idTree = "FX_MAIN_PLATI_RECEPTII" '"Clasificatii" '"frmFX_MAIN"
                _fisier = "C:\AVACONT\RES\tree_FX_MAIN_PLATI_RECEPTII.xml" 'tree_Clasificatii.xml" 'tree_frmFX_MAIN.xml"
            End If
#Else
            If _formHwnd = IntPtr.Zero Or _mainAccessHwnd = IntPtr.Zero Then
                MessageBox.Show("EROARE: Parametrii de lansare invalizi!", $"Tree_Load ({version})", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
            _MonitorTimer = New Timer With {.Interval = 100, .Enabled = False}
            _MonitorTimer.Start()

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
                                                     RemoveProp(_formHwnd, "VBA_READY_" & _idTree)
                                                     OnVbaReady(prop)
                                                 End If
                                             End Sub
            _readyPollTimer.Start()

        Catch ex As Exception
            TreeLogger.Ex(ex, "Tree_Load")
        End Try
    End Sub

    ' =============================================================
    ' MOUSE EVENTS
    ' =============================================================
    Private Sub MyTree_NodeMouseUp(pItem As AdvancedTreeControl.TreeItem, e As MouseEventArgs) Handles MyTree.NodeMouseUp
        If e.Button = MouseButtons.Left Then
            TrimiteMesajAccess("Click", pItem)
        End If

        ' --- POPUP: Închidere după click pe leaf ---
        If MyTree.IsPopupTree AndAlso pItem.Children.Count = 0 AndAlso Not pItem.LazyNode Then
            Me.BeginInvoke(Sub()
                               TreeLogger.Debug("Leaf click în popup - trimit WM_CLOSE", "MyTree_NodeMouseUp")
                               _MonitorTimer.Stop()
                               SendMessage(_formParentHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
                               If Not IsWindow(_formParentHwnd) Then
                                   Application.Exit()
                               End If
                           End Sub)
            Exit Sub
        End If

        If e.Button = MouseButtons.Right Then
            If Not String.IsNullOrEmpty(MyTree.RightClickFunction) AndAlso _accessApp IsNot Nothing Then
                TrimiteMesajAccess("RightClickFunction", pItem)
            End If

            TrimiteMesajAccess("RightClick", pItem, String.Join(",", e.Location.X.ToString(), e.Location.Y.ToString()))
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

    ' =============================================================
    ' TIMER MONITORIZARE RESIZE & FOCUS
    ' =============================================================
    Private Sub MonitorTimer_Tick(sender As Object, e As EventArgs) Handles _MonitorTimer.Tick
        If _formHwnd = IntPtr.Zero Then Return

        ' === VERIFICARE VALIDITATE _formHwnd ===
        If Not IsWindow(_formHwnd) Then
            ' 1. Părintele mai există?
            If _formParentHwnd = IntPtr.Zero OrElse Not IsWindow(_formParentHwnd) Then
                TreeLogger.Warn("Nici _formParentHwnd nu mai e valid, închid aplicația", "MonitorTimer_Tick")
                _MonitorTimer.Stop()
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
                _MonitorTimer.Stop()
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
            Dim foregroundWnd As IntPtr = GetForegroundWindow()

            If foregroundWnd <> _formHwnd AndAlso foregroundWnd <> _formParentHwnd Then
                TreeLogger.Debug($">>> Focus pierdut: {GetWindowInfo(foregroundWnd)}", "MonitorTimer_Tick")
                _MonitorTimer.Stop()
                SendMessage(_formParentHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)

                If Not IsWindow(_formParentHwnd) Then
                    TreeLogger.Info(">>> Access a închis formularul", "MonitorTimer_Tick")
                    Application.Exit()
                    Return
                Else
                    TreeLogger.Info(">>> Access a anulat închiderea, repornesc timer", "MonitorTimer_Tick")
                    SetFocus(_formHwnd)
                    _MonitorTimer.Start()
                End If
            End If
        End If
    End Sub

    Private Function GetAccessFormParent(childHwnd As IntPtr) As IntPtr
        Dim currentHwnd As IntPtr = GetParent(childHwnd)
        Dim maxLevels As Integer = 10
        Dim className As New StringBuilder(256)

        For level As Integer = 0 To maxLevels - 1
            If currentHwnd = IntPtr.Zero Then Exit For

            GetClassName(currentHwnd, className, 256)
            Dim cls As String = className.ToString()
            TreeLogger.Info($"  Nivel {level}: class='{cls}' {GetWindowInfo(currentHwnd)}", "GetAccessFormParent")

            If cls = "OForm" OrElse cls = "OFormPopup" Then
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
    End Sub
End Class