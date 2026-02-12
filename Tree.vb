Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices

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

            ' Inițializare Timer monitorizare
            _MonitorTimer = New Timer With {.Interval = 100, .Enabled = False}

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
                _formHwnd = New IntPtr(2298604) '################
                '################################################
                _mainAccessHwnd = New IntPtr(3213882)
                _idTree = "frmFX_MAIN" '"Clasificatii" '"frmFX_MAIN"
                _fisier = "C:\Avacont\Res\tree_frmFX_MAIN.xml" 'tree_Clasificatii.xml" 'tree_frmFX_MAIN.xml"
            End If
#Else
            If _formHwnd = IntPtr.Zero Or _mainAccessHwnd = IntPtr.Zero Then
                MessageBox.Show("EROARE: Parametrii de lansare invalizi!", $"Tree_Load ({version})", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If
#End If
            ' Conectare COM
            If Not IsWindow(_mainAccessHwnd) Then
                MessageBox.Show("EROARE: Fereastra Access invalida in DEBUG MODE!", $"Tree_Load ({version})", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If

            ConecteazaLaAccess(_mainAccessHwnd)

            ' === GĂSIRE FORMULAR PĂRINTE ACCESS (NOU) ===
            Debug.WriteLine("Căutare formular părinte Access:")
            _formParentHwnd = GetAccessFormParent(_formHwnd)

            If _formParentHwnd = IntPtr.Zero Then
                Debug.WriteLine("AVERTISMENT: Nu s-a găsit formular părinte!")
            Else
                Debug.WriteLine($"Formular părinte găsit: {GetWindowInfo(_formParentHwnd)}")
            End If
            ' ============================================

            Dim spHwnd As IntPtr = SetParent(Me.Handle, _formHwnd)
            'SetParent returneaza HWND-ul anterior al ferestrei copil daca reuseste, sau NULL daca esueaza
            If spHwnd = IntPtr.Zero Then
                Marshal.GetLastWin32Error()
                Dim dllErrInt As Integer = Marshal.GetLastWin32Error()
                Dim dllErr As String = New Win32Exception(dllErrInt).Message
                MessageBox.Show("EROARE: Formularul ACCESS nu este valid!" & ControlChars.CrLf & dllErr & ControlChars.CrLf & $"Form Handle:{_formHwnd}", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End If

            PositioneazaInParent()

            If Not String.IsNullOrEmpty(_fisier) Then
                If LoadXmlData(_fisier) Then
#If DEBUG Then
#Else
                    File.Delete(_fisier)
#End If
                End If
            Else
                MessageBox.Show("ERROR: Nu s-a putut încărca structura arborelui din Access.", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(0)
            End If

            TrimiteMesajAccess("HWND", Nothing, CStr(Me.Handle))

            ' === PORNIRE MONITORIZARE RESIZE ===
            Dim rParent As RECT
            GetClientRect(_formHwnd, rParent)
            _lastParentSize = New Size(rParent.Right - rParent.Left, rParent.Bottom - rParent.Top)
            _MonitorTimer.Start()

            ' _accessApp?.Run("OnTreeEvent", _idTree, "HWND", 0, "x", CStr(Me.Handle))
        Catch ex As Exception
            MessageBox.Show($"ERROR: {ex.Message}", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
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


    Private Sub MonitorTimer_Tick(sender As Object, e As EventArgs) Handles _MonitorTimer.Tick
        If _formHwnd = IntPtr.Zero Then Return

        ' Verificare validitate fereastră Access
        If Not IsWindow(_formHwnd) Then
            _MonitorTimer.Stop()
            CurataResurseSiIesi()
            Application.Exit()
            Return
        End If

        ' Citim dimensiunea curentă a părintelui Access
        Dim rParent As RECT
        GetClientRect(_formHwnd, rParent)
        Dim currentSize As New Size(rParent.Right - rParent.Left, rParent.Bottom - rParent.Top)

        ' Doar dacă s-a schimbat dimensiunea, redimensionăm copilul
        If currentSize <> _lastParentSize Then
            _lastParentSize = currentSize
            PositioneazaInParent()
        End If

        If MyTree.IsPopupTree Then
            ' === VERIFICARE FOCUS ===
            Dim foregroundWnd As IntPtr = GetForegroundWindow()

            If foregroundWnd <> _formHwnd AndAlso foregroundWnd <> _formParentHwnd Then
                Debug.WriteLine($">>> Focus pierdut: {GetWindowInfo(foregroundWnd)}")

                ' OPREȘTE timer-ul ÎNAINTE de SendMessage
                _MonitorTimer.Stop()

                ' Trimite WM_CLOSE și așteaptă răspunsul (blocking)
                SendMessage(_formParentHwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)

                ' Verificăm ce s-a întâmplat
                If Not IsWindow(_formParentHwnd) Then
                    ' Access a acceptat închiderea (Yes/No/fără modificări)
                    Debug.WriteLine(">>> Access a închis formularul")
                    Application.Exit()
                    Return
                Else
                    ' Access a refuzat (Cancel)
                    Debug.WriteLine(">>> Access a anulat închiderea, repornesc timer")
                    SetFocus(_formHwnd)
                    _MonitorTimer.Start()  ' REPORNEȘTE timer-ul
                End If
            End If
        End If

    End Sub

    Private Function GetAccessFormParent(childHwnd As IntPtr) As IntPtr
        ' Urcă ierarhia până găsește fereastra de tip formular Access (top-level sau popup)
        Dim currentHwnd As IntPtr = GetParent(childHwnd)
        Dim maxLevels As Integer = 10 ' Protecție împotriva loop infinit
        Dim level As Integer = 0

        While currentHwnd <> IntPtr.Zero AndAlso level < maxLevels
            ' Debug - vezi ce ferestre găsești
            Debug.WriteLine($"  Nivel {level}: {GetWindowInfo(currentHwnd)}")

            ' Verificăm dacă e fereastră top-level (are WS_CAPTION sau WS_POPUP)
            Dim style As Integer = GetWindowLong(currentHwnd, GWL_STYLE)
            Dim isTopLevel As Boolean = (style And WS_CAPTION) <> 0 OrElse (style And WS_POPUP) <> 0

            If isTopLevel Then
                Debug.WriteLine($"  → Găsit formular părinte la nivel {level}")
                Return currentHwnd
            End If

            currentHwnd = GetParent(currentHwnd)
            level += 1
        End While

        Debug.WriteLine("  → NU s-a găsit formular părinte, returnez IntPtr.Zero")
        Return IntPtr.Zero
    End Function
    Private Function GetWindowInfo(hWnd As IntPtr) As String
        If hWnd = IntPtr.Zero Then Return "NULL"

        ' Obține titlul ferestrei
        Dim length As Integer = GetWindowTextLength(hWnd)
        If length = 0 Then Return $"HWND:{hWnd:X} (fără titlu)"

        Dim sb As New System.Text.StringBuilder(length + 1)
        GetWindowText(hWnd, sb, sb.Capacity)

        ' Obține ProcessID
        Dim processId As Integer = 0
        GetWindowThreadProcessId(hWnd, processId)

        Return $"HWND:{hWnd:X} | PID:{processId} | Title:[{sb}]"
    End Function
End Class