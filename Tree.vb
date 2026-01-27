Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices

' Asumăm că AdvancedTreeControl este definit în proiect
' V.4 - 2024.06.17
' Adugat LAZY LOADING pentru noduri

Partial Public Class Tree
    ' =============================================================
    ' INIT
    ' =============================================================
    Public Sub New()
        ' 1. Deschidem consola imediat pentru a vedea ce se intampla
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
            _MonitorTimer = New Timer With {.Interval = 10, .Enabled = False}

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
                MessageBox.Show("EROARE: Aplicatia poate fi pornita DOAR din AVACONT (/frm:? /acc:? /idt:?!", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                _formHwnd = New IntPtr(6230292) '################
                '################################################
                _mainAccessHwnd = New IntPtr(8458146)
                _idTree = "frmFX_MAIN" '"Clasificatii" '"frmFX_MAIN"
                _fisier = "C:\Avacont\Res\tree_frmFX_MAIN.xml" 'tree_Clasificatii.xml" 'tree_frmFX_MAIN.xml"
            End If
#Else
            If _formHwnd = IntPtr.Zero Or _mainAccessHwnd = IntPtr.Zero Then
                messagebox.show("EROARE: Parametrii de lansare invalizi!",  "Tree_Load", MessageBoxButtons.OK,MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If
#End If
            ' Conectare COM
            If Not IsWindow(_mainAccessHwnd) Then
                MessageBox.Show("EROARE: Fereastra Access invalida in DEBUG MODE!", "Tree_Load", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Environment.Exit(-1)
            End If

            ConecteazaLaAccess(_mainAccessHwnd)
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
End Class