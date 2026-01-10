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
        If _MonitorTimer IsNot Nothing Then
            _MonitorTimer.Stop()
        End If

        ' 2. Eliberare Access COM (foarte important cu Try/Catch)
        If _accessApp IsNot Nothing Then
            Try
                ' Eliberam referinta COM
                Marshal.ReleaseComObject(_accessApp)
            Catch ex As Exception
                ' Aceasta eroare e normala daca Access s-a inchis deja (RPC unavailable)
                'Log("COM Cleanup Info (Access probabil inchis deja): " & ex.Message)
            End Try
            _accessApp = Nothing
        End If
    End Sub

    Private Sub TrimiteMesajAccess(pItem As AdvancedTreeControl.TreeItem)
        If _accessApp IsNot Nothing Then
            Try
                Dim nodeId As String = If(pItem.Tag IsNot Nothing, pItem.Tag.ToString(), "")
                _accessApp.Run("OnTreeEvent", _idTree, nodeId, pItem.Text)
            Catch ex As Exception
                MsgBox("EROARE: " & ex.Message, vbOKOnly + vbCritical, "TrimiteMesajAccess")
            End Try
        End If
    End Sub

    ' =============================================================
    ' WNDPROC - INTERCEPTARE DISTRUGERE FORTATA
    ' =============================================================
    Protected Overrides Sub WndProc(ByRef m As Message)
        ' Interceptam momentul cand Windows vrea sa distruga fereastra
        If m.Msg = WM_DESTROY Then
            If Not _cleaningDone Then
                CurataResurseSiIesi()
            End If
        End If

        MyBase.WndProc(m)
    End Sub
End Class
