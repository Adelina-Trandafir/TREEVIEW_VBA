Imports System.ComponentModel
Imports System.Runtime.InteropServices

Partial Public Class Tree
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetParent(ByVal hWndChild As IntPtr, ByVal hWndNewParent As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function MoveWindow(ByVal hWnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetClientRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetClassLongPtr(hWnd As IntPtr, nIndex As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetClassName(hWnd As IntPtr, lpClassName As System.Text.StringBuilder, nMaxCount As Integer) As Integer
    End Function


    <DllImport("gdi32.dll")>
    Private Shared Function GetObject(hBrush As IntPtr, nCount As Integer, lpObj As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function EnumChildWindows(hWndParent As IntPtr, lpEnumFunc As EnumChildProcDelegate, lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function PrintWindow(hWnd As IntPtr, hdcBlt As IntPtr, nFlags As UInteger) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function IsWindow(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("oleacc.dll")>
    Private Shared Function AccessibleObjectFromWindow(ByVal hwnd As IntPtr, ByVal dwId As UInteger, ByRef riid As Guid, <MarshalAs(UnmanagedType.IDispatch)> ByRef ppvObject As Object) As Integer
    End Function

    Private Delegate Function EnumChildProcDelegate(hWnd As IntPtr, lParam As IntPtr) As Boolean

    <StructLayout(LayoutKind.Sequential)>
    Private Structure LOGBRUSH
        Public lbStyle As Integer
        Public lbColor As UInteger
        Public lbHatch As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public Left, Top, Right, Bottom As Integer
    End Structure

    Private Const GCL_HBRBACKGROUND As Integer = -10
    Private Const OBJID_NATIVEOM As UInteger = &HFFFFFFF0&
    Private Const acSubform As Integer = 112

    Private WithEvents MonitorTimer As System.Windows.Forms.Timer
    Private WithEvents MyTree As AdvancedTreeControl

    Private _formHwnd As IntPtr = IntPtr.Zero
    Private _detailHwnd As IntPtr = IntPtr.Zero
    Private _accessApp As Object = Nothing ' Aici stocăm referința la Access
    Private _mainAccessHwnd As IntPtr = IntPtr.Zero ' Handle-ul ferestrei principale Access
    Private _idTree As String = String.Empty
    Private _fisier As String = String.Empty
    Private _RightClickFunc As String = String.Empty
    ' --- Timer pentru Click/DoubleClick ---

    Private Sub AliniazaLaParinte()
        If _formHwnd = IntPtr.Zero Then Return

        Dim rParent As RECT
        GetClientRect(_formHwnd, rParent)

        ' Verificăm dacă dimensiunea s-a schimbat față de dimensiunea noastră curentă
        ' pentru a evita flickering-ul (redesenarea inutilă)
        If Me.Width <> (rParent.Right - rParent.Left) OrElse Me.Height <> (rParent.Bottom - rParent.Top) Then
            MoveWindow(Me.Handle, 0, 0, rParent.Right - rParent.Left, rParent.Bottom - rParent.Top, True)
        End If
    End Sub

    ' --- Functii pentru detectie copil OFormSub ---
    Private Function FindOFormSub(parentHwnd As IntPtr) As IntPtr
        Dim result As IntPtr = IntPtr.Zero
        Dim callback As EnumChildProcDelegate = Function(hWnd, lParam)
                                                    Dim sb As New System.Text.StringBuilder(256)
                                                    Dim v = GetClassName(hWnd, sb, sb.Capacity)
                                                    If sb.ToString() = "OFormSub" Then
                                                        ' Verificăm dimensiunea
                                                        Dim r As RECT
                                                        If GetClientRect(hWnd, r) Then
                                                            Dim height = r.Bottom - r.Top
                                                            If height >= 2 Then ' prag minim 2 pixeli
                                                                result = hWnd
                                                                Return False ' am găsit unul valid, oprim enumerarea
                                                            End If
                                                        End If
                                                    End If
                                                    Return True ' continuăm enumerarea
                                                End Function

        EnumChildWindows(parentHwnd, callback, IntPtr.Zero)
        Return result
    End Function

    Private Function GetChildWindowBackgroundColor(hWnd As IntPtr) As Color
        Dim r As RECT
        GetClientRect(hWnd, r)
        Dim width = r.Right - r.Left
        Dim height = r.Bottom - r.Top
        Dim bmp As New Bitmap(width, height)
        Using g As Graphics = Graphics.FromImage(bmp)
            Dim hdc = g.GetHdc()
            PrintWindow(hWnd, hdc, 0)
            g.ReleaseHdc(hdc)
        End Using

        ' returneaza pixelul din stanga-sus
        Return bmp.GetPixel(0, 0)
    End Function

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

    Private Function GetFormObjectFromHwnd(hwndCautat As IntPtr) As Object
        If _accessApp Is Nothing Then Return Nothing

        Try
            Dim forms As Object = _accessApp.Forms

            ' 1. Căutăm în formularele principale (Top Level)
            For Each frm As Object In forms
                ' Verificăm dacă acesta este formularul căutat
                If frm.Hwnd = hwndCautat.ToInt32() Then
                    Return frm
                End If

                ' 2. Dacă nu e el, săpăm după subformulare în interiorul lui
                Dim foundSubForm As Object = FindSubFormRecursive(frm, hwndCautat)
                If foundSubForm IsNot Nothing Then
                    Return foundSubForm
                End If
            Next
        Catch ex As Exception
            MsgBox("Eroare la căutarea form-ului: " & ex.Message & vbCrLf)
        End Try

        Return Nothing
    End Function

    Private Function FindSubFormRecursive(parentForm As Object, hwndCautat As IntPtr) As Object
        Try
            Dim controls As Object = parentForm.Controls

            For Each ctl As Object In controls
                ' Verificăm proprietatea ControlType (112 = acSubform)
                ' Folosim Late Binding, deci trebuie să fim atenți la erori
                Dim ctlType As Integer = -1
                Try
                    ctlType = ctl.ControlType
                Catch
                    Continue For
                End Try

                If ctlType = acSubform Then
                    ' Am găsit un container de subform.
                    ' Accesăm proprietatea .Form a controlului (formularul din interior)
                    Dim childForm As Object = Nothing
                    Try
                        childForm = ctl.Form
                    Catch
                        ' E posibil ca subformul să fie gol (SourceObject lipsă)
                        Continue For
                    End Try

                    If childForm IsNot Nothing Then
                        ' Verificăm dacă acesta este cel căutat
                        If childForm.Hwnd = hwndCautat.ToInt32() Then
                            Return childForm
                        End If

                        ' Recursivitate: Căutăm și mai adânc (sub-subform)
                        Dim deepSearch As Object = FindSubFormRecursive(childForm, hwndCautat)
                        If deepSearch IsNot Nothing Then
                            Return deepSearch
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            ' Ignorăm erorile punctuale de acces la controale
        End Try

        Return Nothing
    End Function

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
        ' 1. Oprim timerul ca să nu se mai declanșeze o dată
        If MonitorTimer IsNot Nothing Then
            MonitorTimer.Stop()
        End If

        ' 2. Eliberăm obiectul Access Application
        If _accessApp IsNot Nothing Then
            Try
                ' Scade contorul de referințe COM. 
                ' Repetăm bucla până scade la 0 pentru siguranță maximă.
                While Marshal.ReleaseComObject(_accessApp) > 0
                End While
            Catch ex As Exception
                ' Ignorăm erorile aici, posibil ca Access să fie deja mort
            Finally
                _accessApp = Nothing
            End Try
        End If

        ' 3. TRUCUL ESENȚIAL: "Double GC Collect"
        ' Asta forțează .NET să curețe toate obiectele COM implicite (de ex: forms, controls atinse)
        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        ' 4. Închidem aplicația noastră
        Application.Exit()

        ' 5. (Opțional) Opțiunea Nucleară dacă Application.Exit() nu omoară procesul imediat
        ' Environment.Exit(0) 
    End Sub
End Class
