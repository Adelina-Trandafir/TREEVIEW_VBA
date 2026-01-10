Partial Public Class Tree
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
End Class
