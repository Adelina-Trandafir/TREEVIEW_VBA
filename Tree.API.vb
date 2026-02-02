Imports System.Runtime.InteropServices

Partial Public Class Tree
    ' =============================================================
    ' CONSOLA DEBUG - API
    ' =============================================================
    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As System.Text.StringBuilder, nMaxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetParent(hWnd As IntPtr) As IntPtr
    End Function

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

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As String) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SetFocus(hWnd As IntPtr) As IntPtr
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

    Private Const WM_SETTEXT As Integer = &HC
    Private Const GCL_HBRBACKGROUND As Integer = -10
    Private Const OBJID_NATIVEOM As UInteger = &HFFFFFFF0&
    Private Const acSubform As Integer = 112
    Private Const WM_DESTROY As Integer = &H2
    Private Const WM_CLOSE As Integer = &H10
    Private Const GWL_STYLE As Integer = -16
    Private Const WS_CAPTION As Integer = &HC00000
    Private Const WS_POPUP As Integer = &H80000000
End Class
