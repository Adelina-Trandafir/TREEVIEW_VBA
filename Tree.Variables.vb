Partial Public Class Tree
    Private WithEvents MonitorTimer As System.Windows.Forms.Timer
    Private WithEvents MyTree As AdvancedTreeControl

    Private _formHwnd As IntPtr = IntPtr.Zero
    Private _detailHwnd As IntPtr = IntPtr.Zero
    Private _accessApp As Object = Nothing ' Aici stocăm referința la Access
    Private _mainAccessHwnd As IntPtr = IntPtr.Zero ' Handle-ul ferestrei principale Access
    Private _idTree As String = String.Empty
    Private _fisier As String = String.Empty

    ' Flag pentru a nu rula curatarea de doua ori
    Private _cleaningDone As Boolean = False

    ' Cache pentru imaginile decodate din Base64
    Private _imageCache As New Dictionary(Of String, Image)

    Private DEBUG_MODE As Boolean = True
    Private _manual_params As Boolean = False

    Private Shared ReadOnly inCommandSeparator As String() = New String() {"||"}

End Class
