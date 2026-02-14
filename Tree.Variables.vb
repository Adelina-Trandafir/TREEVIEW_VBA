Partial Public Class Tree
    Private WithEvents _MonitorTimer As System.Windows.Forms.Timer

    Private WithEvents MyTree As AdvancedTreeControl

    Private _formHwnd As IntPtr = IntPtr.Zero
    Private _formParentHwnd As IntPtr = IntPtr.Zero
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

    ' Ultima dimensiune cunoscută a părintelui
    Private _lastParentSize As Size = Size.Empty

    Private _closeRequestSent As Boolean = False

    Private Shared ReadOnly inCommandSeparator As String() = New String() {"||"}

    Public Class NodeDto
        Public Property Key As String
        Public Property Caption As String
        Public Property IconClosed As String
        Public Property IconOpen As String
        Public Property IconRight As String = ""
        Public Property Expanded As Object
        Public Property Tag As String
        Public Property Children As List(Of NodeDto)
        Public Property LazyNode As Object
        Public Property Bold As Object
        Public Property Italic As Object
        Public Property ForeColor As String = ""
        Public Property BackColor As String = ""
    End Class
End Class
