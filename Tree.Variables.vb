''' <summary>Definitia unei coloane in modul TreeListView. Populata din &lt;Columns&gt; XML.</summary>
Friend Structure ColumnDef
    Dim Name As String
    Dim Header As String
    Dim Width As Integer
    Dim ColType As String       ' "Text" | "Number" | "Date" | "Boolean"
    Dim Align As HorizontalAlignment
    Dim Format As String
End Structure

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

    ' === HANDSHAKE VBA READY ===
    Private _vbaReady As Boolean = False
    Private _pendingMessages As New Queue(Of Action)
    Private _readyPollTimer As Timer = Nothing
    Private _handshakeStart As DateTime

    ' === NORMAL VBA COMMUNICATION ===
    Private _vbaBusy As Boolean = False
    Private _eventQueue As New Queue(Of Action)

    Private Const WM_APP_READY As Integer = &H8001  ' WM_APP + 1 (safe custom range)

    ' === POPUP GRACE ===
    Private _popupGraceActive As Boolean = False
    Private _popupGraceTimer As Timer = Nothing

    Private _pendingSelectedNodeId As String = String.Empty

    ' ── TreeListView ────────────────────────────────────────────────────────────
    Private _treeListView As Boolean = False
    Private _columns As New List(Of ColumnDef)

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
        ''' <summary>Dictionar coloana → date celula trimis prin ADD_BATCH_JSON. Nothing = nicio celula.</summary>
        Public Property Cells As Dictionary(Of String, CellDataDto) = Nothing

        Public Class CellDataDto
            Public Property Val As String = ""
            Public Property BackColor As String = ""   ' "#RRGGBB" sau "" = Color.Empty
            Public Property ForeColor As String = ""   ' "#RRGGBB" sau "" = Color.Empty
        End Class
    End Class
End Class
