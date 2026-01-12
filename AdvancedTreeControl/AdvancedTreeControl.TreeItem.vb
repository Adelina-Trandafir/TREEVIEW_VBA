Partial Public Class AdvancedTreeControl
    Public Class TreeItem
        Public Key As String
        Public Caption As String
        Public Children As New List(Of TreeItem)
        Public Expanded As Boolean = True
        Public Parent As TreeItem
        Public Level As Integer
        Public CheckState As TreeCheckState = TreeCheckState.Unchecked
        Public LeftIconClosed As Image
        Public LeftIconOpen As Image
        Public RightIcon As Image

        Private _tag As Object

        ' Cache pentru lățimea textului (performanță la desenare)
        Friend TextWidth As Integer = -1

        ' Proprietate critică pentru desenarea corectă a liniilor verticale
        Public ReadOnly Property IsLastSibling As Boolean
            Get
                If Parent Is Nothing Then
                    ' Dacă e root, verificăm dacă e ultimul în lista principală a controlului
                    ' (Necesită referință la control, dar pentru simplitate desenăm standard)
                    Return True
                End If
                Return Parent.Children.LastOrDefault() Is Me
            End Get
        End Property

        Public Property Tag As Object
            Get
                Return _tag
            End Get
            Set(value As Object)
                _tag = value
            End Set
        End Property
    End Class
End Class
