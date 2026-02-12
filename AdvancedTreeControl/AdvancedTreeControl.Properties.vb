Partial Public Class AdvancedTreeControl
    Public Enum TreeCheckState
        Unchecked = 0       ' Nebifat
        Checked = 1         ' Bifat complet
        Indeterminate = 2   ' Parțial bifat (pătrățel plin sau liniuță)
    End Enum

    Public Indent As Integer = 20
    Public ExpanderSize As Integer = 12
    Public treeID As String

    ' Nodes
    Public ReadOnly Items As New List(Of TreeItem)

    ' Culori
    Public LineColor As Color = Color.FromArgb(160, 160, 160)
    Public HoverBackColor As Color = Color.FromArgb(230, 240, 255)
    Public SelectedBackColor As Color = Color.FromArgb(200, 220, 255)
    Public SelectedBorderColor As Color = Color.FromArgb(150, 180, 255)

    ' Tooltip
    Public TooltipDelayMs As Integer = 600

    Private _checkBoxSize As Integer = 16
    Public Property CheckBoxSize As Integer
        Get
            Return _checkBoxSize
        End Get
        Set(value As Integer)
            _checkBoxSize = value
            Me.Invalidate()
        End Set
    End Property

    ' Înălțimea rândului (calculată automat sau setată manual)
    Private _autoHeight As Boolean = False
    Private _itemHeight As Integer = 22
    Public Property ItemHeight As Integer
        Get
            Return _itemHeight
        End Get
        Set(value As Integer)
            _itemHeight = value
            '_autoHeight = False
            Me.Invalidate()
        End Set
    End Property

    ' Iconițe - Setarea lor declanșează recalcularea înălțimii rândului
    Private _leftIconSize As New Size(24, 24)
    Public Property LeftIconSize As Size
        Get
            Return _leftIconSize
        End Get
        Set(value As Size)
            _leftIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rightIconSize As New Size(14, 14)
    Public Property RightIconSize As Size
        Get
            Return _rightIconSize
        End Get
        Set(value As Size)
            _rightIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rootButton As Boolean = True
    Public Property RootButton As Boolean
        Get
            Return _rootButton
        End Get
        Set(value As Boolean)
            _rootButton = value
            Me.Invalidate()
        End Set
    End Property

    Private _rightClickFunc As String = ""
    Public Property RightClickFunction As String
        Get
            Return _rightClickFunc
        End Get
        Set(value As String)
            _rightClickFunc = value
            Me.Invalidate()
        End Set
    End Property

    Public Property SelectedNode As TreeItem
        Get
            Return pSelectedItem
        End Get
        Set(value As TreeItem)
            If pSelectedItem IsNot value Then
                pSelectedItem = value
                ' Invalidate to trigger redraw and show the new selection
                Me.Invalidate()
            End If
        End Set
    End Property

    Private _checkBoxes As Boolean = False
    Public Property CheckBoxes As Boolean
        Get
            Return _checkBoxes
        End Get
        Set(value As Boolean)
            _checkBoxes = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property

    Private _hasNodeIcons As Boolean = True
    Public Property HasNodeIcons As Boolean
        Get
            Return _hasNodeIcons
        End Get
        Set(value As Boolean)
            _hasNodeIcons = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property

    Private _isPopupTree As Boolean = False
    Public Property IsPopupTree As Boolean
        Get
            Return _isPopupTree
        End Get
        Set(value As Boolean)
            _isPopupTree = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property
End Class
