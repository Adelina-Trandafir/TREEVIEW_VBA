Partial Public Class AdvancedTreeControl
    Public Enum TreeCheckState
        Unchecked = 0       ' Nebifat
        Checked = 1         ' Bifat complet
        Indeterminate = 2   ' Parțial bifat (pătrățel plin sau liniuță)
    End Enum

    Public treeID As String

    ' Nodes
    Public ReadOnly Items As New List(Of TreeItem)

    ' Culori
    Public LineColor As Color = Color.FromArgb(160, 160, 160)
    Public HoverBackColor As Color = Color.FromArgb(230, 240, 255)
    Public SelectedBackColor As Color = Color.FromArgb(200, 220, 255)
    Public SelectedBorderColor As Color = Color.FromArgb(150, 180, 255)
    Public RaiseLeftClickOnRightClick As Boolean = True
    Public ReRaiseClickOnSameNode As Boolean = True

    ' Tooltip
    Public TooltipDelayMs As Integer = 1000

    Private m_TreeFont = New Font("Consolas", 9)
    Public Property TreeFont As Font
        Get
            Return m_TreeFont
        End Get
        Set(value As Font)
            m_TreeFont = value
            m_FontName = m_TreeFont.Name ' Actualizează numele fontului pentru a reflecta schimbarea
            m_FontSize = m_TreeFont.Size ' Actualizează dimensiunea fontului pentru a reflecta schimbarea
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă fontul
        End Set
    End Property

    Private m_FontName As String = "Consolas"
    Public Property FontName As String
        Get
            Return m_FontName
        End Get
        Set(value As String)
            m_FontName = value
            m_TreeFont = New Font(m_FontName, TreeFont.Size) ' Actualizează fontul cu noul nume
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă fontul
        End Set
    End Property

    Private m_FontSize As Single = 9
    Public Property FontSize As Single
        Get
            Return m_FontSize
        End Get
        Set(value As Single)
            m_FontSize = value
            m_TreeFont = New Font(TreeFont.Name, m_FontSize) ' Actualizează fontul cu noua dimensiune
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă fontul
        End Set
    End Property

    Private m_ExpanderSize As Integer = 12
    Public Property ExpanderSize As Integer
        Get
            Return m_ExpanderSize
        End Get
        Set(value As Integer)
            m_ExpanderSize = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property

    Private m_Indent As Integer = 20
    Public Property Indent As Integer
        Get
            Return m_Indent
        End Get
        Set(value As Integer)
            m_Indent = value
            Me.Invalidate() ' Redesenează imediat controlul când se schimbă setarea
        End Set
    End Property

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

    Private _radioButtonLevel As Integer = -1  ' -1 = dezactivat
    Public Property RadioButtonLevel As Integer
        Get
            Return _radioButtonLevel
        End Get
        Set(value As Integer)
            _radioButtonLevel = value
            Me.Invalidate()
        End Set
    End Property

    Private m_BorderColor As Color = Color.Transparent
    Public Property BorderColor As Color
        Get
            Return m_BorderColor
        End Get
        Set(value As Color)
            m_BorderColor = value
            Me.Invalidate()
        End Set
    End Property
End Class
