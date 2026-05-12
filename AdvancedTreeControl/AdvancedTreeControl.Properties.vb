Partial Public Class AdvancedTreeControl
    Public Enum TreeCheckState
        Unchecked = 0       ' Nebifat
        Checked = 1         ' Bifat complet
        Indeterminate = 2   ' Parțial bifat (pătrățel plin sau liniuță)
    End Enum

    Public treeID As String

    ' Nodes
    Public ReadOnly Items As New List(Of TreeItem)

    Public RaiseLeftClickOnRightClick As Boolean = True
    Public ReRaiseClickOnSameNode As Boolean = True

    ' Tooltip
    Public AutoHideTooltipMs As Integer = 5000

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

    Private _RootExpander As Boolean = True
    Public Property RootExpander As Boolean
        Get
            Return _RootExpander
        End Get
        Set(value As Boolean)
            _RootExpander = value
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

    Private _popupGraceMs As Integer = 1500
    Public Property PopupGraceMs() As Integer
        Get
            Return _popupGraceMs
        End Get
        Set(value As Integer)
            _popupGraceMs = value
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

    Private m_HoverBackColor As Color = Color.FromArgb(230, 240, 255)
    Public Property HoverBackColor As Color
        Get
            Return m_HoverBackColor
        End Get
        Set(value As Color)
            m_HoverBackColor = value
            Me.Invalidate()
        End Set
    End Property

    Private m_SelectedBackColor As Color = Color.FromArgb(200, 220, 255)
    Public Property SelectedBackColor As Color
        Get
            Return m_SelectedBackColor
        End Get
        Set(value As Color)
            m_SelectedBackColor = value
            Me.Invalidate()
        End Set
    End Property

    Private m_SelectedBorderColor As Color = Color.FromArgb(150, 180, 255)
    Public Property SelectedBorderColor As Color
        Get
            Return m_SelectedBorderColor
        End Get
        Set(value As Color)
            m_SelectedBorderColor = value
            Me.Invalidate()
        End Set
    End Property

    Private m_LineColor As Color = Color.FromArgb(160, 160, 160)
    Public Property LineColor As Color
        Get
            Return m_LineColor
        End Get
        Set(value As Color)
            m_LineColor = value
            Me.Invalidate()
        End Set
    End Property

    Private _tooltipDelayMs As Integer = 600
    Public Property TooltipDelayMs As Integer
        Get
            Return _tooltipDelayMs
        End Get
        Set(value As Integer)
            _tooltipDelayMs = value
            pTooltipTimer.Interval = value
        End Set
    End Property

    Private m_leftTextWidth As Integer = 0
    Public Property LeftTextWidth As Integer
        Get
            Return m_leftTextWidth
        End Get
        Set(value As Integer)
            m_leftTextWidth = Math.Max(0, value)
            Me.Invalidate()
        End Set
    End Property

    ' Lățime fixă rezervată pentru textul drept din caption cu separator ~~~
    ' 0 = nelimitat (dinamic)
    Private m_rightTextWidth As Integer = 0
    Public Property RightTextWidth As Integer
        Get
            Return m_rightTextWidth
        End Get
        Set(value As Integer)
            m_rightTextWidth = Math.Max(0, value)
            Me.Invalidate()
        End Set
    End Property
    Public ReadOnly Property OldSelectedNode As TreeItem
        Get
            Return pOldSelectedItem
        End Get
    End Property

    ' Când True, iconița din dreapta este vizibilă DOAR la hover pe nodul respectiv.
    ' Spațiul din dreapta e rezervat întotdeauna (textul nu sare la hover).
    ' Per-nod: TreeItem.ShowRightIconOnHover suprascrie globalul DOAR pentru nodul respectiv.
    Private _showRightIconOnHover As Boolean = False
    Public Property ShowRightIconOnHover As Boolean
        Get
            Return _showRightIconOnHover
        End Get
        Set(value As Boolean)
            _showRightIconOnHover = value
            Me.Invalidate()
        End Set
    End Property

    ' ══════════════════════════════════════════════════
    ' HEADER PROPERTIES
    ' ══════════════════════════════════════════════════

    Private _headerVisible As Boolean = False
    Public Property HeaderVisible As Boolean
        Get
            Return _headerVisible
        End Get
        Set(value As Boolean)
            _headerVisible = value
            Me.Invalidate()
        End Set
    End Property

    Private _headerHeight As Integer = 32
    Public Property HeaderHeight As Integer
        Get
            Return _headerHeight
        End Get
        Set(value As Integer)
            _headerHeight = Math.Max(16, value)
            Me.Invalidate()
        End Set
    End Property

    Private _headerCaption As String = ""
    Public Property HeaderCaption As String
        Get
            Return _headerCaption
        End Get
        Set(value As String)
            _headerCaption = value
            Me.Invalidate()
        End Set
    End Property

    ' Resolved images — set directly or via ResolveHeaderIcons()
    Private _headerLeftIcon As Image = Nothing
    Private _headerRightIcon As Image = Nothing
    Private _headerSearchIcon As Image = Nothing

    Public Property HeaderLeftIcon As Image
        Get
            Return _headerLeftIcon
        End Get
        Set(value As Image)
            _headerLeftIcon = value : Me.Invalidate()
        End Set
    End Property
    Public Property HeaderRightIcon As Image
        Get
            Return _headerRightIcon
        End Get
        Set(value As Image)
            _headerRightIcon = value
            Me.Invalidate()
        End Set
    End Property
    Public Property HeaderSearchIcon As Image
        Get
            Return _headerSearchIcon
        End Get
        Set(value As Image)
            _headerSearchIcon = value
            Me.Invalidate()
        End Set
    End Property

    ' Icon keys — stored for resolution after image cache is loaded
    Private _headerLeftIconKey As String = ""
    Private _headerRightIconKey As String = ""
    Private _headerSearchIconKey As String = ""

    Public Property HeaderLeftIconKey As String
        Get
            Return _headerLeftIconKey
        End Get
        Set(value As String)
            _headerLeftIconKey = value
        End Set
    End Property
    Public Property HeaderRightIconKey As String
        Get
            Return _headerRightIconKey
        End Get
        Set(value As String)
            _headerRightIconKey = value
        End Set
    End Property
    Public Property HeaderSearchIconKey As String
        Get
            Return _headerSearchIconKey
        End Get
        Set(value As String)
            _headerSearchIconKey = value
        End Set
    End Property

    Private _headerIconSize As New Size(16, 16)
    Public Property HeaderIconSize As Size
        Get
            Return _headerIconSize
        End Get
        Set(value As Size)
            _headerIconSize = value : Me.Invalidate()
        End Set
    End Property

    Private _headerBackColor As Color = Color.FromArgb(240, 240, 245)
    Public Property HeaderBackColor As Color
        Get
            Return _headerBackColor
        End Get
        Set(value As Color)
            _headerBackColor = value : Me.Invalidate()
        End Set
    End Property

    Private _headerForeColor As Color = Color.FromArgb(50, 50, 60)
    Public Property HeaderForeColor As Color
        Get
            Return _headerForeColor
        End Get
        Set(value As Color)
            _headerForeColor = value
            Me.Invalidate() : End Set
    End Property

    ' ══════════════════════════════════════════════════
    ' SEARCH PROPERTIES
    ' ══════════════════════════════════════════════════

    Private _searchType As en_Tree_SearchType = en_Tree_SearchType.SearchType_Contains
    Public Property SearchType As en_Tree_SearchType
        Get
            Return _searchType
        End Get
        Set(value As en_Tree_SearchType)
            _searchType = value
        End Set
    End Property

    Private _searchIn As en_Tree_SearchIn = en_Tree_SearchIn.SearchIn_Caption
    Public Property SearchIn As en_Tree_SearchIn
        Get
            Return _searchIn
        End Get
        Set(value As en_Tree_SearchIn)
            _searchIn = value
        End Set
    End Property

    Private _searchMode As en_Tree_SearchMode = en_Tree_SearchMode.SearchMode_Tree
    Public Property SearchMode As en_Tree_SearchMode
        Get
            Return _searchMode
        End Get
        Set(value As en_Tree_SearchMode)
            _searchMode = value
        End Set
    End Property

    Private _searchDropdownHeight As Integer = 220
    Public Property SearchDropdownHeight As Integer
        Get
            Return _searchDropdownHeight
        End Get
        Set(value As Integer)
            _searchDropdownHeight = Math.Max(60, value)
            Me.Invalidate()
        End Set
    End Property
End Class
