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

    Private m_Indent As Integer = 10
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
    Private _leftIconSize As New Size(18, 18)
    Public Property LeftIconSize As Size
        Get
            Return _leftIconSize
        End Get
        Set(value As Size)
            _leftIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rightIconSize As New Size(18, 18)
    Public Property RightIconSize As Size
        Get
            Return _rightIconSize
        End Get
        Set(value As Size)
            _rightIconSize = value
            RecalculateItemHeight()
        End Set
    End Property

    Private _rightIconRightPadding As Integer = 6
    Public Property RightIconRightPadding As Integer
        Get
            Return _rightIconRightPadding
        End Get
        Set(value As Integer)
            _rightIconRightPadding = Math.Max(0, value)
            Me.Invalidate()
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
            TooltipTimer.Interval = value
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

    Private _headerBackColor As Color = Color.FromArgb(222, 222, 222)
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

    Private _searchPropertiesConfigured As Boolean = False

    Private _searchShow As Boolean = False
    Public Property SearchShow As Boolean
        Get
            Return _searchShow
        End Get
        Set(value As Boolean)
            _searchShow = value
        End Set
    End Property

    Private _searchDefaultText As String = ""
    Public Property SearchDefaultText As String
        Get
            Return _searchDefaultText
        End Get
        Set(value As String)
            _searchDefaultText = value
            ApplySearchPlaceholder()
        End Set
    End Property

    Private _searchType As En_Tree_SearchType = En_Tree_SearchType.SearchType_Contains
    Public Property SearchType As En_Tree_SearchType
        Get
            Return _searchType
        End Get
        Set(value As En_Tree_SearchType)
            _searchType = value
            _searchPropertiesConfigured = True
        End Set
    End Property

    Private _searchIn As En_Tree_SearchIn = En_Tree_SearchIn.SearchIn_Caption
    Public Property SearchIn As En_Tree_SearchIn
        Get
            Return _searchIn
        End Get
        Set(value As En_Tree_SearchIn)
            _searchIn = value
            _searchPropertiesConfigured = True
        End Set
    End Property

    Private _searchMode As En_Tree_SearchMode = En_Tree_SearchMode.SearchMode_Tree
    Public Property SearchMode As En_Tree_SearchMode
        Get
            Return _searchMode
        End Get
        Set(value As En_Tree_SearchMode)
            _searchMode = value
            _searchPropertiesConfigured = True
        End Set
    End Property


    Private _searchBackColor As Color = Color.FromArgb(222, 222, 222)
    Public Property SearchBackColor As Color
        Get
            Return _searchBackColor
        End Get
        Set(value As Color)
            _searchBackColor = value : Me.Invalidate()
        End Set
    End Property

    Private _searchBoxBackColor As Color = Color.Empty
    Public Property SearchBoxBackColor As Color
        Get
            Return _searchBoxBackColor
        End Get
        Set(value As Color)
            _searchBoxBackColor = value
            If _searchTextBox IsNot Nothing Then
                _searchTextBox.BackColor = If(value = Color.Empty, Me.BackColor, value)
            End If
            Me.Invalidate()
        End Set
    End Property

    Private _searchBarLabelText As String = "Cautare: "
    Public Property SearchBarLabelText As String
        Get
            Return _searchBarLabelText
        End Get
        Set(value As String)
            _searchBarLabelText = value
            _searchPropertiesConfigured = True
            If _searchBarLabel IsNot Nothing Then _searchBarLabel.Text = value
            Me.Invalidate()
        End Set
    End Property

    Private _searchBarLabelForeColor As Color = Color.Empty
    Public Property SearchBarLabelForeColor As Color
        Get
            Return _searchBarLabelForeColor
        End Get
        Set(value As Color)
            _searchBarLabelForeColor = value
            _searchPropertiesConfigured = True
            If _searchBarLabel IsNot Nothing Then _searchBarLabel.ForeColor = If(value = Color.Empty, _headerForeColor, value)
            Me.Invalidate()
        End Set
    End Property

    Private _searchBarLabelBold As Boolean = False
    Public Property SearchBarLabelBold As Boolean
        Get
            Return _searchBarLabelBold
        End Get
        Set(value As Boolean)
            _searchBarLabelBold = value
            _searchPropertiesConfigured = True
            UpdateSearchBarLabelFont()
            Me.Invalidate()
        End Set
    End Property

    Private _searchBarLabelItalic As Boolean = False
    Public Property SearchBarLabelItalic As Boolean
        Get
            Return _searchBarLabelItalic
        End Get
        Set(value As Boolean)
            _searchBarLabelItalic = value
            _searchPropertiesConfigured = True
            UpdateSearchBarLabelFont()
            Me.Invalidate()
        End Set
    End Property

    Private Sub UpdateSearchBarLabelFont()
        If _searchBarLabel Is Nothing Then Return
        Dim style As FontStyle = FontStyle.Regular
        If _searchBarLabelBold Then style = style Or FontStyle.Bold
        If _searchBarLabelItalic Then style = style Or FontStyle.Italic
        _searchBarLabel.Font = New Font(Me.Font, style)
    End Sub

    Friend Sub MarkSearchConfigured()
        _searchPropertiesConfigured = True
    End Sub

    ' ── Search TextBox font ───────────────────────────────────────────────────
    Private _searchBarFontName As String = "Calibri"
    Public Property SearchBarFontName As String
        Get
            Return _searchBarFontName
        End Get
        Set(value As String)
            _searchBarFontName = value
            _searchPropertiesConfigured = True
            UpdateSearchTextBoxFont()
        End Set
    End Property

    Private _searchBarFontSize As Single = 10
    Public Property SearchBarFontSize As Single
        Get
            Return _searchBarFontSize
        End Get
        Set(value As Single)
            _searchBarFontSize = value
            _searchPropertiesConfigured = True
            UpdateSearchTextBoxFont()
        End Set
    End Property

    Private _searchClearButton As Boolean = False
    Public Property SearchClearButton As Boolean
        Get
            Return _searchClearButton
        End Get
        Set(value As Boolean)
            _searchClearButton = value
        End Set
    End Property

    Private _scrollBarTheme As En_ScrollBarTheme = En_ScrollBarTheme.Explorer
    Public Property ScrollBarTheme As En_ScrollBarTheme
        Get
            Return _scrollBarTheme
        End Get
        Set(value As En_ScrollBarTheme)
            _scrollBarTheme = value
            ApplyScrollBarTheme()
        End Set
    End Property

    Private _tooltipShow As Boolean = True
    Public Property TooltipShow As Boolean
        Get
            Return _tooltipShow
        End Get
        Set(value As Boolean)
            _tooltipShow = value
        End Set
    End Property

    Private _tooltipBackColor As Color = Color.FromArgb(255, 255, 232)
    Public Property TooltipBackColor As Color
        Get
            Return _tooltipBackColor
        End Get
        Set(value As Color)
            _tooltipBackColor = value
        End Set
    End Property

    Private _tooltipForeColor As Color = Color.FromArgb(50, 50, 60)
    Public Property TooltipForeColor As Color
        Get
            Return _tooltipForeColor
        End Get
        Set(value As Color)
            _tooltipForeColor = value
        End Set
    End Property

    Private _treeListViewEnabled As Boolean = False          ' master switch
    Public Property TreeListView As Boolean
        Get
            Return _treeListViewEnabled
        End Get
        Set(value As Boolean)
            _treeListViewEnabled = value
            If Not value Then
                _activeColFilters.Clear()
                _colFilterActive = False
                _colFilterSet.Clear()
                _activeColFilterPopup?.Close()
                _activeColFilterPopup = Nothing
            End If
            Me.Invalidate()
        End Set
    End Property


    Friend Sub UpdateSearchTextBoxFont()
        If _searchTextBox Is Nothing Then Return
        Dim name As String = If(String.IsNullOrEmpty(_searchBarFontName), Me.Font.Name, _searchBarFontName)
        Dim size As Single = If(_searchBarFontSize <= 0, Me.Font.Size, _searchBarFontSize)
        _searchTextBox.Font = New Font(name, size)
    End Sub

    Private ReadOnly Property ScrollBarWidth As Integer
        Get
            Return If(_vScroll IsNot Nothing AndAlso _vScroll.Visible, _vScroll.Width, 0)
        End Get
    End Property
End Class
