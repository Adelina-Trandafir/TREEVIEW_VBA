Imports System.ComponentModel
Imports System.Drawing

' ==============================================================================
' NodeDebugInfo.vb
'
' Clasa descriptor pentru PropertyGrid in FrmNodeDebug.
' Toate proprietatile vin direct din TreeItem si din calculul vizual
' al AdvancedTreeControl — zero placeholders, zero ADAPT.
'
' 4 categorii:
'   1. Model         — toate campurile din TreeItem (publice + Friend)
'   2. Layout        — dreptunghiuri calculate exact ca in DrawItem / Painting.vb
'   3. Cells         — celulele TreeListView ale nodului
'   4. Renderer      — setarile controlului care afecteaza desenarea acestui nod
' ==============================================================================

Public Class NodeDebugInfo

#Region "── 1. MODEL ─────────────────────────────────────────────────────────"

    <Category("1. Model"),
     DisplayName("Key"),
     Description("Cheia unica a nodului.")>
    Public Property Key As String = ""

    <Category("1. Model"),
     DisplayName("Caption"),
     Description("Textul afisat al nodului.")>
    Public Property Caption As String = ""

    <Category("1. Model"),
     DisplayName("Level"),
     Description("Nivelul ierarhic (0 = root).")>
    Public Property Level As Integer

    <Category("1. Model"),
     DisplayName("Expanded"),
     Description("Nodul este expandat?")>
    Public Property Expanded As Boolean

    <Category("1. Model"),
     DisplayName("ParentKey"),
     Description("Cheia nodului parinte. String.Empty daca este root.")>
    Public Property ParentKey As String = ""

    <Category("1. Model"),
     DisplayName("ChildCount"),
     Description("Numar copii directi (Children.Count).")>
    Public Property ChildCount As Integer

    <Category("1. Model"),
     DisplayName("CheckState"),
     Description("Starea checkbox-ului: Unchecked / Checked / Indeterminate.")>
    Public Property CheckState As String = "Unchecked"

    <Category("1. Model"),
     DisplayName("HasCheckBox"),
     Description("Nodul afiseaza un checkbox (TreeItem.HasCheckBox)?")>
    Public Property HasCheckBox As Boolean

    <Category("1. Model"),
     DisplayName("Bold"),
     Description("Text bold?")>
    Public Property Bold As Boolean

    <Category("1. Model"),
     DisplayName("Italic"),
     Description("Text italic?")>
    Public Property Italic As Boolean

    <Category("1. Model"),
     DisplayName("NodeForeColor"),
     Description("Culoarea textului nodului. Color.Empty = mosteneste ForeColor-ul controlului.")>
    Public Property NodeForeColor As String = "Empty (mosteneste control)"

    <Category("1. Model"),
     DisplayName("NodeBackColor"),
     Description("Culoarea de fundal a nodului. Color.Empty = transparent.")>
    Public Property NodeBackColor As String = "Empty (transparent)"

    <Category("1. Model"),
     DisplayName("Tooltip"),
     Description("Textul tooltip (= VBA ControlTipText, trimis ca Tooltip in XML).")>
    Public Property Tooltip As String = ""

    <Category("1. Model"),
     DisplayName("Tag"),
     Description("Campul Tag al nodului (Object).")>
    Public Property Tag As String = ""

    <Category("1. Model"),
     DisplayName("LazyNode"),
     Description("Copiii se incarca din VBA la primul expand?")>
    Public Property LazyNode As Boolean

    <Category("1. Model"),
     DisplayName("ShowRightIconOnHover"),
     Description("Iconita dreapta vizibila doar la hover pe ACEST nod (suprascrie setarea globala a controlului)?")>
    Public Property ShowRightIconOnHover As Boolean

    <Category("1. Model"),
     DisplayName("IsLoader"),
     Description("Nod de tip loader/spinner (afisat in timp ce se incarca copiii LazyNode)?")>
    Public Property IsLoader As Boolean

    <Category("1. Model"),
     DisplayName("IsRadioSelected"),
     Description("Nodul este selectat in grupul RadioButton?")>
    Public Property IsRadioSelected As Boolean

    <Category("1. Model"),
     DisplayName("IsLastSibling"),
     Description("Este ultimul copil in lista parintelui sau, daca e root, ultimul root? (afecteaza desenarea liniilor de arbore)")>
    Public Property IsLastSibling As Boolean

    <Category("1. Model"),
     DisplayName("ColHeaderText"),
     Description("Coloane dinamice TreeListView, pipe-separated (ex: ""Col1|Col2""). Populat rar.")>
    Public Property ColHeaderText As String = ""

    <Category("1. Model"),
     DisplayName("HasLeftIconClosed"),
     Description("Nodul are imaginea LeftIconClosed incarcata?")>
    Public Property HasLeftIconClosed As Boolean

    <Category("1. Model"),
     DisplayName("HasLeftIconOpen"),
     Description("Nodul are imaginea LeftIconOpen incarcata?")>
    Public Property HasLeftIconOpen As Boolean

    <Category("1. Model"),
     DisplayName("HasRightIcon"),
     Description("Nodul are imaginea RightIcon incarcata?")>
    Public Property HasRightIcon As Boolean

    <Category("1. Model"),
     DisplayName("TextWidth_Cache"),
     Description("Cache pentru latimea textului in pixeli (Friend TextWidth). -1 = necalculat inca.")>
    Public Property TextWidth_Cache As Integer = -1

    <Category("1. Model"),
     DisplayName("LastClickedColumnIndex"),
     Description("Indexul ultimei coloane pe care s-a dat click (Friend, reset la -1 la fiecare MouseDown).")>
    Public Property LastClickedColumnIndex As Integer = -1

    <Category("1. Model"),
     DisplayName("LastClickedColumnName"),
     Description("Numele ultimei coloane pe care s-a dat click (Friend).")>
    Public Property LastClickedColumnName As String = ""

#End Region

#Region "── 2. LAYOUT ────────────────────────────────────────────────────────"
    <Category("2. Layout"),
     DisplayName("SelectionBounds"),
     Description("Dreptunghiul rounded al selectiei/hover (fullRowRect din DrawSelection). " &
                 "X = gridLeft + ExpanderSize*2 - 3 (sau gridLeft daca Level=0 si RootExpander=False).")>
    Public Property SelectionBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("NodeBounds"),
     Description("Dreptunghiul intregului rand (X=0, Y=GetItemY, W=control.Width, H=ItemHeight).")>
    Public Property NodeBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("ExpanderBounds"),
     Description("Dreptunghiul expanderului [+]/[-] (GetExpanderRect). Rectangle.Empty daca nodul nu are copii sau RootExpander=False.")>
    Public Property ExpanderBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("CheckBoxBounds"),
     Description("Dreptunghiul checkbox-ului (GetCheckBoxRect). Rectangle.Empty daca nu exista checkbox.")>
    Public Property CheckBoxBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("LeftIconBounds"),
     Description("Dreptunghiul iconitei din stanga (calculat ca in DrawContent). Rectangle.Empty daca lipseste iconita sau HasNodeIcons=False.")>
    Public Property LeftIconBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("TextBounds"),
     Description("Dreptunghiul disponibil pentru text (calculat ca in DrawContent). Include toata zona de la textX pana la RightIcon sau marginea dreapta.")>
    Public Property TextBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("RightIconBounds"),
     Description("Dreptunghiul iconitei din dreapta (calculat ca in DrawRightIcon). Rectangle.Empty daca nodul nu are RightIcon.")>
    Public Property RightIconBounds As Rectangle

    <Category("2. Layout"),
     DisplayName("GridLeft"),
     Description("X-ul de start al grilei = (Level * Indent) + AutoScrollPosition.X + PADDING_TREE_START(10).")>
    Public Property GridLeft As Integer

    <Category("2. Layout"),
     DisplayName("XBase"),
     Description("X-ul de start al continutului (dupa expander+gap). = GridLeft + Indent + PADDING_EXPANDER_GAP(12), sau GridLeft daca Level=0 si RootExpander=False.")>
    Public Property XBase As Integer

    <Category("2. Layout"),
     DisplayName("MidY"),
     Description("Y-ul central al randului = NodeBounds.Y + ItemHeight / 2. Folosit pentru alinierea expander, checkbox, icon.")>
    Public Property MidY As Integer

    <Category("2. Layout"),
     DisplayName("IndexInVisibleList"),
     Description("Pozitia nodului in lista GetVisibleItems() (0-based). Tinand cont de scroll si expandare.")>
    Public Property IndexInVisibleList As Integer = -1

    <Category("2. Layout"),
     DisplayName("IsInViewport"),
     Description("Randul este in zona vizibila a controlului (NodeBounds.Bottom > 0 si NodeBounds.Y < control.Height)?")>
    Public Property IsInViewport As Boolean

    <Category("2. Layout"),
     DisplayName("ExpanderCenterX"),
     Description("X-ul central al expanderului = GridLeft + Indent / 2. Folosit si pentru liniile de arbore (trunchiul vertical).")>
    Public Property ExpanderCenterX As Integer

#End Region

#Region "── 3. CELLS (TreeListView) ─────────────────────────────────────────"

    <Category("3. Cells"),
     DisplayName("CellCount"),
     Description("Numarul de celule TreeListView ale nodului (Cells.Count).")>
    Public Property CellCount As Integer

    <Category("3. Cells"),
     DisplayName("CellsData"),
     Description("Continutul celulelor in format: ""Col: Val [BackColor] [ForeColor]"". Gol daca nodul nu are celule.")>
    Public Property CellsData As String = ""

#End Region

#Region "── 4. RENDERER (setarile controlului pentru acest nod) ──────────────"
    <Category("4. Renderer"),
     DisplayName("TreeControl_Bounds"),
     Description("Bounds-ul AdvancedTreeControl relativ la form-ul parinte (Location + Size).")>
    Public Property TreeControl_Bounds As Rectangle

    <Category("4. Renderer"),
     DisplayName("ParentForm_Size"),
     Description("Dimensiunea form-ului parinte (FindForm().Size).")>
    Public Property ParentForm_Size As String = ""

    <Category("4. Renderer"),
     DisplayName("ParentForm_Bounds"),
     Description("Bounds-ul form-ului parinte pe ecran (FindForm().Bounds — pozitie absoluta + dimensiune).")>
    Public Property ParentForm_Bounds As String = ""

    <Category("4. Renderer"),
     DisplayName("ItemHeight"),
     Description("Inaltimea randului in pixeli (AdvancedTreeControl.ItemHeight).")>
    Public Property ItemHeight As Integer

    <Category("4. Renderer"),
     DisplayName("Indent"),
     Description("Indentarea pe nivel in pixeli (AdvancedTreeControl.Indent, default 20).")>
    Public Property Indent As Integer

    <Category("4. Renderer"),
     DisplayName("ExpanderSize"),
     Description("Dimensiunea patratului expanderului in pixeli (AdvancedTreeControl.ExpanderSize, default 12).")>
    Public Property ExpanderSize As Integer

    <Category("4. Renderer"),
     DisplayName("CheckBoxSize"),
     Description("Dimensiunea checkbox-ului in pixeli (AdvancedTreeControl.CheckBoxSize, default 16).")>
    Public Property CheckBoxSize As Integer

    <Category("4. Renderer"),
     DisplayName("LeftIconSize"),
     Description("Dimensiunea iconitei din stanga (AdvancedTreeControl.LeftIconSize, default 24x24).")>
    Public Property LeftIconSize As String = ""

    <Category("4. Renderer"),
     DisplayName("RightIconSize"),
     Description("Dimensiunea iconitei din dreapta (AdvancedTreeControl.RightIconSize, default 14x14).")>
    Public Property RightIconSize As String = ""

    <Category("4. Renderer"),
     DisplayName("HasNodeIcons"),
     Description("Controlul afiseaza iconite pe noduri (AdvancedTreeControl.HasNodeIcons)?")>
    Public Property HasNodeIcons As Boolean

    <Category("4. Renderer"),
     DisplayName("CheckBoxes"),
     Description("Controlul are modul CheckBoxes activ (AdvancedTreeControl.CheckBoxes)?")>
    Public Property CheckBoxes As Boolean

    <Category("4. Renderer"),
     DisplayName("RootExpander"),
     Description("Nodurile root au expander (AdvancedTreeControl.RootExpander)?")>
    Public Property RootExpander As Boolean

    <Category("4. Renderer"),
     DisplayName("ShowRightIconOnHover_Global"),
     Description("Setarea globala ShowRightIconOnHover a controlului (iconitele dreapta sunt hover-only global)?")>
    Public Property ShowRightIconOnHover_Global As Boolean

    <Category("4. Renderer"),
     DisplayName("ControlWidth"),
     Description("Latimea curenta a controlului AdvancedTreeControl in pixeli.")>
    Public Property ControlWidth As Integer

    <Category("4. Renderer"),
     DisplayName("ControlHeight"),
     Description("Inaltimea curenta a controlului AdvancedTreeControl in pixeli.")>
    Public Property ControlHeight As Integer

    <Category("4. Renderer"),
     DisplayName("ScrollBarVisible"),
     Description("Bara de scroll verticala este vizibila?")>
    Public Property ScrollBarVisible As Boolean

    <Category("4. Renderer"),
     DisplayName("ScrollOffsetY"),
     Description("Valoarea curenta a scroll-ului vertical (AutoScrollPosition.Y negativ → stocam pozitiv).")>
    Public Property ScrollOffsetY As Integer

    <Category("4. Renderer"),
     DisplayName("IsSelectedNode"),
     Description("Nodul este cel selectat curent in control (pSelectedItem)?")>
    Public Property IsSelectedNode As Boolean

    <Category("4. Renderer"),
     DisplayName("IsHoveredNode"),
     Description("Nodul este cel pe care se afla cursorul (pHoveredItem)?")>
    Public Property IsHoveredNode As Boolean

#End Region

End Class