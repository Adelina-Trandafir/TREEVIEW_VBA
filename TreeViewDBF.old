Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' TreeView complet custom (DoubleBufferedTreeView):
'''  - DESENEAZĂ MANUAL:
'''      * linii ierarhie
'''      * expandere (+ / -)
'''      * indentare
'''      * icon principal stânga
'''      * icon opțional în dreapta (per nod)
'''      * text + background (normal / hover / pressed / selected)
'''  - double-buffer (Win32 + .NET)
'''  - evenimente:
'''      * NodeMouseDown / NodeMouseUp / NodeClick
'''      * DrawNodeExternal(sender, e, ByRef handled)
'''      * QueryNodeSelection(sender, e, ByRef paintSelection)
'''      * QueryAllowHoverOnSelectedNode(sender, e, ByRef allow)
'''      * RightIconMouseDown / RightIconMouseUp / RightIconClick / RightIconMouseMove
'''  - suport pentru imagini per nod (ImageKey / ImageIndex) + icon suplimentar în dreapta (tot din ImageList).
''' </summary>
Public Class DoubleBufferedTreeView
    Inherits TreeView

    ' ============================================================
    ' STRUCT / P/INVOKE
    ' ============================================================
    <StructLayout(LayoutKind.Sequential)>
    Private Structure TRACKMOUSEEVENTSTRUCT
        Public pCbSize As UInteger
        Public pDwFlags As UInteger
        Public pHwndTrack As IntPtr
        Public pDwHoverTime As UInteger
    End Structure

    Private Const pTME_LEAVE As Integer = &H2
    Private Const pWM_MOUSELEAVE As Integer = &H2A3

    Private Const pTV_FIRST As Integer = &H1100
    Private Const pTVM_SETEXTENDEDSTYLE As Integer = pTV_FIRST + 44
    Private Const pTVS_EX_DOUBLEBUFFER As Integer = &H4

    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202
    Private Const WM_LBUTTONDBLCLK As Integer = &H203

    <DllImport("user32.dll")>
    Private Shared Function TrackMouseEvent(ByRef pLpEventTrack As TRACKMOUSEEVENTSTRUCT) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function SendMessage(pHWnd As IntPtr, pMsg As Integer, pWParam As Integer, pLParam As Integer) As IntPtr
    End Function

    ' ============================================================
    ' CONSTANTE DESEN
    ' ============================================================
    Private Const pExpanderSize As Integer = 10     ' dimensiune pătrat expander
    Private Const pLineToTextGap As Integer = 8     ' distanța dintre linia principală și text
    Private Const pIconSize As Integer = 48         ' dimensiune "logică" icon stânga/dreapta
    Private Const pRightIconPadding As Integer = 4  ' padding față de marginea din dreapta

    ' ============================================================
    ' TIP INTERN: date pentru iconul din dreapta
    ' ============================================================
    ''' <summary>
    ''' Stochează informații pentru iconul din dreapta per nod.
    ''' </summary>
    Private Class RightIconData
        Public Property ImageKey As String = Nothing
        Public Property ImageIndex As Integer = -1
        Public Property Visible As Boolean = False
    End Class

    ' ============================================================
    ' EVENIMENTE PUBLICE
    ' ============================================================
    ''' <summary>
    ''' Se declanșează la MouseDown pe un nod (zona text/icoană/linie, NU pe right-icon).
    ''' </summary>
    Public Event NodeMouseDown(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' Se declanșează la MouseUp pe un nod (zona text/icoană/linie, NU pe right-icon).
    ''' </summary>
    Public Event NodeMouseUp(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' Click de nod (MouseDown + MouseUp pe același nod, în afara expander / right-icon).
    ''' </summary>
    Public Event NodeClick(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' Permite desen extern (custom). Dacă pHandled = True, controlul nu mai desenează textul/default.
    ''' </summary>
    Public Event DrawNodeExternal(pSender As Object, pE As DrawTreeNodeEventArgs, ByRef pHandled As Boolean)

    ''' <summary>
    ''' Permite să decizi dacă un nod se desenează ca "selectat".
    ''' </summary>
    Public Event QueryNodeSelection(pSender As Object, pE As DrawTreeNodeEventArgs, ByRef pPaintSelection As Boolean)

    ''' <summary>
    ''' Permite să decizi dacă efectul de hover se aplică și pe nodurile deja selectate.
    ''' </summary>
    Public Event QueryAllowHoverOnSelectedNode(pSender As Object, pE As DrawTreeNodeEventArgs, ByRef pAllow As Boolean)

    ''' <summary>
    ''' MouseDown pe iconul din dreapta al nodului.
    ''' </summary>
    Public Event RightIconMouseDown(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' MouseUp pe iconul din dreapta al nodului.
    ''' </summary>
    Public Event RightIconMouseUp(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' Click (MouseDown + MouseUp) pe iconul din dreapta al nodului.
    ''' </summary>
    Public Event RightIconClick(pNode As TreeNode, pE As MouseEventArgs)

    ''' <summary>
    ''' MouseMove peste iconul din dreapta al nodului.
    ''' </summary>
    Public Event RightIconMouseMove(pNode As TreeNode, pE As MouseEventArgs)

    ' ============================================================
    ' STARE INTERNĂ
    ' ============================================================
    Private pHoverNode As TreeNode = Nothing
    Private pPressedNode As TreeNode = Nothing
    Private pWasPressed As Boolean = False
    Private pTracking As Boolean = False

    ' stare pentru right-icon
    Private pRightIconPressedNode As TreeNode = Nothing
    Private pRightIconWasPressed As Boolean = False
    Private pLastRightIconHoverNode As TreeNode = Nothing

    ' map nod → date icon dreapta
    Private ReadOnly pRightIcons As New Dictionary(Of TreeNode, RightIconData)

    ' Dicționar global: TreeNode → NodeData
    Private ReadOnly NodeInfo As New Dictionary(Of TreeNode, NodeData)()

    ' Fade animation pentru right icon
    Private ReadOnly pRightIconAlphaCurrent As New Dictionary(Of TreeNode, Single)()
    Private ReadOnly pRightIconAlphaTarget As New Dictionary(Of TreeNode, Single)()
    Private WithEvents pFadeTimer As Timer
    Private Const pFadeSpeed As Single = 0.15F ' viteza fade (0.0-1.0 per tick)

    ' Variabila pentru functia de click dreapta
    Private _RightClickFunc As String = ""
    Private _useFader As Boolean = False

    ' ============================================================
    ' NODEDATA — date extinse ale nodului
    ' ============================================================
    Private Class NodeData
        Public Property NodeRect As Rectangle
        Public Property TextRect As Rectangle
        Public Property LeftIconRect As Rectangle
        Public Property RightIconRect As Rectangle
        Public Property ExpanderRect As Rectangle

        Public Property IsHovered As Boolean
        Public Property IsPressed As Boolean
        Public Property IsRightIconHovered As Boolean

        Public Property Payload As Object
    End Class

    ''' <summary>
    ''' Un TreeNode specializat care știe ce iconiță are când e deschis sau închis.
    ''' </summary>
    Public Class ExtendedTreeNode
        Inherits TreeNode

        ' Proprietăți noi
        Public Property IconClosed As String = ""
        Public Property IconOpen As String = ""

        Public Sub New(text As String)
            MyBase.New(text)
        End Sub
    End Class

    ' ============================================================
    ' CTOR
    ' ============================================================
    ''' <summary>
    ''' Inițializează un nou DoubleBufferedTreeView cu desen complet custom și double-buffer.
    ''' </summary>
    Public Sub New()
        MyBase.New()

        ' 1. Activăm desenarea custom
        Me.DrawMode = TreeViewDrawMode.OwnerDrawAll
        Me.ShowLines = False
        Me.ShowPlusMinus = False
        Me.ShowRootLines = False

        ' 2. OPRIREM DoubleBuffer-ul .NET pentru a evita conflictul și lag-ul
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, False)
        Me.DoubleBuffered = False

        ' 3. Păstrăm AllPaintingInWmPaint pentru a opri sclipirea albă de fundal
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)

        ' 4. Activăm redraw la redimensionare
        Me.SetStyle(ControlStyles.ResizeRedraw, True)

        ' 5. Activăm Double-Buffering NATIV (Win32) - mult mai rapid pentru TreeView
        ' Acesta desenează totul intern și pune rezultatul pe ecran instant
        SendMessage(Me.Handle, pTVM_SETEXTENDEDSTYLE, pTVS_EX_DOUBLEBUFFER, pTVS_EX_DOUBLEBUFFER)

        ' Timer animație
        If _useFader Then
            pFadeTimer = New Timer With {
            .Interval = 16,
            .Enabled = False
        }
        End If

    End Sub

    ''' <summary>
    ''' Obține dreptunghiul complet al nodului (inclusiv zona de text + background).
    ''' </summary>
    ''' <param name="n"></param>
    ''' <returns></returns>
    Public Function NodeRect(Optional n As TreeNode = Nothing) As Rectangle
        If n Is Nothing Then
            Return GetData(pHoverNode).NodeRect
        Else
            Return GetData(n).NodeRect
        End If
    End Function

    ''' <summary>
    ''' Obține dreptunghiul iconiței din dreapta (daca exista).
    ''' </summary>
    Public Function RightIconRect(Optional n As TreeNode = Nothing) As Rectangle
        ' Determinăm nodul țintă: cel primit ca parametru SAU cel peste care a fost mouse-ul
        Dim targetNode As TreeNode = If(n, pLastRightIconHoverNode)

        ' Dacă nu avem niciun nod valid, returnăm Empty (evităm eroarea)
        If targetNode Is Nothing Then
            Return Rectangle.Empty
        End If

        ' Obținem datele. GetData poate returna Nothing în cazuri rare, verificăm și asta
        Dim d = GetData(targetNode)
        If d Is Nothing Then
            Return Rectangle.Empty
        End If

        Return d.RightIconRect
    End Function

    ''' <summary>
    ''' Adaugă imagine în ImageList-ul intern, configurând automat ImageList dacă nu există încă.
    ''' </summary>
    ''' <param name="pKey"></param>
    ''' <param name="pImage"></param>
    Public Sub AddImage(pKey As String, pImage As Image)

        ' creează ImageList dacă nu există
        If Me.ImageList Is Nothing Then
            ' setează parametrii corecți pentru PNG 32-bit
            ' IMPORTANT: dimensiunea trebuie setată înainte de Add()
            Me.ImageList = New ImageList With {
                .ColorDepth = ColorDepth.Depth32Bit,
                .ImageSize = New Size(pIconSize, pIconSize)
            }
        End If

        ' dacă imaginea nu există deja, o adaugă
        If Not Me.ImageList.Images.ContainsKey(pKey) Then
            Me.ImageList.Images.Add(pKey, pImage)
        End If

    End Sub


    ' ============================================================
    ' API PUBLIC pentru ICON DREAPTA
    ' ============================================================
    ''' <summary>
    ''' Setează iconul din dreapta pentru un nod, folosind ImageKey din ImageList.
    ''' </summary>
    Public Sub SetRightIconKey(pNode As TreeNode, pImageKey As String, Optional pVisible As Boolean = True)
        If pNode Is Nothing Then Return

        Dim pData As RightIconData = Nothing
        If Not pRightIcons.TryGetValue(pNode, pData) Then
            pData = New RightIconData()
            pRightIcons(pNode) = pData
        End If

        pData.ImageKey = pImageKey
        pData.ImageIndex = -1
        pData.Visible = pVisible

        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Setează iconul din dreapta pentru un nod, folosind ImageIndex din ImageList.
    ''' </summary>
    Public Sub SetRightIconIndex(pNode As TreeNode, pImageIndex As Integer, Optional pVisible As Boolean = True)
        If pNode Is Nothing Then Return

        Dim pData As RightIconData = Nothing
        If Not pRightIcons.TryGetValue(pNode, pData) Then
            pData = New RightIconData()
            pRightIcons(pNode) = pData
        End If

        pData.ImageIndex = pImageIndex
        pData.ImageKey = Nothing
        pData.Visible = pVisible

        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Ascunde iconul din dreapta pentru un nod (nu șterge configurația).
    ''' </summary>
    Public Sub HideRightIcon(pNode As TreeNode)
        Dim pData As RightIconData = Nothing
        If pNode Is Nothing OrElse Not pRightIcons.TryGetValue(pNode, pData) Then Return

        pData.Visible = False
        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Șterge complet configurarea iconului din dreapta pentru un nod.
    ''' </summary>
    Public Sub ClearRightIcon(pNode As TreeNode)
        If pNode Is Nothing Then Return
        If pRightIcons.Remove(pNode) Then
            Me.Invalidate()
        End If
    End Sub

    ''' <summary>
    ''' Curăță cache-ul de date. Apelează asta după Nodes.Clear().
    ''' </summary>
    Public Sub ClearCache()
        NodeInfo.Clear()
        pRightIcons.Clear()
        pRightIconAlphaCurrent.Clear()
        pRightIconAlphaTarget.Clear()
        pLastRightIconHoverNode = Nothing
        pHoverNode = Nothing
    End Sub

    ' ============================================================
    ' DRAW NODE
    ' ============================================================
    ''' <summary>
    ''' Desenează nodul cu:
    '''  - background custom (normal / hover / pressed / selected)
    '''  - linii ierarhie
    '''  - expander (+/-)
    '''  - icon stânga (ImageKey/ImageIndex)
    '''  - icon dreapta (dacă este configurat)
    '''  - text
    '''  - delegare DrawNodeExternal
    ''' </summary>
    <System.Diagnostics.DebuggerStepThrough()>
    Protected Overrides Sub OnDrawNode(pE As DrawTreeNodeEventArgs)
        If pE Is Nothing OrElse pE.Node Is Nothing Then
            MyBase.OnDrawNode(pE)
            Return
        End If

        ' Protecție critică: Dacă nodul nu e vizibil sau bounds sunt 0, nu desenăm
        If pE.Node Is Nothing OrElse pE.Bounds.Width <= 0 OrElse pE.Bounds.Height <= 0 Then
            Return
        End If

        Dim pNode As TreeNode = pE.Node
        Dim pG As Graphics = pE.Graphics

        Dim pNodeBounds As Rectangle = pNode.Bounds
        If pNodeBounds.Height <= 0 Then
            pNodeBounds = New Rectangle(0, pE.Bounds.Top, Me.ClientSize.Width, Me.ItemHeight)
        End If

        Dim pIsSelectedNative As Boolean = (Me.SelectedNode Is pNode)
        Dim pPaintSelection As Boolean = pIsSelectedNative
        Dim pAllowHoverOnSelected As Boolean = False

        RaiseEvent QueryAllowHoverOnSelectedNode(Me, pE, pAllowHoverOnSelected)
        RaiseEvent QueryNodeSelection(Me, pE, pPaintSelection)

        Dim pIsSelected As Boolean = pPaintSelection
        Dim pIsHover As Boolean = (pNode Is pHoverNode) AndAlso (Not pIsSelectedNative OrElse pAllowHoverOnSelected)
        Dim pIsPressed As Boolean = (pNode Is pPressedNode)

        ' === CALCUL LAYOUT BAZĂ ===
        Dim pTextRect As Rectangle = GetTextRect(pNode)
        Dim pExpanderRect As Rectangle = GetExpanderRect(pNode)
        Dim pFullRowRect As New Rectangle(pTextRect.Left, pTextRect.Top, Me.ClientSize.Width - pTextRect.Left, pTextRect.Height)

        ' === 1. BACKGROUND (TEXT AREA) ===
        If pIsPressed Then
            Using pBrush As New SolidBrush(Color.FromArgb(150, SystemColors.Highlight))
                pG.FillRectangle(pBrush, pFullRowRect)
            End Using
        ElseIf pIsHover Then
            Using pBrush As New SolidBrush(Color.FromArgb(40, SystemColors.Highlight))
                pG.FillRectangle(pBrush, pFullRowRect)
            End Using
        ElseIf pIsSelected Then
            Using pBrush As New SolidBrush(SystemColors.Highlight)
                pG.FillRectangle(pBrush, pFullRowRect)
            End Using
        Else
            Using pBrush As New SolidBrush(Me.BackColor)
                pG.FillRectangle(pBrush, pFullRowRect)
            End Using
        End If

        ' === 2. LINES + EXPANDER + ICON STÂNGA ===
        DrawHierarchyLines(pG, pNode, pNodeBounds, pExpanderRect)
        DrawExpander(pG, pNode, pExpanderRect)
        Dim pLeftIconRect As Rectangle = DrawLeftIcon(pG, pNode, pTextRect)

        ' Ajustăm zona de text după iconul din stânga.
        Dim pEffectiveTextRect As Rectangle = pTextRect
        If Not pLeftIconRect.IsEmpty Then
            pEffectiveTextRect.X = pLeftIconRect.Right + 2
            pEffectiveTextRect.Width = Me.ClientSize.Width - pEffectiveTextRect.X
        End If

        ' === 3. ICON DREAPTA ===
        Dim pRightIconRect As Rectangle = DrawRightIcon(pG, pNode, pEffectiveTextRect)

        ' === 4. DESEN EXTERN (DELEGAT) ===
        Dim pHandled As Boolean = False
        Dim pArgs As New DrawTreeNodeEventArgs(
            pG,
            pNode,
            pEffectiveTextRect,
            pE.State
        )

        RaiseEvent DrawNodeExternal(Me, pArgs, pHandled)
        If pHandled Then
            Return
        End If

        ' === 5. TEXT ===
        Dim pForeColor As Color = If(pIsSelected, SystemColors.HighlightText, Me.ForeColor)

        TextRenderer.DrawText(
            pG,
            pNode.Text,
            Me.Font,
            pEffectiveTextRect,
            pForeColor,
            TextFormatFlags.VerticalCenter Or TextFormatFlags.Left Or TextFormatFlags.NoPrefix
        )

        ' === 6. SALVARE DATE NOD ===
        Dim d = GetData(pNode)
        d.NodeRect = pFullRowRect
        d.TextRect = pEffectiveTextRect
        d.LeftIconRect = pLeftIconRect
        d.RightIconRect = pRightIconRect
        d.ExpanderRect = pExpanderRect
    End Sub

    ' ============================================================
    ' CALCUL RECT TEXT
    ' ============================================================
    ''' <summary>
    ''' Calculează dreptunghiul de bază pentru textul nodului (fără icon dreapta).
    ''' </summary>
    Private Function GetTextRect(pNode As TreeNode) As Rectangle
        Dim pBounds As Rectangle = pNode.Bounds
        If pBounds.Height <= 0 Then
            pBounds = New Rectangle(0, pBounds.Top, Me.ClientSize.Width, Me.ItemHeight)
        End If

        Dim pTextLeft As Integer = pBounds.Left
        Dim pTextTop As Integer = pBounds.Top
        Dim pTextWidth As Integer = Me.ClientSize.Width - pTextLeft
        Dim pTextHeight As Integer = pBounds.Height
        If pTextHeight <= 0 Then pTextHeight = Me.ItemHeight

        Return New Rectangle(pTextLeft, pTextTop, pTextWidth, pTextHeight)
    End Function

    ' ============================================================
    ' CALCUL RECT EXPANDER
    ' ============================================================
    ''' <summary>
    ''' Calculează dreptunghiul pentru expander (+/-) al nodului.
    ''' </summary>
    Private Shared Function GetExpanderRect(pNode As TreeNode) As Rectangle
        If pNode Is Nothing Then Return Rectangle.Empty
        If pNode.Bounds.Height <= 0 Then Return Rectangle.Empty

        Dim pBounds As Rectangle = pNode.Bounds
        Dim pCenterY As Integer = pBounds.Top + pBounds.Height \ 2

        ' linia principală este la stânga textului
        Dim pLineX As Integer = pBounds.Left - pLineToTextGap

        Return New Rectangle(
            pLineX - pExpanderSize \ 2,
            pCenterY - pExpanderSize \ 2,
            pExpanderSize,
            pExpanderSize
        )
    End Function

    ' ============================================================
    ' DESENARE ICON STÂNGA
    ' ============================================================
    ''' <summary>
    ''' Desenează iconul standard al nodului în stânga textului (ImageList de pe control).
    ''' </summary>
    Private Function DrawLeftIcon(pG As Graphics, pNode As TreeNode, pTextRect As Rectangle) As Rectangle
        Dim pResult As Rectangle = Rectangle.Empty
        If Me.ImageList Is Nothing Then Return pResult

        Dim pImg As Image = Nothing
        If Not String.IsNullOrEmpty(pNode.ImageKey) AndAlso Me.ImageList.Images.ContainsKey(pNode.ImageKey) Then
            pImg = Me.ImageList.Images(pNode.ImageKey)
        ElseIf pNode.ImageIndex >= 0 AndAlso pNode.ImageIndex < Me.ImageList.Images.Count Then
            pImg = Me.ImageList.Images(pNode.ImageIndex)
        End If

        If pImg Is Nothing Then Return pResult

        Dim pSize As Integer = Math.Min(pIconSize, Math.Min(pImg.Width, pImg.Height))
        Dim pX As Integer = pTextRect.Left
        Dim pY As Integer = pTextRect.Top + (pTextRect.Height - pSize) \ 2

        pResult = New Rectangle(pX, pY, pSize, pSize)
        pG.DrawImage(pImg, pResult)

        Return pResult
    End Function

    ' ============================================================
    ' DESENARE ICON DREAPTA
    ' ============================================================
    ''' <summary>
    ''' Desenează iconul din dreapta pentru nod, dacă este configurat și vizibil.
    ''' Folosește alpha blending pentru fade-in/fade-out.
    ''' </summary>
    Private Function DrawRightIcon(pG As Graphics, pNode As TreeNode, pTextRect As Rectangle) As Rectangle
        Dim pResult As Rectangle = Rectangle.Empty
        If Me.ImageList Is Nothing Then Return pResult

        Dim pData As RightIconData = Nothing
        If Not pRightIcons.TryGetValue(pNode, pData) OrElse pData Is Nothing OrElse Not pData.Visible Then
            Return pResult
        End If

        Dim pImg As Image = Nothing
        If pData.ImageKey IsNot Nothing AndAlso Me.ImageList.Images.ContainsKey(pData.ImageKey) Then
            pImg = Me.ImageList.Images(pData.ImageKey)
        ElseIf pData.ImageIndex >= 0 AndAlso pData.ImageIndex < Me.ImageList.Images.Count Then
            pImg = Me.ImageList.Images(pData.ImageIndex)
        End If

        If pImg Is Nothing Then Return pResult

        Dim pSize As Integer = Math.Min(pIconSize, Math.Min(pImg.Width, pImg.Height))
        Dim pX As Integer = Me.ClientSize.Width - pSize - pRightIconPadding
        Dim pY As Integer = pTextRect.Top + (pTextRect.Height - pSize) \ 2

        pResult = New Rectangle(pX, pY, pSize, pSize)

        ' Obținem alpha curent pentru acest nod
        Dim alpha As Single = 0.0F
        If Not pRightIconAlphaCurrent.TryGetValue(pNode, alpha) Then
            alpha = 0.0F
            pRightIconAlphaCurrent(pNode) = alpha
        End If

        ' Desenăm doar dacă alpha > 0
        If alpha > 0.01F Then
            ' Creăm ImageAttributes pentru alpha blending
            Using imgAttr As New Imaging.ImageAttributes()
                Dim colorMatrix As New Imaging.ColorMatrix With {
                    .Matrix33 = alpha ' Setăm alpha-ul
                    }
                imgAttr.SetColorMatrix(colorMatrix, Imaging.ColorMatrixFlag.Default, Imaging.ColorAdjustType.Bitmap)

                ' Desenăm imaginea cu alpha
                pG.DrawImage(pImg, pResult, 0, 0, pImg.Width, pImg.Height, GraphicsUnit.Pixel, imgAttr)
            End Using
        End If

        Return pResult
    End Function

    ''' <summary>
    ''' Calculează rect-ul iconului din dreapta pentru un nod (fără desen).
    ''' </summary>
    Private Function GetRightIconRect(pNode As TreeNode) As Rectangle
        If Me.ImageList Is Nothing OrElse pNode Is Nothing Then Return Rectangle.Empty

        Dim pData As RightIconData = Nothing
        If Not pRightIcons.TryGetValue(pNode, pData) OrElse pData Is Nothing OrElse Not pData.Visible Then
            Return Rectangle.Empty
        End If

        Dim pBounds As Rectangle = pNode.Bounds
        If pBounds.Height <= 0 Then
            pBounds = New Rectangle(0, pBounds.Top, Me.ClientSize.Width, Me.ItemHeight)
        End If

        Dim pSize As Integer = pIconSize
        Dim pX As Integer = Me.ClientSize.Width - pSize - pRightIconPadding
        Dim pY As Integer = pBounds.Top + (pBounds.Height - pSize) \ 2

        Return New Rectangle(pX, pY, pSize, pSize)
    End Function

    ' ============================================================
    ' DESENARE EXPANDER (+/-)
    ' ============================================================
    ''' <summary>
    ''' Desenează expander-ul (+/-) pentru nod, dacă are copii.
    ''' </summary>
    Private Shared Sub DrawExpander(pG As Graphics, pNode As TreeNode, pExpanderRect As Rectangle)
        If pNode.Nodes.Count = 0 Then
            Return
        End If

        Using pPen As New Pen(SystemColors.ControlDark),
              pBrush As New SolidBrush(SystemColors.Window)
            pG.FillRectangle(pBrush, pExpanderRect)
            pG.DrawRectangle(pPen, pExpanderRect)
        End Using

        Dim pCenterX As Integer = pExpanderRect.Left + pExpanderRect.Width \ 2
        Dim pCenterY As Integer = pExpanderRect.Top + pExpanderRect.Height \ 2

        Using pPenLine As New Pen(SystemColors.ControlText)
            ' linie orizontală (minus)
            pG.DrawLine(pPenLine,
                        pCenterX - (pExpanderRect.Width \ 3),
                        pCenterY,
                        pCenterX + (pExpanderRect.Width \ 3),
                        pCenterY)

            ' linie verticală (plus) doar dacă e închis
            If Not pNode.IsExpanded Then
                pG.DrawLine(pPenLine,
                            pCenterX,
                            pCenterY - (pExpanderRect.Height \ 3),
                            pCenterX,
                            pCenterY + (pExpanderRect.Height \ 3))
            End If
        End Using
    End Sub

    ' ============================================================
    ' DESENARE LINES (IERARHIE)
    ' ============================================================
    ''' <summary>
    ''' Desenează liniile ierarhice (verticale/orizontale) pentru nod.
    ''' </summary>
    Private Sub DrawHierarchyLines(pG As Graphics, pNode As TreeNode, pNodeBounds As Rectangle, pExpanderRect As Rectangle)
        Using pPen As New Pen(SystemColors.ControlDark)
            Dim pTextRect As Rectangle = GetTextRect(pNode)
            Dim pCenterY As Integer = pTextRect.Top + pTextRect.Height \ 2

            ' linia orizontală de la expander spre text
            If Not pExpanderRect.IsEmpty Then
                Dim pLineX As Integer = pExpanderRect.Left + pExpanderRect.Width \ 2
                pG.DrawLine(pPen, pLineX, pCenterY, pTextRect.Left, pCenterY)
            End If

            ' linia verticală pentru nodul curent
            If Not pExpanderRect.IsEmpty Then
                Dim pLineX As Integer = pExpanderRect.Left + pExpanderRect.Width \ 2
                Dim pTopY As Integer = pNodeBounds.Top
                Dim pBottomY As Integer = pNodeBounds.Bottom

                ' sus: până la mijloc
                If pNode.Parent IsNot Nothing OrElse pNode.PrevNode IsNot Nothing Then
                    pG.DrawLine(pPen, pLineX, pTopY, pLineX, pCenterY)
                End If

                ' jos: doar dacă are NextNode (mai urmează frați)
                If pNode.NextNode IsNot Nothing Then
                    pG.DrawLine(pPen, pLineX, pCenterY, pLineX, pBottomY)
                End If
            End If

            ' linii verticale pentru fiecare strămoș care are frați după el
            Dim pAncestor As TreeNode = pNode.Parent
            Dim pLevel As Integer = pNode.Level - 1

            While pAncestor IsNot Nothing AndAlso pLevel >= 0
                If pAncestor.NextNode IsNot Nothing Then
                    Dim pAncestorBounds As Rectangle = pAncestor.Bounds
                    Dim pAncestorTextRect As Rectangle = GetTextRect(pAncestor)
                    Dim pLineX As Integer = pAncestorTextRect.Left - pLineToTextGap

                    pG.DrawLine(pPen,
                                pLineX,
                                pNodeBounds.Top,
                                pLineX,
                                pNodeBounds.Bottom)
                End If

                pAncestor = pAncestor.Parent
                pLevel -= 1
            End While
        End Using
    End Sub

    ''' <summary>
    ''' Procesează MouseDown pe noduri / right-icon / expander.
    ''' Vine din WndProc, nu MyBase.OnMouseDown
    ''' </summary>
    ''' <param name="pE"></param>
    Private Sub HandleMouseDownInternal(pE As MouseEventArgs)
        ' Prioritate 1: Expander
        Dim pExpanderNode As TreeNode = HitTestExpander(pE.Location)
        If pExpanderNode IsNot Nothing Then
            If pExpanderNode.IsExpanded Then
                pExpanderNode.Collapse()
            Else
                pExpanderNode.Expand()
            End If
            Me.Invalidate()
            Return
        End If

        ' Prioritate 2: Right-icon
        Dim pRightIconNode As TreeNode = HitTestNodeForRightIcon(pE.Location)
        If pRightIconNode IsNot Nothing Then
            pRightIconPressedNode = pRightIconNode
            pRightIconWasPressed = True
            Me.SelectedNode = pRightIconNode
            UpdateRightIconAlphaTargets() ' Actualizăm alpha pentru iconițe
            RaiseEvent RightIconMouseDown(pRightIconNode, pE)
            Me.Invalidate()
            Return
        End If

        ' Prioritate 3: Restul nodului
        Dim pNode As TreeNode = HitTestNode(pE.Location)
        If pNode IsNot Nothing Then
            Me.SelectedNode = pNode
            UpdateRightIconAlphaTargets() ' Actualizăm alpha pentru iconițe
            pPressedNode = pNode
            pWasPressed = True
            GetData(pNode).IsPressed = True
            RaiseEvent NodeMouseDown(pNode, pE)
            Me.Invalidate()
        End If
    End Sub

    ''' <summary>
    ''' Procesează MouseUp pe noduri / right-icon.
    ''' Vine din WndProc, nu MyBase.OnMouseUp
    ''' </summary>
    ''' <param name="pE"></param>
    Private Sub HandleMouseUpInternal(pE As MouseEventArgs)
        ' Prioritate 1: Right-icon
        Dim pRightIconNode As TreeNode = HitTestNodeForRightIcon(pE.Location)
        If pRightIconNode IsNot Nothing Then

            ' === FIX: Eliberăm captura și resetăm cursorul AICI ===
            ' Deoarece în WndProc sărim peste MyBase.WndProc, trebuie să facem noi curățenie.
            If Me.Capture Then Me.Capture = False
            Me.Cursor = Cursors.Default
            ' ======================================================

            RaiseEvent RightIconMouseUp(pRightIconNode, pE)
            If pRightIconWasPressed AndAlso Object.ReferenceEquals(pRightIconPressedNode, pRightIconNode) Then
                RaiseEvent RightIconClick(pRightIconNode, pE)
            End If
            pRightIconWasPressed = False
            pRightIconPressedNode = Nothing
            Me.Invalidate()
            Return
        End If

        ' Prioritate 2: Restul nodului
        Dim pNode As TreeNode = HitTestNode(pE.Location)
        If pNode IsNot Nothing Then
            RaiseEvent NodeMouseUp(pNode, pE)
            If pWasPressed AndAlso Object.ReferenceEquals(pNode, pPressedNode) Then
                RaiseEvent NodeClick(pNode, pE)
            End If
            GetData(pNode).IsPressed = False
        End If

        pWasPressed = False
        pPressedNode = Nothing
        Me.Invalidate()
    End Sub

    ''' <summary>
    ''' Obține datele extinse ale unui nod.
    ''' </summary>
    ''' <param name="n"></param>
    ''' <returns></returns>
    Private Function GetData(n As TreeNode) As NodeData
        If n Is Nothing Then Return Nothing

        Dim d As NodeData = Nothing
        If Not NodeInfo.TryGetValue(n, d) Then
            d = New NodeData()
            NodeInfo(n) = d
        End If
        Return d
    End Function

    ''' <summary>
    ''' Hit-test pentru right-icon - returnează nodul dacă punctul e în RightIconRect.
    ''' </summary>
    Private Function HitTestNodeForRightIcon(pt As Point) As TreeNode
        For Each kv In NodeInfo
            Dim n = kv.Key
            Dim d = kv.Value
            If d.RightIconRect.Contains(pt) Then
                Return n
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Hit-test pentru expander - returnează nodul dacă punctul e în ExpanderRect.
    ''' </summary>
    Private Function HitTestExpander(pt As Point) As TreeNode
        For Each kv In NodeInfo
            Dim n = kv.Key
            Dim d = kv.Value
            If Not d.ExpanderRect.IsEmpty AndAlso d.ExpanderRect.Contains(pt) Then
                Return n
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Hit-test pentru întregul nod (expander + icon stânga + text + icon dreapta).
    ''' </summary>
    Private Function HitTestNode(pt As Point) As TreeNode
        For Each kv In NodeInfo
            Dim n = kv.Key
            Dim d = kv.Value

            ' Verifică expander
            If Not d.ExpanderRect.IsEmpty AndAlso d.ExpanderRect.Contains(pt) Then
                Return n
            End If

            ' Verifică icon stânga
            If Not d.LeftIconRect.IsEmpty AndAlso d.LeftIconRect.Contains(pt) Then
                Return n
            End If

            ' Verifică zona text (care include tot spațiul până la right icon)
            If Not d.TextRect.IsEmpty AndAlso d.TextRect.Contains(pt) Then
                Return n
            End If

            ' Verifică right icon
            If Not d.RightIconRect.IsEmpty AndAlso d.RightIconRect.Contains(pt) Then
                Return n
            End If
        Next
        Return Nothing
    End Function

    ' ============================================================
    ' MOUSE MOVE / HOVER
    ' ============================================================
    ''' <summary>
    ''' Gestionează starea de hover pentru noduri și right-icon, pornește TrackMouseEvent pentru WM_MOUSELEAVE.
    ''' </summary>
    <System.Diagnostics.DebuggerStepThrough()>
    Protected Overrides Sub OnMouseMove(pE As MouseEventArgs)
        MyBase.OnMouseMove(pE)

        If Not pTracking Then
            Dim pT As New TRACKMOUSEEVENTSTRUCT With {
                .pCbSize = CUInt(Marshal.SizeOf(GetType(TRACKMOUSEEVENTSTRUCT))),
                .pDwFlags = pTME_LEAVE,
                .pHwndTrack = Me.Handle,
                .pDwHoverTime = 0
            }
            TrackMouseEvent(pT)
            pTracking = True
        End If

        Dim pNode As TreeNode = If(HitTestNodeForRightIcon(pE.Location), Me.GetNodeAt(pE.X, pE.Y))

        If pNode IsNot Nothing Then
            Dim d = GetData(pNode)
            d.IsHovered = True

            ' reset previous hover
            If pHoverNode IsNot Nothing AndAlso Not Object.ReferenceEquals(pHoverNode, pNode) Then
                GetData(pHoverNode).IsHovered = False
            End If
        End If

        ' Right-icon hover & mouse move
        If pNode IsNot Nothing Then
            Dim pRightRect As Rectangle = GetRightIconRect(pNode)
            If Not pRightRect.IsEmpty AndAlso pRightRect.Contains(pE.Location) Then
                If Not Object.ReferenceEquals(pLastRightIconHoverNode, pNode) Then
                    pLastRightIconHoverNode = pNode
                End If
                Me.Cursor = Cursors.Hand
                'RaiseEvent RightIconMouseMove(pNode, pE)
            Else
                Me.Cursor = Cursors.Default
                pLastRightIconHoverNode = Nothing
            End If
        Else
            pLastRightIconHoverNode = Nothing
        End If

        ' Hover pe nod (pentru background)
        If pNode Is Nothing Then
            If pHoverNode IsNot Nothing Then
                pHoverNode = Nothing
                Me.Invalidate()
            End If
            ' Resetăm alpha țintă pentru toate nodurile când nu e hover pe nimic
            UpdateRightIconAlphaTargets()
            Return
        End If

        If Not Object.ReferenceEquals(pNode, pHoverNode) Then
            pHoverNode = pNode
            Me.Invalidate()
        End If

        ' Actualizăm alpha țintă pentru iconițe
        UpdateRightIconAlphaTargets()
    End Sub

    ' ============================================================
    ' UPDATE ALPHA TARGETS
    ' ============================================================
    ''' <summary>
    ''' Actualizează alpha țintă pentru toate nodurile în funcție de starea lor (hover/selected).
    ''' </summary>
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub UpdateRightIconAlphaTargets()
        ' Procesăm toate nodurile din NodeInfo
        For Each kv In NodeInfo
            Dim node As TreeNode = kv.Key
            Dim isSelected As Boolean = (Me.SelectedNode Is node)
            Dim isHovered As Boolean = (pHoverNode Is node)

            ' Setăm alpha țintă:
            ' - 1.0 dacă nodul e selectat sau hover
            ' - 0.0 altfel
            Dim targetAlpha As Single = If(isSelected OrElse isHovered, 1.0F, 0.0F)
            pRightIconAlphaTarget(node) = targetAlpha

            If _useFader AndAlso Not pFadeTimer.Enabled Then
                pFadeTimer.Start()
            End If
        Next
    End Sub

    ''' <summary>
    ''' Override pentru a actualiza alpha targets când se schimbă selecția.
    ''' </summary>
    Protected Overrides Sub OnAfterSelect(e As TreeViewEventArgs)
        MyBase.OnAfterSelect(e)
        UpdateRightIconAlphaTargets()
        Me.Invalidate()
    End Sub

    ' ============================================================
    ' WNDPROC — MOUSELEAVE
    ' ============================================================
    ''' <summary>
    ''' Interceptează WM_MOUSELEAVE pentru a reseta hover-ul.
    ''' Interceptează și WM_LBUTTONDOWN/UP pentru a gestiona corect MouseDown/Up.
    ''' </summary>
    '<System.Diagnostics.DebuggerStepThrough()>
    Protected Overrides Sub WndProc(ByRef pM As Message)

        ' Interceptează MouseDown
        If pM.Msg = WM_LBUTTONDOWN Then
            Dim pt As Point = Me.PointToClient(Cursor.Position)
            Dim foundSomething As Boolean = False

            If HitTestExpander(pt) IsNot Nothing Then
                foundSomething = True
            ElseIf HitTestNodeForRightIcon(pt) IsNot Nothing Then
                foundSomething = True
            ElseIf HitTestNode(pt) IsNot Nothing Then
                foundSomething = True
            End If

            HandleMouseDownInternal(New MouseEventArgs(MouseButtons.Left, 1, pt.X, pt.Y, 0))

            If Not foundSomething Then
                MyBase.WndProc(pM)
            End If
            Return
        End If

        ' Interceptează MouseUp
        If pM.Msg = WM_LBUTTONUP Then
            Dim pt As Point = Me.PointToClient(Cursor.Position)
            Dim foundSomething As Boolean = False

            If HitTestNodeForRightIcon(pt) IsNot Nothing Then
                foundSomething = True
            ElseIf HitTestNode(pt) IsNot Nothing Then
                foundSomething = True
            End If

            HandleMouseUpInternal(New MouseEventArgs(MouseButtons.Left, 1, pt.X, pt.Y, 0))

            If Not foundSomething Then
                MyBase.WndProc(pM)
            End If
            Return
        End If

        ' MouseLeave
        If pM.Msg = pWM_MOUSELEAVE Then
            pTracking = False
            pHoverNode = Nothing
            pLastRightIconHoverNode = Nothing
            For Each kv In NodeInfo
                kv.Value.IsHovered = False
                kv.Value.IsRightIconHovered = False
            Next
            UpdateRightIconAlphaTargets()
            Me.Invalidate()
        End If

        MyBase.WndProc(pM)
    End Sub

    ' ============================================================
    ' FADE ANIMATION TIMER
    ' ============================================================
    ''' <summary>
    ''' Handler pentru timer-ul de fade animation. Interpolează alpha curent către alpha țintă.
    ''' </summary>
    Private Sub pFadeTimer_Tick(sender As Object, e As EventArgs) Handles pFadeTimer.Tick
        Dim needsRedraw As Boolean = False

        ' Listăm toate nodurile din NodeInfo pentru a procesa alpha-ul lor
        Dim nodesToUpdate As New List(Of TreeNode)()
        For Each kv In NodeInfo
            nodesToUpdate.Add(kv.Key)
        Next

        For Each node In nodesToUpdate
            ' Obținem alpha curent și țintă
            Dim currentAlpha As Single = 0.0F
            Dim targetAlpha As Single = 0.0F

            If Not pRightIconAlphaCurrent.TryGetValue(node, currentAlpha) Then
                currentAlpha = 0.0F
                pRightIconAlphaCurrent(node) = currentAlpha
            End If

            If Not pRightIconAlphaTarget.TryGetValue(node, targetAlpha) Then
                targetAlpha = 0.0F
                pRightIconAlphaTarget(node) = targetAlpha
            End If

            ' Interpolăm
            If Math.Abs(currentAlpha - targetAlpha) > 0.01F Then
                If currentAlpha < targetAlpha Then
                    currentAlpha = Math.Min(currentAlpha + pFadeSpeed, targetAlpha)
                Else
                    currentAlpha = Math.Max(currentAlpha - pFadeSpeed, targetAlpha)
                End If
                pRightIconAlphaCurrent(node) = currentAlpha
                needsRedraw = True
            End If
        Next

        If needsRedraw Then
            Me.Invalidate()
        Else
            pFadeTimer.Stop()
        End If
    End Sub
End Class