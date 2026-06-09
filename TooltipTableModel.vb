' ============================================================================
'  TooltipTableModel.vb
'  TREEVIEW_VBA - model intern + parser pentru tooltip-uri de tip <table>
'
'  Conține:
'    - TableConfig          : configurarea globală a tabelului (cu default-uri)
'    - TtCell               : model per-celulă
'    - TtRow                : model per-rând (celule + fundal)
'    - TooltipTableModel    : modelul complet (config + header + rows + footer)
'    - TooltipTableParser   : IsTableXml() + TryParse() (XmlDocument-based)
'
'  CONTRACT DPI: toate dimensiunile în pixeli (CellPaddingH/V, RowHeight,
'  Width pe celulă, MaxWidth pe <table>) sosesc DEJA scalate de partea VBA.
'  VB.NET nu mai aplică nicio scalare. FontSize rămâne în puncte (GraphicsUnit.Point),
'  GDI+ îl tratează DPI-aware singur.
'
'  CONTRACT BUILDER (Faza 2 - clsTT_Table.ToXml):
'    - emite elementul <table> FĂRĂ prolog <?xml ?> (folosește documentElement.xml,
'      nu xmlDoc.xml). IsTableXml tolerează totuși un prolog, dacă apare.
'    - emite FontSize cu PUNCT zecimal (InvariantCulture), nu cu virgulă - altfel
'      Single.TryParse(InvariantCulture) de mai jos îl ignoră și cade pe default 8.5.
' ============================================================================

Imports System.Globalization
Imports System.Xml

' ----------------------------------------------------------------------------
'  CONFIGURARE GLOBALĂ TABEL
'  Default-urile de mai jos sunt UNICA sursă de adevăr pentru valorile implicite
'  (am consolidat aici constantele TBL_DEFAULT_* din planul §9, ca să nu existe
'   două surse care pot diverge).
' ----------------------------------------------------------------------------
Friend Class TableConfig
    Public FontName As String = "Segoe UI"
    Public FontSize As Single = 8.5F
    Public CellPaddingH As Integer = 6
    Public CellPaddingV As Integer = 3
    Public RowHeight As Integer = 0                              ' 0 = auto (din font)
    Public MaxWidth As Integer = 0                              ' 0 = nelimitat (clamp exterior, S2)
    Public GridVisible As Boolean = True
    Public GridColor As Color = Color.FromArgb(204, 204, 204)  ' #CCCCCC
    Public HeaderBackColor As Color = Color.FromArgb(68, 114, 196)   ' #4472C4
    Public HeaderForeColor As Color = Color.White                    ' #FFFFFF
    Public FooterItalic As Boolean = False
    Public FooterSeparator As Boolean = True
End Class

' ----------------------------------------------------------------------------
'  CELULĂ
' ----------------------------------------------------------------------------
Friend Class TtCell
    Public Text As String = ""
    Public Width As Integer = 0                        ' 0 = auto (px deja scalați DPI)
    Public Align As HorizontalAlignment = HorizontalAlignment.Left
    Public Bold As Boolean = False
    Public Italic As Boolean = False
    Public ForeColor As Color = Color.Empty              ' Empty = moștenit
    Public BackColor As Color = Color.Empty              ' Empty = transparent
End Class

' ----------------------------------------------------------------------------
'  RÂND
' ----------------------------------------------------------------------------
Friend Class TtRow
    Public Cells As New List(Of TtCell)
    Public BackColor As Color = Color.Empty                            ' Empty = fără fundal de rând
End Class

' ----------------------------------------------------------------------------
'  MODEL COMPLET
' ----------------------------------------------------------------------------
Friend Class TooltipTableModel
    Public Config As New TableConfig
    Public HeaderRow As TtRow = Nothing                                ' Nothing dacă lipsește <header>
    Public Rows As New List(Of TtRow)
    Public FooterRow As TtRow = Nothing                                ' Nothing dacă lipsește <footer>
    Public ColCount As Integer = 0                                    ' calculat la parse (max celule/rând)
End Class

' ----------------------------------------------------------------------------
'  PARSER
' ----------------------------------------------------------------------------
Friend NotInheritable Class TooltipTableParser

    Private Sub New() ' clasă pur statică - fără instanțiere
    End Sub

    ''' <summary>
    ''' Detecție rapidă, fără parse complet. Tolerează un eventual prolog &lt;?xml ?&gt;.
    ''' </summary>
    Friend Shared Function IsTableXml(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False

        Dim t As String = text.TrimStart()

        ' Sări peste eventualul prolog XML (<?xml ... ?>)
        If t.StartsWith("<?xml", StringComparison.OrdinalIgnoreCase) Then
            Dim idx As Integer = t.IndexOf("?>", StringComparison.Ordinal)
            If idx >= 0 Then t = t.Substring(idx + 2).TrimStart()
        End If

        Return t.StartsWith("<table", StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' Parse complet. Returnează True + model, sau False + errorMsg dacă XML-ul e invalid.
    ''' Rândurile cu număr inegal de celule NU sunt eroare (se tolerează la Measure/Paint).
    ''' </summary>
    Friend Shared Function TryParse(text As String,
                                ByRef model As TooltipTableModel,
                                ByRef errorMsg As String) As Boolean
        model = Nothing
        errorMsg = Nothing
        Try
            Dim doc As New XmlDocument()
            doc.LoadXml(text)

            Dim root As XmlElement = doc.DocumentElement
            If root Is Nothing OrElse root.Name <> "table" Then
                errorMsg = "Elementul rădăcină nu este <table>."
                Return False
            End If

            Dim m As New TooltipTableModel()
            ReadTableConfig(root, m.Config)

            Dim headerNode As XmlNode = root.SelectSingleNode("header")
            If headerNode IsNot Nothing Then m.HeaderRow = ReadRow(headerNode)

            For Each rowNode As XmlNode In root.SelectNodes("row")
                m.Rows.Add(ReadRow(rowNode))
            Next

            Dim footerNode As XmlNode = root.SelectSingleNode("footer")
            If footerNode IsNot Nothing Then m.FooterRow = ReadRow(footerNode)

            Dim colCount As Integer = 0
            If m.HeaderRow IsNot Nothing Then colCount = Math.Max(colCount, m.HeaderRow.Cells.Count)
            For Each r As TtRow In m.Rows
                colCount = Math.Max(colCount, r.Cells.Count)
            Next
            If m.FooterRow IsNot Nothing Then colCount = Math.Max(colCount, m.FooterRow.Cells.Count)
            m.ColCount = colCount

            If colCount = 0 Then
                errorMsg = "Tabelul nu conține nicio celulă."
                Return False
            End If

            model = m
            Return True
        Catch ex As Exception
            errorMsg = ex.Message
            Return False
        End Try
    End Function

    ' ========================================================================
    '  HELPERS PRIVATE
    ' ========================================================================

    Private Shared Sub ReadTableConfig(root As XmlElement, cfg As TableConfig)
        Dim s As String

        s = AttrStr(root, "FontName")
        If Not String.IsNullOrEmpty(s) Then cfg.FontName = s

        Dim fSize As Single
        If Single.TryParse(AttrStr(root, "FontSize"), NumberStyles.Float,
                           CultureInfo.InvariantCulture, fSize) AndAlso fSize > 0 Then
            cfg.FontSize = fSize
        End If

        Dim iv As Integer
        If Integer.TryParse(AttrStr(root, "CellPaddingH"), iv) AndAlso iv >= 0 Then cfg.CellPaddingH = iv
        If Integer.TryParse(AttrStr(root, "CellPaddingV"), iv) AndAlso iv >= 0 Then cfg.CellPaddingV = iv
        If Integer.TryParse(AttrStr(root, "RowHeight"), iv) AndAlso iv >= 0 Then cfg.RowHeight = iv
        If Integer.TryParse(AttrStr(root, "MaxWidth"), iv) AndAlso iv >= 0 Then cfg.MaxWidth = iv

        cfg.GridVisible = ParseBool(AttrStr(root, "GridVisible"), cfg.GridVisible)

        s = AttrStr(root, "GridColor")
        If Not String.IsNullOrEmpty(s) Then cfg.GridColor = AdvancedTreeControl.ParseColor(s, cfg.GridColor)

        s = AttrStr(root, "HeaderBackColor")
        If Not String.IsNullOrEmpty(s) Then cfg.HeaderBackColor = AdvancedTreeControl.ParseColor(s, cfg.HeaderBackColor)

        s = AttrStr(root, "HeaderForeColor")
        If Not String.IsNullOrEmpty(s) Then cfg.HeaderForeColor = AdvancedTreeControl.ParseColor(s, cfg.HeaderForeColor)

        cfg.FooterItalic = ParseBool(AttrStr(root, "FooterItalic"), cfg.FooterItalic)
        cfg.FooterSeparator = ParseBool(AttrStr(root, "FooterSeparator"), cfg.FooterSeparator)
    End Sub

    Private Shared Function ReadRow(rowNode As XmlNode) As TtRow
        Dim row As New TtRow()

        Dim bc As String = AttrStr(rowNode, "BackColor")
        If Not String.IsNullOrEmpty(bc) Then row.BackColor = AdvancedTreeControl.ParseColor(bc, Color.Empty)

        For Each cellNode As XmlNode In rowNode.SelectNodes("cell")
            row.Cells.Add(ReadCell(cellNode))
        Next

        Return row
    End Function

    Private Shared Function ReadCell(cellNode As XmlNode) As TtCell
        Dim c As New TtCell() With {.Text = cellNode.InnerText}

        Dim wv As Integer
        If Integer.TryParse(AttrStr(cellNode, "Width"), wv) AndAlso wv > 0 Then c.Width = wv

        c.Align = ParseAlign(AttrStr(cellNode, "Align"))
        c.Bold = ParseBool(AttrStr(cellNode, "Bold"), False)
        c.Italic = ParseBool(AttrStr(cellNode, "Italic"), False)

        Dim fc As String = AttrStr(cellNode, "ForeColor")
        If Not String.IsNullOrEmpty(fc) Then c.ForeColor = AdvancedTreeControl.ParseColor(fc, Color.Empty)

        Dim bc As String = AttrStr(cellNode, "BackColor")
        If Not String.IsNullOrEmpty(bc) Then c.BackColor = AdvancedTreeControl.ParseColor(bc, Color.Empty)

        Return c
    End Function

    Private Shared Function AttrStr(node As XmlNode, attrName As String) As String
        If node Is Nothing OrElse node.Attributes Is Nothing Then Return ""
        Dim a As XmlAttribute = node.Attributes(attrName)
        If a Is Nothing Then Return ""
        Return a.Value
    End Function

    ''' <summary>Convenție boolean identică cu restul applier-elor: "1" / "-1" / "true".</summary>
    Private Shared Function ParseBool(value As String, fallback As Boolean) As Boolean
        If String.IsNullOrEmpty(value) Then Return fallback
        Dim v As String = value.Trim().ToLowerInvariant()
        Return (v = "1" OrElse v = "-1" OrElse v = "true")
    End Function

    Private Shared Function ParseAlign(value As String) As HorizontalAlignment
        If String.IsNullOrEmpty(value) Then Return HorizontalAlignment.Left
        Select Case value.Trim().ToLowerInvariant()
            Case "right" : Return HorizontalAlignment.Right
            Case "center" : Return HorizontalAlignment.Center
            Case Else : Return HorizontalAlignment.Left
        End Select
    End Function

End Class