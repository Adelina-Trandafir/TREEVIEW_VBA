' ============================================================================
'  TooltipTableParseTest.vb   (TEMPORAR - de șters după validare)
'
'  Test in-assembly pentru TooltipTableParser. NU folosește VBA / Shell /
'  WM_SETTEXT - doar string-uri hardcodate. Trebuie să stea în proiectul
'  TREEVIEW_VBA_64 fiindcă TableConfig / TtCell / TtRow / TooltipTableModel /
'  TooltipTableParser sunt Friend.
'
'  CUM RULEZI:
'    1. Adaugă acest fișier în proiect.
'    2. Apelează o singură dată TooltipTableParseTest.Run() dintr-un loc care
'       se execută la pornire în Debug - de ex. în handler-ul Load al formei
'       gazdă, sau dintr-un buton temporar:
'           Private Sub Form_Load(...) Handles MyBase.Load
'               TooltipTableParseTest.Run()
'           End Sub
'    3. Citește rezultatul din MessageBox-ul afișat.
'    4. Șterge fișierul + apelul după validare.
' ============================================================================

Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

Friend Module TooltipTableParseTest

    Friend Sub Run()
        Dim report As New StringBuilder()

        ' ── XML VALID (produse) ─────────────────────────────────────────────
        Dim xml As String =
            <table CellPaddingH="6" CellPaddingV="3" MaxWidth="360" HeaderBackColor="#2E75B6"><header><cell>Produs</cell><cell Align="Right">Cant.</cell><cell Align="Right">Valoare</cell></header><row><cell>Laptop Dell XPS</cell><cell>3</cell><cell>12.450,00</cell></row><row BackColor="#F2F2F2"><cell>Mouse Logitech</cell><cell>10</cell><cell>890,00</cell></row><footer><cell Bold="1">TOTAL</cell><cell></cell><cell Align="Right" Bold="1" ForeColor="#CC0000">13.340,00</cell></footer></table>.ToString()

        report.AppendLine("=== TEST 1: XML valid ===")
        report.AppendLine("IsTableXml = " & TooltipTableParser.IsTableXml(xml).ToString())

        Dim model As TooltipTableModel = Nothing
        Dim err As String = Nothing
        If TooltipTableParser.TryParse(xml, model, err) Then
            report.AppendLine("TryParse = True")
            report.AppendLine()
            report.Append(Dump(model))
        Else
            report.AppendLine("TryParse = False  ->  " & err)
        End If

        ' ── NEGATIV 1: XML malformat (lipsește </row>) ──────────────────────
        report.AppendLine()
        report.AppendLine("=== TEST 2: XML malformat ===")
        Dim badXml As String = "<table><row><cell>x</cell></table>"
        report.AppendLine("IsTableXml = " & TooltipTableParser.IsTableXml(badXml).ToString())
        Dim m2 As TooltipTableModel = Nothing
        Dim e2 As String = Nothing
        report.AppendLine("TryParse = " & TooltipTableParser.TryParse(badXml, m2, e2).ToString() &
                          "  ->  " & If(e2, "(fără mesaj)"))

        ' ── NEGATIV 2: nu e tabel (RichText normal) ─────────────────────────
        report.AppendLine()
        report.AppendLine("=== TEST 3: RichText normal (nu tabel) ===")
        Dim rich As String = "Salut <b>lume</b> <color=#FF0000>roșu</color>"
        report.AppendLine("IsTableXml = " & TooltipTableParser.IsTableXml(rich).ToString() &
                          "   (așteptat: False)")

        ' ── NEGATIV 3: prolog <?xml ?> tolerat ──────────────────────────────
        report.AppendLine()
        report.AppendLine("=== TEST 4: prolog <?xml ?> tolerat ===")
        Dim withProlog As String = "<?xml version=""1.0""?><table><row><cell>a</cell></row></table>"
        report.AppendLine("IsTableXml = " & TooltipTableParser.IsTableXml(withProlog).ToString() &
                          "   (așteptat: True)")
        Dim m4 As TooltipTableModel = Nothing
        Dim e4 As String = Nothing
        report.AppendLine("TryParse = " & TooltipTableParser.TryParse(withProlog, m4, e4).ToString())

        MessageBox.Show(report.ToString(), "TooltipTableParser - rezultate test",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ========================================================================
    '  DUMP MODEL
    ' ========================================================================

    Private Function Dump(m As TooltipTableModel) As String
        Dim sb As New StringBuilder()
        Dim cfg As TableConfig = m.Config

        sb.AppendLine("-- CONFIG --")
        sb.AppendLine($"FontName={cfg.FontName}  FontSize={cfg.FontSize}")
        sb.AppendLine($"PadH={cfg.CellPaddingH} PadV={cfg.CellPaddingV} RowHeight={cfg.RowHeight} MaxWidth={cfg.MaxWidth}")
        sb.AppendLine($"GridVisible={cfg.GridVisible} GridColor={FmtColor(cfg.GridColor)}")
        sb.AppendLine($"HdrBack={FmtColor(cfg.HeaderBackColor)} HdrFore={FmtColor(cfg.HeaderForeColor)}")
        sb.AppendLine($"FooterItalic={cfg.FooterItalic} FooterSep={cfg.FooterSeparator}")
        sb.AppendLine($"ColCount = {m.ColCount}")
        sb.AppendLine()

        If m.HeaderRow IsNot Nothing Then DumpRow(sb, "HEADER", m.HeaderRow)

        Dim i As Integer = 0
        For Each r As TtRow In m.Rows
            DumpRow(sb, $"ROW[{i}]", r)
            i += 1
        Next

        If m.FooterRow IsNot Nothing Then DumpRow(sb, "FOOTER", m.FooterRow)

        Return sb.ToString()
    End Function

    Private Sub DumpRow(sb As StringBuilder, label As String, r As TtRow)
        sb.AppendLine($"{label}  (BackColor={FmtColor(r.BackColor)}, cells={r.Cells.Count})")
        Dim j As Integer = 0
        For Each c As TtCell In r.Cells
            sb.AppendLine($"   [{j}] '{c.Text}'  W={c.Width} Align={c.Align} Bold={c.Bold} Italic={c.Italic} Fore={FmtColor(c.ForeColor)} Back={FmtColor(c.BackColor)}")
            j += 1
        Next
    End Sub

    Private Function FmtColor(c As Color) As String
        If c.IsEmpty Then Return "(empty)"
        Return $"#{c.R:X2}{c.G:X2}{c.B:X2}"
    End Function

End Module