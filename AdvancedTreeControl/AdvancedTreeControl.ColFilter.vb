Imports System.Linq

''' <summary>
''' Popup flotant pentru filtrarea pe o coloană specifică din TreeListView.
''' Conține un TextBox (Enter = filtrare) și un ListBox cu valori distincte.
''' Se închide automat la OnDeactivate.
''' </summary>
Partial Public Class AdvancedTreeControl
    Private NotInheritable Class ColFilterPopup
        Inherits Form

        Private ReadOnly _owner   As AdvancedTreeControl
        Private ReadOnly _colName As String
        Private _lblTitle  As Label
        Private _textBox   As TextBox
        Private _separator As Panel
        Private _listBox   As ListBox

        ' ────────────────────────────────────────────────────────────────
        Friend Sub New(owner As AdvancedTreeControl, colName As String, screenPos As Point)
            _owner   = owner
            _colName = colName

            ' ── Form ────────────────────────────────────────────────────
            Me.FormBorderStyle = FormBorderStyle.None
            Me.ShowInTaskbar   = False
            Me.TopMost         = True
            Me.StartPosition   = FormStartPosition.Manual
            Me.BackColor       = Color.FromArgb(250, 250, 252)
            Me.Width           = 230

            ' ── Title ───────────────────────────────────────────────────
            _lblTitle = New Label() With {
                .Text      = "  " & colName,
                .Font      = New Font(owner.Font, FontStyle.Bold),
                .BackColor = Color.FromArgb(228, 228, 244),
                .ForeColor = Color.FromArgb(40, 40, 80),
                .Height    = 24,
                .Width     = 230,
                .Location  = New Point(0, 0),
                .TextAlign = ContentAlignment.MiddleLeft
            }

            ' ── TextBox ─────────────────────────────────────────────────
            _textBox = New TextBox() With {
                .BorderStyle = BorderStyle.FixedSingle,
                .Font        = owner.Font,
                .Width       = 218,
                .Location    = New Point(6, 28)
            }
            ' Pre-populează cu filtrul activ (dacă există)
            If owner._activeColFilters.ContainsKey(colName) Then
                _textBox.Text = owner._activeColFilters(colName)
            End If

            ' ── Separator ───────────────────────────────────────────────
            Dim tbBottom As Integer = _textBox.Top + _textBox.PreferredHeight + 6
            _separator = New Panel() With {
                .BackColor = Color.FromArgb(200, 200, 215),
                .Height    = 1,
                .Width     = 230,
                .Location  = New Point(0, tbBottom)
            }

            ' ── ListBox ─────────────────────────────────────────────────
            _listBox = New ListBox() With {
                .BorderStyle    = BorderStyle.None,
                .Font           = owner.Font,
                .IntegralHeight = True,
                .Location       = New Point(0, tbBottom + 1),
                .Width          = 230
            }
            _listBox.Items.Add("(Toate)")
            For Each v In owner.GetDistinctColumnValues(colName)
                _listBox.Items.Add(v)
            Next
            ' Pre-selectează valoarea curentă (dacă există în listă)
            If owner._activeColFilters.ContainsKey(colName) Then
                Dim cur As String = owner._activeColFilters(colName)
                Dim idx As Integer = _listBox.Items.IndexOf(cur)
                If idx >= 0 Then _listBox.SelectedIndex = idx
            End If
            ' Înălțime: max 8 linii vizibile
            Dim visItems As Integer = Math.Min(_listBox.Items.Count, 8)
            _listBox.Height = _listBox.ItemHeight * visItems + 2

            ' ── Form height ─────────────────────────────────────────────
            Me.Height = _listBox.Top + _listBox.Height + 3

            ' ── Adaugă controalele ──────────────────────────────────────
            Me.Controls.AddRange(New Control() {_lblTitle, _textBox, _separator, _listBox})

            ' ── Poziționare — ajustare dacă iese din ecran ──────────────
            Me.Location = screenPos
            Dim scr As Rectangle = Screen.FromPoint(screenPos).WorkingArea
            If Me.Right  > scr.Right  Then Me.Left = scr.Right  - Me.Width
            If Me.Bottom > scr.Bottom Then Me.Top  = screenPos.Y - Me.Height

            ' ── Events ──────────────────────────────────────────────────
            AddHandler _textBox.KeyDown, AddressOf OnTextBoxKeyDown
            AddHandler _listBox.Click,   AddressOf OnListBoxClick
        End Sub

        ' ────────────────────────────────────────────────────────────────
        Protected Overrides Sub OnShown(e As EventArgs)
            MyBase.OnShown(e)
            _textBox.Focus()
            _textBox.SelectAll()
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            MyBase.OnPaint(e)
            Using pen As New Pen(Color.FromArgb(160, 160, 200), 1)
                e.Graphics.DrawRectangle(pen, 0, 0, Me.Width - 1, Me.Height - 1)
            End Using
        End Sub

        Protected Overrides Sub OnDeactivate(e As EventArgs)
            MyBase.OnDeactivate(e)
            Me.Close()
        End Sub

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing Then
                If _owner._activeColFilterPopup Is Me Then
                    _owner._activeColFilterPopup = Nothing
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        ' ────────────────────────────────────────────────────────────────
        Private Sub OnTextBoxKeyDown(sender As Object, e As KeyEventArgs)
            Select Case e.KeyCode
                Case Keys.Return
                    e.SuppressKeyPress = True
                    ApplyFilter(_textBox.Text.Trim())
                    Me.Close()
                Case Keys.Escape
                    Me.Close()
            End Select
        End Sub

        Private Sub OnListBoxClick(sender As Object, e As EventArgs)
            If _listBox.SelectedIndex < 0 Then Return
            Dim selected As String = _listBox.SelectedItem.ToString()
            ApplyFilter(If(selected = "(Toate)", "", selected))
            Me.Close()
        End Sub

        Private Sub ApplyFilter(text As String)
            If String.IsNullOrEmpty(text) Then
                _owner._activeColFilters.Remove(_colName)
            Else
                _owner._activeColFilters(_colName) = text
            End If
            _owner.ApplyColumnFilters()
        End Sub

    End Class  ' ColFilterPopup
End Class  ' AdvancedTreeControl (partial)
