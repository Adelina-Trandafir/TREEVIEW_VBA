Imports System.ComponentModel
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms

' ==============================================================================
' frmNodeDebug.vb  (logica)
' frmNodeDebug.Designer.vb  (layout — editabil vizual in VS Designer)
' ==============================================================================

Public Class frmNodeDebug

    ' ── Singleton ───────────────────────────────────────────────────────────────
    Private Shared _instance As frmNodeDebug
    Private WithEvents _resetTimer As New System.Windows.Forms.Timer() With {.Interval = 1500}

    Public Shared Sub ShowForNode(info As NodeDebugInfo,
                                  Optional owner As IWin32Window = Nothing)
        If _instance Is Nothing OrElse _instance.IsDisposed Then
            _instance = New frmNodeDebug()
        End If

        _instance.LoadInfo(info)

        If Not _instance.Visible Then
            Dim cur = Cursor.Position
            _instance.Location = New Point(cur.X + 24, Math.Max(0, cur.Y - 60))
            _instance.Show(owner)
        Else
            _instance.BringToFront()
        End If
    End Sub

    ' ── Load ────────────────────────────────────────────────────────────────────
    Public Sub LoadInfo(info As NodeDebugInfo)
        Dim childStr As String = If(info.ChildCount > 0, $"  |  {info.ChildCount} copii", "")
        Dim selStr As String = If(info.IsSelectedNode, "  ★ SELECTED", "")
        _lblTitle.Text = $"  [{info.Key}]   L{info.Level}{childStr}{selStr}"
        _pg.SelectedObject = info
        _pg.ExpandAllGridItems()
    End Sub

    ' ── Handlers ────────────────────────────────────────────────────────────────
    Private Sub _btnClose_Click(s As Object, e As EventArgs) Handles _btnClose.Click
        Me.Hide()
    End Sub

    Private Sub _btnCopy_Click(s As Object, e As EventArgs) Handles _btnCopy.Click
        Dim info = TryCast(_pg.SelectedObject, NodeDebugInfo)
        If info Is Nothing Then Return

        Clipboard.SetText(BuildReport(info))

        _btnCopy.Text = "✓  Copiat!"
        _btnCopy.BackColor = Color.FromArgb(15, 75, 15)
        _resetTimer.Start()
    End Sub

    Private Sub _resetTimer_Tick(s As Object, e As EventArgs) Handles _resetTimer.Tick
        _resetTimer.Stop()
        _btnCopy.Text = "📋  Copiaza tot"
        _btnCopy.BackColor = Color.FromArgb(45, 45, 55)
    End Sub

    Private Sub frmNodeDebug_Resize(s As Object, e As EventArgs) Handles Me.Resize
        If _btnClose IsNot Nothing AndAlso _pnlBottom IsNot Nothing Then
            _btnClose.Location = New Point(_pnlBottom.ClientSize.Width - _btnClose.Width - 8, 7)
        End If
    End Sub

    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Me.Hide()
        End If
        MyBase.OnFormClosing(e)
    End Sub

    ' ── Report clipboard ────────────────────────────────────────────────────────
    Private Shared Function BuildReport(info As NodeDebugInfo) As String
        Dim sb As New StringBuilder()
        sb.AppendLine(New String("="c, 62))
        sb.AppendLine($"  NODE INSPECTOR  ·  [{info.Key}]")
        sb.AppendLine($"  {DateTime.Now:yyyy-MM-dd  HH:mm:ss}")
        sb.AppendLine(New String("="c, 62))

        Dim lastCat = ""
        Dim props = TypeDescriptor.GetProperties(info) _
                        .Cast(Of PropertyDescriptor)() _
                        .OrderBy(Function(p) p.Category) _
                        .ThenBy(Function(p) p.DisplayName) _
                        .ToList()

        For Each prop In props
            If prop.Category <> lastCat Then
                sb.AppendLine()
                sb.AppendLine($"  {New String("─"c, 58)}")
                sb.AppendLine($"  {prop.Category}")
                sb.AppendLine($"  {New String("─"c, 58)}")
                lastCat = prop.Category
            End If

            Dim val = prop.GetValue(info)
            Dim valStr As String

            If TypeOf val Is Rectangle Then
                Dim r = CType(val, Rectangle)
                valStr = If(r = Rectangle.Empty, "Empty",
                            $"X={r.X}  Y={r.Y}  W={r.Width}  H={r.Height}")
            Else
                valStr = If(val Is Nothing, "(null)", val.ToString())
            End If

            sb.AppendLine($"    {prop.DisplayName,-34}  {valStr}")
        Next

        sb.AppendLine()
        Return sb.ToString()
    End Function

End Class