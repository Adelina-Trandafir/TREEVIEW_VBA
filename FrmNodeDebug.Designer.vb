<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Public Class frmNodeDebug
    Inherits System.Windows.Forms.Form

    ' Necesar pentru suportul Windows Forms Designer
    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private Sub InitializeComponent()
        _pnlTop = New Panel()
        _lblTitle = New Label()
        _pg = New PropertyGrid()
        _pnlBottom = New Panel()
        _btnClose = New Button()
        _btnCopy = New Button()
        _pnlTop.SuspendLayout()
        _pnlBottom.SuspendLayout()
        SuspendLayout()
        ' 
        ' _pnlTop
        ' 
        _pnlTop.BackColor = Color.FromArgb(CByte(16), CByte(16), CByte(16))
        _pnlTop.Controls.Add(_lblTitle)
        _pnlTop.Dock = DockStyle.Top
        _pnlTop.Location = New Point(0, 0)
        _pnlTop.Margin = New Padding(4, 5, 4, 5)
        _pnlTop.Name = "_pnlTop"
        _pnlTop.Size = New Size(638, 57)
        _pnlTop.TabIndex = 0
        ' 
        ' _lblTitle
        ' 
        _lblTitle.Dock = DockStyle.Fill
        _lblTitle.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        _lblTitle.ForeColor = Color.FromArgb(CByte(90), CByte(170), CByte(255))
        _lblTitle.Location = New Point(0, 0)
        _lblTitle.Margin = New Padding(4, 0, 4, 0)
        _lblTitle.Name = "_lblTitle"
        _lblTitle.Padding = New Padding(11, 0, 0, 0)
        _lblTitle.Size = New Size(638, 57)
        _lblTitle.TabIndex = 0
        _lblTitle.Text = "Node Inspector"
        _lblTitle.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' _pg
        ' 
        _pg.BackColor = Color.FromArgb(CByte(24), CByte(24), CByte(24))
        _pg.CategoryForeColor = Color.FromArgb(CByte(90), CByte(170), CByte(255))
        _pg.DisabledItemForeColor = Color.FromArgb(CByte(127), CByte(220), CByte(220), CByte(220))
        _pg.Dock = DockStyle.Fill
        _pg.HelpBackColor = Color.FromArgb(CByte(16), CByte(16), CByte(16))
        _pg.HelpForeColor = Color.FromArgb(CByte(160), CByte(160), CByte(160))
        _pg.LineColor = Color.FromArgb(CByte(50), CByte(50), CByte(50))
        _pg.Location = New Point(0, 57)
        _pg.Margin = New Padding(4, 5, 4, 5)
        _pg.Name = "_pg"
        _pg.PropertySort = PropertySort.Categorized
        _pg.SelectedItemWithFocusBackColor = Color.FromArgb(CByte(0), CByte(78), CByte(140))
        _pg.SelectedItemWithFocusForeColor = Color.White
        _pg.Size = New Size(638, 509)
        _pg.TabIndex = 1
        _pg.ToolbarVisible = False
        _pg.ViewBackColor = Color.FromArgb(CByte(34), CByte(34), CByte(34))
        _pg.ViewForeColor = Color.FromArgb(CByte(220), CByte(220), CByte(220))
        ' 
        ' _pnlBottom
        ' 
        _pnlBottom.BackColor = Color.FromArgb(CByte(16), CByte(16), CByte(16))
        _pnlBottom.Controls.Add(_btnClose)
        _pnlBottom.Controls.Add(_btnCopy)
        _pnlBottom.Dock = DockStyle.Bottom
        _pnlBottom.Location = New Point(0, 566)
        _pnlBottom.Margin = New Padding(4, 5, 4, 5)
        _pnlBottom.Name = "_pnlBottom"
        _pnlBottom.Size = New Size(638, 70)
        _pnlBottom.TabIndex = 2
        ' 
        ' _btnClose
        ' 
        _btnClose.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        _btnClose.BackColor = Color.FromArgb(CByte(75), CByte(25), CByte(25))
        _btnClose.Cursor = Cursors.Hand
        _btnClose.FlatAppearance.BorderColor = Color.FromArgb(CByte(110), CByte(50), CByte(50))
        _btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(CByte(100), CByte(35), CByte(35))
        _btnClose.FlatStyle = FlatStyle.Flat
        _btnClose.ForeColor = Color.FromArgb(CByte(230), CByte(170), CByte(170))
        _btnClose.Location = New Point(513, 12)
        _btnClose.Margin = New Padding(4, 5, 4, 5)
        _btnClose.Name = "_btnClose"
        _btnClose.Size = New Size(114, 47)
        _btnClose.TabIndex = 1
        _btnClose.Text = "Inchide"
        _btnClose.UseVisualStyleBackColor = False
        ' 
        ' _btnCopy
        ' 
        _btnCopy.BackColor = Color.FromArgb(CByte(45), CByte(45), CByte(55))
        _btnCopy.Cursor = Cursors.Hand
        _btnCopy.FlatAppearance.BorderColor = Color.FromArgb(CByte(75), CByte(75), CByte(95))
        _btnCopy.FlatAppearance.MouseOverBackColor = Color.FromArgb(CByte(60), CByte(60), CByte(80))
        _btnCopy.FlatStyle = FlatStyle.Flat
        _btnCopy.ForeColor = Color.FromArgb(CByte(210), CByte(210), CByte(210))
        _btnCopy.Location = New Point(11, 12)
        _btnCopy.Margin = New Padding(4, 5, 4, 5)
        _btnCopy.Name = "_btnCopy"
        _btnCopy.Size = New Size(186, 47)
        _btnCopy.TabIndex = 0
        _btnCopy.Text = "📋  Copiaza tot"
        _btnCopy.UseVisualStyleBackColor = False
        ' 
        ' frmNodeDebug
        ' 
        AutoScaleDimensions = New SizeF(10F, 25F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(24), CByte(24), CByte(24))
        ClientSize = New Size(638, 636)
        Controls.Add(_pg)
        Controls.Add(_pnlBottom)
        Controls.Add(_pnlTop)
        Font = New Font("Segoe UI", 9F)
        ForeColor = Color.FromArgb(CByte(220), CByte(220), CByte(220))
        FormBorderStyle = FormBorderStyle.SizableToolWindow
        Margin = New Padding(4, 5, 4, 5)
        MinimumSize = New Size(533, 629)
        Name = "frmNodeDebug"
        ShowInTaskbar = False
        StartPosition = FormStartPosition.Manual
        Text = "Node Inspector"
        TopMost = True
        _pnlTop.ResumeLayout(False)
        _pnlBottom.ResumeLayout(False)
        ResumeLayout(False)
    End Sub

    ' ── Declaratii controale ────────────────────────────────────────────────────
    Friend WithEvents _pg As System.Windows.Forms.PropertyGrid
    Friend WithEvents _btnClose As System.Windows.Forms.Button
    Friend WithEvents _btnCopy As System.Windows.Forms.Button
    Friend _lblTitle As System.Windows.Forms.Label
    Friend _pnlTop As System.Windows.Forms.Panel
    Friend _pnlBottom As System.Windows.Forms.Panel

End Class