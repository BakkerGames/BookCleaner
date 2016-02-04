<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBoxFromPath = New System.Windows.Forms.TextBox()
        Me.ButtonStart = New System.Windows.Forms.Button()
        Me.TextBoxFileName = New System.Windows.Forms.TextBox()
        Me.MenuStripMain = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TextBoxResults = New System.Windows.Forms.TextBox()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabelMain = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ButtonStop = New System.Windows.Forms.Button()
        Me.MenuStripMain.SuspendLayout()
        Me.StatusStripMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBoxFromPath
        '
        Me.TextBoxFromPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxFromPath.Location = New System.Drawing.Point(12, 27)
        Me.TextBoxFromPath.Name = "TextBoxFromPath"
        Me.TextBoxFromPath.Size = New System.Drawing.Size(536, 20)
        Me.TextBoxFromPath.TabIndex = 1
        '
        'ButtonStart
        '
        Me.ButtonStart.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonStart.Location = New System.Drawing.Point(243, 53)
        Me.ButtonStart.Name = "ButtonStart"
        Me.ButtonStart.Size = New System.Drawing.Size(75, 23)
        Me.ButtonStart.TabIndex = 2
        Me.ButtonStart.Text = "Start"
        Me.ButtonStart.UseVisualStyleBackColor = True
        '
        'TextBoxFileName
        '
        Me.TextBoxFileName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBoxFileName.Location = New System.Drawing.Point(12, 82)
        Me.TextBoxFileName.Name = "TextBoxFileName"
        Me.TextBoxFileName.ReadOnly = True
        Me.TextBoxFileName.Size = New System.Drawing.Size(536, 20)
        Me.TextBoxFileName.TabIndex = 4
        '
        'MenuStripMain
        '
        Me.MenuStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStripMain.Location = New System.Drawing.Point(0, 0)
        Me.MenuStripMain.Name = "MenuStripMain"
        Me.MenuStripMain.Size = New System.Drawing.Size(560, 24)
        Me.MenuStripMain.TabIndex = 0
        Me.MenuStripMain.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'TextBoxResults
        '
        Me.TextBoxResults.Location = New System.Drawing.Point(12, 108)
        Me.TextBoxResults.Multiline = True
        Me.TextBoxResults.Name = "TextBoxResults"
        Me.TextBoxResults.ReadOnly = True
        Me.TextBoxResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBoxResults.Size = New System.Drawing.Size(536, 229)
        Me.TextBoxResults.TabIndex = 5
        '
        'StatusStripMain
        '
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabelMain})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 345)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Padding = New System.Windows.Forms.Padding(1, 0, 10, 0)
        Me.StatusStripMain.Size = New System.Drawing.Size(560, 22)
        Me.StatusStripMain.SizingGrip = False
        Me.StatusStripMain.TabIndex = 6
        '
        'ToolStripStatusLabelMain
        '
        Me.ToolStripStatusLabelMain.Name = "ToolStripStatusLabelMain"
        Me.ToolStripStatusLabelMain.Size = New System.Drawing.Size(549, 17)
        Me.ToolStripStatusLabelMain.Spring = True
        Me.ToolStripStatusLabelMain.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ButtonStop
        '
        Me.ButtonStop.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.ButtonStop.Enabled = False
        Me.ButtonStop.Location = New System.Drawing.Point(243, 53)
        Me.ButtonStop.Name = "ButtonStop"
        Me.ButtonStop.Size = New System.Drawing.Size(75, 23)
        Me.ButtonStop.TabIndex = 3
        Me.ButtonStop.Text = "Stop"
        Me.ButtonStop.UseVisualStyleBackColor = True
        Me.ButtonStop.Visible = False
        '
        'FormMain
        '
        Me.AcceptButton = Me.ButtonStart
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(560, 367)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.TextBoxResults)
        Me.Controls.Add(Me.TextBoxFileName)
        Me.Controls.Add(Me.ButtonStart)
        Me.Controls.Add(Me.TextBoxFromPath)
        Me.Controls.Add(Me.MenuStripMain)
        Me.Controls.Add(Me.ButtonStop)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FormMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Book Cleaner"
        Me.MenuStripMain.ResumeLayout(False)
        Me.MenuStripMain.PerformLayout()
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBoxFromPath As System.Windows.Forms.TextBox
    Friend WithEvents ButtonStart As System.Windows.Forms.Button
    Friend WithEvents TextBoxFileName As System.Windows.Forms.TextBox
    Friend WithEvents MenuStripMain As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TextBoxResults As System.Windows.Forms.TextBox
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StatusStripMain As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripStatusLabelMain As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ButtonStop As System.Windows.Forms.Button

End Class
