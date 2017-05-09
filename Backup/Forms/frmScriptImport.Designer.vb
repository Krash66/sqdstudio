<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScriptImport
    Inherits SQDStudio.frmBlank

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmScriptImport))
        Me.gbInputScript = New System.Windows.Forms.GroupBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtScriptPath = New System.Windows.Forms.TextBox
        Me.lvInScript = New System.Windows.Forms.ListView
        Me.colLineNo = New System.Windows.Forms.ColumnHeader
        Me.colScriptText = New System.Windows.Forms.ColumnHeader
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.NewToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.OpenToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.SaveToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.PrintToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.toolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator
        Me.HelpToolStripButton = New System.Windows.Forms.ToolStripButton
        Me.gbOption = New System.Windows.Forms.GroupBox
        Me.btnAnalyze = New System.Windows.Forms.Button
        Me.gbScriptTree = New System.Windows.Forms.GroupBox
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.gbProperties = New System.Windows.Forms.GroupBox
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.colProp = New System.Windows.Forms.ColumnHeader
        Me.colValue = New System.Windows.Forms.ColumnHeader
        Me.scScriptTree = New System.Windows.Forms.SplitContainer
        Me.OFD1 = New System.Windows.Forms.OpenFileDialog
        Me.gbAnalyzedScript = New System.Windows.Forms.GroupBox
        Me.ListView2 = New System.Windows.Forms.ListView
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Panel1.SuspendLayout()
        Me.gbInputScript.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.gbOption.SuspendLayout()
        Me.gbScriptTree.SuspendLayout()
        Me.gbProperties.SuspendLayout()
        Me.scScriptTree.Panel1.SuspendLayout()
        Me.scScriptTree.Panel2.SuspendLayout()
        Me.scScriptTree.SuspendLayout()
        Me.gbAnalyzedScript.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Location = New System.Drawing.Point(-1, -1)
        Me.Panel1.Size = New System.Drawing.Size(453, 66)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(0, 506)
        Me.GroupBox1.Size = New System.Drawing.Size(841, 10)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(571, 518)
        Me.cmdOk.Size = New System.Drawing.Size(85, 24)
        Me.cmdOk.Text = "&Import"
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(662, 518)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(748, 518)
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.Size = New System.Drawing.Size(375, 16)
        Me.Label1.Text = "SQData Script Importer"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(377, 39)
        Me.Label2.Text = "This tool allows SQData Scripts modified or created outside of Design Studio to b" & _
            "e imported as a full or partial SQData Project within Design Studio."
        '
        'gbInputScript
        '
        Me.gbInputScript.Controls.Add(Me.Label3)
        Me.gbInputScript.Controls.Add(Me.txtScriptPath)
        Me.gbInputScript.Controls.Add(Me.lvInScript)
        Me.gbInputScript.Controls.Add(Me.ToolStrip1)
        Me.gbInputScript.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbInputScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbInputScript.Location = New System.Drawing.Point(0, 0)
        Me.gbInputScript.Name = "gbInputScript"
        Me.gbInputScript.Size = New System.Drawing.Size(448, 214)
        Me.gbInputScript.TabIndex = 58
        Me.gbInputScript.TabStop = False
        Me.gbInputScript.Text = "Script to Import"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(3, 47)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(74, 13)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Script Path:"
        '
        'txtScriptPath
        '
        Me.txtScriptPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtScriptPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtScriptPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtScriptPath.Location = New System.Drawing.Point(83, 44)
        Me.txtScriptPath.Name = "txtScriptPath"
        Me.txtScriptPath.ReadOnly = True
        Me.txtScriptPath.Size = New System.Drawing.Size(359, 20)
        Me.txtScriptPath.TabIndex = 3
        '
        'lvInScript
        '
        Me.lvInScript.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvInScript.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colLineNo, Me.colScriptText})
        Me.lvInScript.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvInScript.GridLines = True
        Me.lvInScript.HideSelection = False
        Me.lvInScript.Location = New System.Drawing.Point(3, 70)
        Me.lvInScript.Name = "lvInScript"
        Me.lvInScript.Size = New System.Drawing.Size(442, 138)
        Me.lvInScript.TabIndex = 2
        Me.lvInScript.UseCompatibleStateImageBehavior = False
        Me.lvInScript.View = System.Windows.Forms.View.Details
        '
        'colLineNo
        '
        Me.colLineNo.Text = "Line#"
        Me.colLineNo.Width = 55
        '
        'colScriptText
        '
        Me.colScriptText.Text = "Script Text"
        Me.colScriptText.Width = 650
        '
        'ToolStrip1
        '
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripButton, Me.OpenToolStripButton, Me.SaveToolStripButton, Me.PrintToolStripButton, Me.toolStripSeparator1, Me.HelpToolStripButton})
        Me.ToolStrip1.Location = New System.Drawing.Point(3, 16)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(442, 25)
        Me.ToolStrip1.TabIndex = 0
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'NewToolStripButton
        '
        Me.NewToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.NewToolStripButton.Enabled = False
        Me.NewToolStripButton.Image = CType(resources.GetObject("NewToolStripButton.Image"), System.Drawing.Image)
        Me.NewToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.NewToolStripButton.Name = "NewToolStripButton"
        Me.NewToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.NewToolStripButton.Text = "&New"
        Me.NewToolStripButton.Visible = False
        '
        'OpenToolStripButton
        '
        Me.OpenToolStripButton.Image = CType(resources.GetObject("OpenToolStripButton.Image"), System.Drawing.Image)
        Me.OpenToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.OpenToolStripButton.Name = "OpenToolStripButton"
        Me.OpenToolStripButton.Size = New System.Drawing.Size(83, 22)
        Me.OpenToolStripButton.Text = "&Open Script"
        '
        'SaveToolStripButton
        '
        Me.SaveToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.SaveToolStripButton.Enabled = False
        Me.SaveToolStripButton.Image = CType(resources.GetObject("SaveToolStripButton.Image"), System.Drawing.Image)
        Me.SaveToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.SaveToolStripButton.Name = "SaveToolStripButton"
        Me.SaveToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.SaveToolStripButton.Text = "&Save"
        Me.SaveToolStripButton.Visible = False
        '
        'PrintToolStripButton
        '
        Me.PrintToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.PrintToolStripButton.Image = CType(resources.GetObject("PrintToolStripButton.Image"), System.Drawing.Image)
        Me.PrintToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.PrintToolStripButton.Name = "PrintToolStripButton"
        Me.PrintToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.PrintToolStripButton.Text = "&Print"
        '
        'toolStripSeparator1
        '
        Me.toolStripSeparator1.Name = "toolStripSeparator1"
        Me.toolStripSeparator1.Size = New System.Drawing.Size(6, 25)
        '
        'HelpToolStripButton
        '
        Me.HelpToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.HelpToolStripButton.Enabled = False
        Me.HelpToolStripButton.Image = CType(resources.GetObject("HelpToolStripButton.Image"), System.Drawing.Image)
        Me.HelpToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.HelpToolStripButton.Name = "HelpToolStripButton"
        Me.HelpToolStripButton.Size = New System.Drawing.Size(23, 22)
        Me.HelpToolStripButton.Text = "He&lp"
        Me.HelpToolStripButton.Visible = False
        '
        'gbOption
        '
        Me.gbOption.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbOption.Controls.Add(Me.btnAnalyze)
        Me.gbOption.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbOption.Location = New System.Drawing.Point(462, 3)
        Me.gbOption.Name = "gbOption"
        Me.gbOption.Size = New System.Drawing.Size(366, 85)
        Me.gbOption.TabIndex = 59
        Me.gbOption.TabStop = False
        Me.gbOption.Text = "Options"
        '
        'btnAnalyze
        '
        Me.btnAnalyze.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAnalyze.Location = New System.Drawing.Point(285, 9)
        Me.btnAnalyze.Name = "btnAnalyze"
        Me.btnAnalyze.Size = New System.Drawing.Size(75, 70)
        Me.btnAnalyze.TabIndex = 0
        Me.btnAnalyze.Text = "Analyze Script"
        Me.btnAnalyze.UseVisualStyleBackColor = True
        '
        'gbScriptTree
        '
        Me.gbScriptTree.BackColor = System.Drawing.SystemColors.Control
        Me.gbScriptTree.Controls.Add(Me.TreeView1)
        Me.gbScriptTree.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbScriptTree.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbScriptTree.Location = New System.Drawing.Point(0, 0)
        Me.gbScriptTree.Name = "gbScriptTree"
        Me.gbScriptTree.Size = New System.Drawing.Size(366, 203)
        Me.gbScriptTree.TabIndex = 60
        Me.gbScriptTree.TabStop = False
        Me.gbScriptTree.Text = "Script Tree"
        '
        'TreeView1
        '
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeView1.Location = New System.Drawing.Point(3, 16)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.Size = New System.Drawing.Size(360, 184)
        Me.TreeView1.TabIndex = 0
        '
        'gbProperties
        '
        Me.gbProperties.BackColor = System.Drawing.SystemColors.Control
        Me.gbProperties.Controls.Add(Me.ListView1)
        Me.gbProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProperties.Location = New System.Drawing.Point(0, 0)
        Me.gbProperties.Name = "gbProperties"
        Me.gbProperties.Size = New System.Drawing.Size(366, 197)
        Me.gbProperties.TabIndex = 61
        Me.gbProperties.TabStop = False
        Me.gbProperties.Text = "Properties"
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colProp, Me.colValue})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(3, 16)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(360, 178)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'colProp
        '
        Me.colProp.Text = "Property"
        Me.colProp.Width = 110
        '
        'colValue
        '
        Me.colValue.Text = "Current Value"
        Me.colValue.Width = 222
        '
        'scScriptTree
        '
        Me.scScriptTree.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scScriptTree.Location = New System.Drawing.Point(462, 94)
        Me.scScriptTree.Name = "scScriptTree"
        Me.scScriptTree.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scScriptTree.Panel1
        '
        Me.scScriptTree.Panel1.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.scScriptTree.Panel1.Controls.Add(Me.gbScriptTree)
        '
        'scScriptTree.Panel2
        '
        Me.scScriptTree.Panel2.BackColor = System.Drawing.SystemColors.InactiveCaptionText
        Me.scScriptTree.Panel2.Controls.Add(Me.gbProperties)
        Me.scScriptTree.Size = New System.Drawing.Size(366, 406)
        Me.scScriptTree.SplitterDistance = 203
        Me.scScriptTree.SplitterWidth = 6
        Me.scScriptTree.TabIndex = 62
        '
        'OFD1
        '
        Me.OFD1.DefaultExt = "inl"
        Me.OFD1.Filter = "SQData Scripts|*.inl|Parsed Scripts|*.rpt|All Files|*.*"
        Me.OFD1.Title = "Open SQData Script for Import"
        '
        'gbAnalyzedScript
        '
        Me.gbAnalyzedScript.Controls.Add(Me.ListView2)
        Me.gbAnalyzedScript.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAnalyzedScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbAnalyzedScript.Location = New System.Drawing.Point(0, 0)
        Me.gbAnalyzedScript.Name = "gbAnalyzedScript"
        Me.gbAnalyzedScript.Size = New System.Drawing.Size(448, 210)
        Me.gbAnalyzedScript.TabIndex = 63
        Me.gbAnalyzedScript.TabStop = False
        Me.gbAnalyzedScript.Text = "Analyzed Script"
        '
        'ListView2
        '
        Me.ListView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView2.Location = New System.Drawing.Point(3, 16)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(442, 191)
        Me.ListView2.TabIndex = 0
        Me.ListView2.UseCompatibleStateImageBehavior = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SplitContainer1.Location = New System.Drawing.Point(6, 72)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbInputScript)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gbAnalyzedScript)
        Me.SplitContainer1.Size = New System.Drawing.Size(448, 428)
        Me.SplitContainer1.SplitterDistance = 214
        Me.SplitContainer1.TabIndex = 64
        '
        'frmScriptImport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(840, 554)
        Me.Controls.Add(Me.gbOption)
        Me.Controls.Add(Me.scScriptTree)
        Me.Controls.Add(Me.SplitContainer1)
        Me.MinimumSize = New System.Drawing.Size(562, 291)
        Me.Name = "frmScriptImport"
        Me.Text = "Script Importer"
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.SplitContainer1, 0)
        Me.Controls.SetChildIndex(Me.scScriptTree, 0)
        Me.Controls.SetChildIndex(Me.gbOption, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbInputScript.ResumeLayout(False)
        Me.gbInputScript.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.gbOption.ResumeLayout(False)
        Me.gbScriptTree.ResumeLayout(False)
        Me.gbProperties.ResumeLayout(False)
        Me.scScriptTree.Panel1.ResumeLayout(False)
        Me.scScriptTree.Panel2.ResumeLayout(False)
        Me.scScriptTree.ResumeLayout(False)
        Me.gbAnalyzedScript.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbInputScript As System.Windows.Forms.GroupBox
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents NewToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents OpenToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents SaveToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents PrintToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents toolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents HelpToolStripButton As System.Windows.Forms.ToolStripButton
    Friend WithEvents gbOption As System.Windows.Forms.GroupBox
    Friend WithEvents gbScriptTree As System.Windows.Forms.GroupBox
    Friend WithEvents gbProperties As System.Windows.Forms.GroupBox
    Friend WithEvents scScriptTree As System.Windows.Forms.SplitContainer
    Friend WithEvents colLineNo As System.Windows.Forms.ColumnHeader
    Friend WithEvents colScriptText As System.Windows.Forms.ColumnHeader
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents colProp As System.Windows.Forms.ColumnHeader
    Friend WithEvents colValue As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnAnalyze As System.Windows.Forms.Button
    Friend WithEvents OFD1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtScriptPath As System.Windows.Forms.TextBox
    Friend WithEvents gbAnalyzedScript As System.Windows.Forms.GroupBox
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Private WithEvents lvInScript As System.Windows.Forms.ListView
End Class
