<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDMLInfo
    Inherits SQDStudio.frmBlank

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.BottomToolStripPanel = New System.Windows.Forms.ToolStripPanel()
        Me.TopToolStripPanel = New System.Windows.Forms.ToolStripPanel()
        Me.RightToolStripPanel = New System.Windows.Forms.ToolStripPanel()
        Me.LeftToolStripPanel = New System.Windows.Forms.ToolStripPanel()
        Me.ContentPanel = New System.Windows.Forms.ToolStripContentPanel()
        Me.TabSelectTable = New System.Windows.Forms.TabPage()
        Me.gbChseConn = New System.Windows.Forms.GroupBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.cbConn = New System.Windows.Forms.ComboBox()
        Me.gbTableList = New System.Windows.Forms.GroupBox()
        Me.dgvTables = New System.Windows.Forms.DataGridView()
        Me.gbNarrowSearch = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.txtSchemaTab2 = New System.Windows.Forms.TextBox()
        Me.txtPartTableName = New System.Windows.Forms.TextBox()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.btnRemoveTab = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        Me.TabSelectTable.SuspendLayout()
        Me.gbChseConn.SuspendLayout()
        Me.gbTableList.SuspendLayout()
        CType(Me.dgvTables, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbNarrowSearch.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Size = New System.Drawing.Size(800, 84)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 740)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox1.Size = New System.Drawing.Size(803, 9)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(424, 768)
        Me.cmdOk.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdOk.Size = New System.Drawing.Size(107, 30)
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(552, 768)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCancel.Size = New System.Drawing.Size(107, 30)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(680, 768)
        Me.cmdHelp.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdHelp.Size = New System.Drawing.Size(107, 30)
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(96, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Size = New System.Drawing.Size(576, 20)
        Me.Label1.Text = "Choose Table Properties for new DML Structure"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(96, 30)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Size = New System.Drawing.Size(699, 48)
        Me.Label2.Text = "Select ODBC Source, Table Name, and Columns from an existing Oracle or DB2 Databa" & _
    "se. These will be used to Model a custom Structure."
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.Location = New System.Drawing.Point(21, 767)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(4)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(217, 33)
        Me.btnNext.TabIndex = 36
        Me.btnNext.Text = "&Next >>>"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'BottomToolStripPanel
        '
        Me.BottomToolStripPanel.Location = New System.Drawing.Point(0, 0)
        Me.BottomToolStripPanel.Name = "BottomToolStripPanel"
        Me.BottomToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.BottomToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.BottomToolStripPanel.Size = New System.Drawing.Size(0, 0)
        '
        'TopToolStripPanel
        '
        Me.TopToolStripPanel.Location = New System.Drawing.Point(0, 0)
        Me.TopToolStripPanel.Name = "TopToolStripPanel"
        Me.TopToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.TopToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.TopToolStripPanel.Size = New System.Drawing.Size(0, 0)
        '
        'RightToolStripPanel
        '
        Me.RightToolStripPanel.Location = New System.Drawing.Point(0, 0)
        Me.RightToolStripPanel.Name = "RightToolStripPanel"
        Me.RightToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.RightToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.RightToolStripPanel.Size = New System.Drawing.Size(0, 0)
        '
        'LeftToolStripPanel
        '
        Me.LeftToolStripPanel.Location = New System.Drawing.Point(0, 0)
        Me.LeftToolStripPanel.Name = "LeftToolStripPanel"
        Me.LeftToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal
        Me.LeftToolStripPanel.RowMargin = New System.Windows.Forms.Padding(3, 0, 0, 0)
        Me.LeftToolStripPanel.Size = New System.Drawing.Size(0, 0)
        '
        'ContentPanel
        '
        Me.ContentPanel.Size = New System.Drawing.Size(6, 148)
        '
        'TabSelectTable
        '
        Me.TabSelectTable.BackColor = System.Drawing.SystemColors.Control
        Me.TabSelectTable.Controls.Add(Me.gbChseConn)
        Me.TabSelectTable.Controls.Add(Me.gbTableList)
        Me.TabSelectTable.Controls.Add(Me.gbNarrowSearch)
        Me.TabSelectTable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabSelectTable.Location = New System.Drawing.Point(4, 26)
        Me.TabSelectTable.Margin = New System.Windows.Forms.Padding(4)
        Me.TabSelectTable.Name = "TabSelectTable"
        Me.TabSelectTable.Padding = New System.Windows.Forms.Padding(4)
        Me.TabSelectTable.Size = New System.Drawing.Size(760, 612)
        Me.TabSelectTable.TabIndex = 1
        Me.TabSelectTable.Text = "Select Tables"
        '
        'gbChseConn
        '
        Me.gbChseConn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbChseConn.Controls.Add(Me.Label5)
        Me.gbChseConn.Controls.Add(Me.Label4)
        Me.gbChseConn.Controls.Add(Me.cbConn)
        Me.gbChseConn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbChseConn.Location = New System.Drawing.Point(8, 7)
        Me.gbChseConn.Margin = New System.Windows.Forms.Padding(4)
        Me.gbChseConn.Name = "gbChseConn"
        Me.gbChseConn.Padding = New System.Windows.Forms.Padding(4)
        Me.gbChseConn.Size = New System.Drawing.Size(741, 80)
        Me.gbChseConn.TabIndex = 80
        Me.gbChseConn.TabStop = False
        Me.gbChseConn.Text = "Choose a Connection"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label5.Location = New System.Drawing.Point(8, 55)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(303, 17)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "This will only support ODBC Connections"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(8, 27)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(213, 17)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Select Database Connection"
        '
        'cbConn
        '
        Me.cbConn.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cbConn.FormattingEnabled = True
        Me.cbConn.Location = New System.Drawing.Point(348, 23)
        Me.cbConn.Margin = New System.Windows.Forms.Padding(4)
        Me.cbConn.Name = "cbConn"
        Me.cbConn.Size = New System.Drawing.Size(384, 25)
        Me.cbConn.TabIndex = 0
        '
        'gbTableList
        '
        Me.gbTableList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbTableList.Controls.Add(Me.dgvTables)
        Me.gbTableList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbTableList.Location = New System.Drawing.Point(8, 235)
        Me.gbTableList.Margin = New System.Windows.Forms.Padding(4)
        Me.gbTableList.Name = "gbTableList"
        Me.gbTableList.Padding = New System.Windows.Forms.Padding(4)
        Me.gbTableList.Size = New System.Drawing.Size(741, 368)
        Me.gbTableList.TabIndex = 77
        Me.gbTableList.TabStop = False
        Me.gbTableList.Text = "List of Tables"
        '
        'dgvTables
        '
        Me.dgvTables.AllowUserToAddRows = False
        Me.dgvTables.AllowUserToDeleteRows = False
        Me.dgvTables.AllowUserToOrderColumns = True
        Me.dgvTables.AllowUserToResizeRows = False
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.InactiveBorder
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ActiveCaptionText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTables.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.dgvTables.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvTables.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.dgvTables.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.Disable
        Me.dgvTables.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.InactiveBorder
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.NullValue = Nothing
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTables.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvTables.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvTables.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically
        Me.dgvTables.GridColor = System.Drawing.SystemColors.InactiveCaption
        Me.dgvTables.Location = New System.Drawing.Point(4, 20)
        Me.dgvTables.Margin = New System.Windows.Forms.Padding(4)
        Me.dgvTables.Name = "dgvTables"
        Me.dgvTables.ReadOnly = True
        Me.dgvTables.RowHeadersVisible = False
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.NullValue = Nothing
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTables.RowsDefaultCellStyle = DataGridViewCellStyle3
        Me.dgvTables.RowTemplate.ReadOnly = True
        Me.dgvTables.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvTables.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvTables.ShowCellErrors = False
        Me.dgvTables.ShowCellToolTips = False
        Me.dgvTables.ShowEditingIcon = False
        Me.dgvTables.ShowRowErrors = False
        Me.dgvTables.Size = New System.Drawing.Size(733, 344)
        Me.dgvTables.StandardTab = True
        Me.dgvTables.TabIndex = 1
        '
        'gbNarrowSearch
        '
        Me.gbNarrowSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbNarrowSearch.Controls.Add(Me.Label7)
        Me.gbNarrowSearch.Controls.Add(Me.Label3)
        Me.gbNarrowSearch.Controls.Add(Me.Label8)
        Me.gbNarrowSearch.Controls.Add(Me.txtSchemaTab2)
        Me.gbNarrowSearch.Controls.Add(Me.txtPartTableName)
        Me.gbNarrowSearch.Controls.Add(Me.btnRefresh)
        Me.gbNarrowSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbNarrowSearch.Location = New System.Drawing.Point(8, 95)
        Me.gbNarrowSearch.Margin = New System.Windows.Forms.Padding(4)
        Me.gbNarrowSearch.Name = "gbNarrowSearch"
        Me.gbNarrowSearch.Padding = New System.Windows.Forms.Padding(4)
        Me.gbNarrowSearch.Size = New System.Drawing.Size(741, 133)
        Me.gbNarrowSearch.TabIndex = 76
        Me.gbNarrowSearch.TabStop = False
        Me.gbNarrowSearch.Text = "Narrow Your Search"
        '
        'Label7
        '
        Me.Label7.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label7.Location = New System.Drawing.Point(8, 90)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(571, 37)
        Me.Label7.TabIndex = 75
        Me.Label7.Text = "Enter a Schema Name and/or partial table name to narrow your search, then click R" & _
    "efresh Table. NOTE: ""%"" is the wildcard for Partial Names."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(8, 27)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(243, 17)
        Me.Label3.TabIndex = 71
        Me.Label3.Text = "Enter a Schema (case sensitive)"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(8, 59)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(320, 17)
        Me.Label8.TabIndex = 73
        Me.Label8.Text = "Enter a Partial TableName (case sensitive)"
        '
        'txtSchemaTab2
        '
        Me.txtSchemaTab2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSchemaTab2.Location = New System.Drawing.Point(348, 23)
        Me.txtSchemaTab2.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSchemaTab2.Name = "txtSchemaTab2"
        Me.txtSchemaTab2.Size = New System.Drawing.Size(384, 23)
        Me.txtSchemaTab2.TabIndex = 70
        '
        'txtPartTableName
        '
        Me.txtPartTableName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPartTableName.Location = New System.Drawing.Point(348, 55)
        Me.txtPartTableName.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPartTableName.Name = "txtPartTableName"
        Me.txtPartTableName.Size = New System.Drawing.Size(384, 23)
        Me.txtPartTableName.TabIndex = 72
        '
        'btnRefresh
        '
        Me.btnRefresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefresh.Location = New System.Drawing.Point(588, 90)
        Me.btnRefresh.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(147, 36)
        Me.btnRefresh.TabIndex = 74
        Me.btnRefresh.Text = "&Refresh Table"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabSelectTable)
        Me.TabControl1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.HotTrack = True
        Me.TabControl1.Location = New System.Drawing.Point(16, 91)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(4)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(768, 642)
        Me.TabControl1.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.TabControl1.TabIndex = 1
        '
        'btnRemoveTab
        '
        Me.btnRemoveTab.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRemoveTab.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRemoveTab.Location = New System.Drawing.Point(253, 769)
        Me.btnRemoveTab.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRemoveTab.Name = "btnRemoveTab"
        Me.btnRemoveTab.Size = New System.Drawing.Size(145, 28)
        Me.btnRemoveTab.TabIndex = 59
        Me.btnRemoveTab.Text = "Remove Tab"
        Me.btnRemoveTab.UseVisualStyleBackColor = True
        '
        'frmDMLInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 816)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnNext)
        Me.Controls.Add(Me.btnRemoveTab)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MinimumSize = New System.Drawing.Size(805, 847)
        Me.Name = "frmDMLInfo"
        Me.Text = "Add Relational DML Structure"
        Me.Controls.SetChildIndex(Me.btnRemoveTab, 0)
        Me.Controls.SetChildIndex(Me.btnNext, 0)
        Me.Controls.SetChildIndex(Me.TabControl1, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.TabSelectTable.ResumeLayout(False)
        Me.gbChseConn.ResumeLayout(False)
        Me.gbChseConn.PerformLayout()
        Me.gbTableList.ResumeLayout(False)
        CType(Me.dgvTables, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbNarrowSearch.ResumeLayout(False)
        Me.gbNarrowSearch.PerformLayout()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents BottomToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents TopToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents RightToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents LeftToolStripPanel As System.Windows.Forms.ToolStripPanel
    Friend WithEvents ContentPanel As System.Windows.Forms.ToolStripContentPanel
    Friend WithEvents TabSelectTable As System.Windows.Forms.TabPage
    Friend WithEvents gbChseConn As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbConn As System.Windows.Forms.ComboBox
    Friend WithEvents gbTableList As System.Windows.Forms.GroupBox
    Friend WithEvents dgvTables As System.Windows.Forms.DataGridView
    Friend WithEvents gbNarrowSearch As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtSchemaTab2 As System.Windows.Forms.TextBox
    Friend WithEvents txtPartTableName As System.Windows.Forms.TextBox
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents btnRemoveTab As System.Windows.Forms.Button
End Class
