<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ctlSelectTab
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctlSelectTab))
        Me.dgvColumns = New System.Windows.Forms.DataGridView
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SelectAllColumnsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.SelectAllColumnsToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem
        Me.btnSelectAll = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.gblvColumns = New System.Windows.Forms.GroupBox
        Me.ToolStripContainer2 = New System.Windows.Forms.ToolStripContainer
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip
        Me.btnUp = New System.Windows.Forms.ToolStripButton
        Me.btnDown = New System.Windows.Forms.ToolStripButton
        Me.btnDel = New System.Windows.Forms.ToolStripButton
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnName = New System.Windows.Forms.ColumnHeader
        Me.ColumnType = New System.Windows.Forms.ColumnHeader
        Me.ColumnLength = New System.Windows.Forms.ColumnHeader
        Me.ColumnScale = New System.Windows.Forms.ColumnHeader
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuRemove = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuMoveUp = New System.Windows.Forms.ToolStripMenuItem
        Me.mnuMoveDown = New System.Windows.Forms.ToolStripMenuItem
        Me.gbColumns = New System.Windows.Forms.GroupBox
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.btnAddAll = New System.Windows.Forms.Button
        CType(Me.dgvColumns, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.gblvColumns.SuspendLayout()
        Me.ToolStripContainer2.ContentPanel.SuspendLayout()
        Me.ToolStripContainer2.SuspendLayout()
        Me.ToolStrip1.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.gbColumns.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvColumns
        '
        Me.dgvColumns.AllowUserToAddRows = False
        Me.dgvColumns.AllowUserToDeleteRows = False
        Me.dgvColumns.AllowUserToOrderColumns = True
        Me.dgvColumns.AllowUserToResizeRows = False
        Me.dgvColumns.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvColumns.BackgroundColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.dgvColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvColumns.ContextMenuStrip = Me.ContextMenuStrip2
        Me.dgvColumns.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvColumns.GridColor = System.Drawing.SystemColors.InactiveCaption
        Me.dgvColumns.Location = New System.Drawing.Point(3, 16)
        Me.dgvColumns.MultiSelect = False
        Me.dgvColumns.Name = "dgvColumns"
        Me.dgvColumns.ReadOnly = True
        Me.dgvColumns.RowHeadersVisible = False
        Me.dgvColumns.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvColumns.Size = New System.Drawing.Size(534, 262)
        Me.dgvColumns.TabIndex = 0
        Me.dgvColumns.VirtualMode = True
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectAllColumnsToolStripMenuItem, Me.SelectAllColumnsToolStripMenuItem1})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(184, 48)
        '
        'SelectAllColumnsToolStripMenuItem
        '
        Me.SelectAllColumnsToolStripMenuItem.Name = "SelectAllColumnsToolStripMenuItem"
        Me.SelectAllColumnsToolStripMenuItem.Size = New System.Drawing.Size(183, 22)
        Me.SelectAllColumnsToolStripMenuItem.Text = "Unselect All Columns"
        '
        'SelectAllColumnsToolStripMenuItem1
        '
        Me.SelectAllColumnsToolStripMenuItem1.Name = "SelectAllColumnsToolStripMenuItem1"
        Me.SelectAllColumnsToolStripMenuItem1.Size = New System.Drawing.Size(183, 22)
        Me.SelectAllColumnsToolStripMenuItem1.Text = "Select All Columns"
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelectAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelectAll.Location = New System.Drawing.Point(442, 287)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.Size = New System.Drawing.Size(98, 23)
        Me.btnSelectAll.TabIndex = 66
        Me.btnSelectAll.Text = "Unselect All"
        Me.btnSelectAll.UseVisualStyleBackColor = True
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label11.Location = New System.Drawing.Point(140, 287)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(133, 26)
        Me.Label11.TabIndex = 65
        Me.Label11.Text = "Right-Click above for" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "column options"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gblvColumns
        '
        Me.gblvColumns.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gblvColumns.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gblvColumns.Controls.Add(Me.ToolStripContainer2)
        Me.gblvColumns.Controls.Add(Me.ListView1)
        Me.gblvColumns.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gblvColumns.Location = New System.Drawing.Point(6, 316)
        Me.gblvColumns.Name = "gblvColumns"
        Me.gblvColumns.Size = New System.Drawing.Size(537, 182)
        Me.gblvColumns.TabIndex = 64
        Me.gblvColumns.TabStop = False
        Me.gblvColumns.Text = "Selected Columns"
        '
        'ToolStripContainer2
        '
        Me.ToolStripContainer2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ToolStripContainer2.BottomToolStripPanelVisible = False
        '
        'ToolStripContainer2.ContentPanel
        '
        Me.ToolStripContainer2.ContentPanel.Controls.Add(Me.ToolStrip1)
        Me.ToolStripContainer2.ContentPanel.Size = New System.Drawing.Size(31, 160)
        Me.ToolStripContainer2.LeftToolStripPanelVisible = False
        Me.ToolStripContainer2.Location = New System.Drawing.Point(6, 16)
        Me.ToolStripContainer2.Name = "ToolStripContainer2"
        Me.ToolStripContainer2.RightToolStripPanelVisible = False
        Me.ToolStripContainer2.Size = New System.Drawing.Size(31, 160)
        Me.ToolStripContainer2.TabIndex = 59
        Me.ToolStripContainer2.Text = "ToolStripContainer2"
        Me.ToolStripContainer2.TopToolStripPanelVisible = False
        '
        'ToolStrip1
        '
        Me.ToolStrip1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStrip1.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnUp, Me.btnDown, Me.btnDel})
        Me.ToolStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.VerticalStackWithOverflow
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(31, 160)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btnUp
        '
        Me.btnUp.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnUp.Image = CType(resources.GetObject("btnUp.Image"), System.Drawing.Image)
        Me.btnUp.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnUp.Name = "btnUp"
        Me.btnUp.Size = New System.Drawing.Size(29, 20)
        Me.btnUp.Text = "Move Up Button"
        Me.btnUp.ToolTipText = "Move Column Up"
        '
        'btnDown
        '
        Me.btnDown.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnDown.Image = CType(resources.GetObject("btnDown.Image"), System.Drawing.Image)
        Me.btnDown.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDown.Name = "btnDown"
        Me.btnDown.Size = New System.Drawing.Size(29, 20)
        Me.btnDown.Text = "Move Down Button"
        Me.btnDown.ToolTipText = "Move Column Down"
        '
        'btnDel
        '
        Me.btnDel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnDel.Image = CType(resources.GetObject("btnDel.Image"), System.Drawing.Image)
        Me.btnDel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnDel.Name = "btnDel"
        Me.btnDel.Size = New System.Drawing.Size(29, 20)
        Me.btnDel.Text = "Remove Button"
        Me.btnDel.ToolTipText = "Remove Column"
        '
        'ListView1
        '
        Me.ListView1.Alignment = System.Windows.Forms.ListViewAlignment.Left
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.AutoArrange = False
        Me.ListView1.BackColor = System.Drawing.SystemColors.Window
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnName, Me.ColumnType, Me.ColumnLength, Me.ColumnScale})
        Me.ListView1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.LabelWrap = False
        Me.ListView1.Location = New System.Drawing.Point(40, 16)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.ShowItemToolTips = True
        Me.ListView1.Size = New System.Drawing.Size(491, 160)
        Me.ListView1.TabIndex = 0
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnName
        '
        Me.ColumnName.Text = "Column Name"
        Me.ColumnName.Width = 197
        '
        'ColumnType
        '
        Me.ColumnType.Text = "Data Type"
        Me.ColumnType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnType.Width = 140
        '
        'ColumnLength
        '
        Me.ColumnLength.Text = "Length"
        Me.ColumnLength.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnLength.Width = 80
        '
        'ColumnScale
        '
        Me.ColumnScale.Text = "Scale"
        Me.ColumnScale.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.ColumnScale.Width = 68
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuRemove, Me.mnuMoveUp, Me.mnuMoveDown})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(142, 70)
        '
        'mnuRemove
        '
        Me.mnuRemove.Name = "mnuRemove"
        Me.mnuRemove.Size = New System.Drawing.Size(141, 22)
        Me.mnuRemove.Text = "Remove"
        '
        'mnuMoveUp
        '
        Me.mnuMoveUp.Name = "mnuMoveUp"
        Me.mnuMoveUp.Size = New System.Drawing.Size(141, 22)
        Me.mnuMoveUp.Text = "Move Up"
        '
        'mnuMoveDown
        '
        Me.mnuMoveDown.Name = "mnuMoveDown"
        Me.mnuMoveDown.Size = New System.Drawing.Size(141, 22)
        Me.mnuMoveDown.Text = "Move Down"
        '
        'gbColumns
        '
        Me.gbColumns.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbColumns.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gbColumns.Controls.Add(Me.dgvColumns)
        Me.gbColumns.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbColumns.Location = New System.Drawing.Point(3, 3)
        Me.gbColumns.Name = "gbColumns"
        Me.gbColumns.Size = New System.Drawing.Size(540, 281)
        Me.gbColumns.TabIndex = 63
        Me.gbColumns.TabStop = False
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "delete.ico")
        Me.ImageList1.Images.SetKeyName(2, "UpArrow.ico")
        Me.ImageList1.Images.SetKeyName(3, "DownArrow.ico")
        '
        'btnAddAll
        '
        Me.btnAddAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAddAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnAddAll.Location = New System.Drawing.Point(338, 287)
        Me.btnAddAll.Name = "btnAddAll"
        Me.btnAddAll.Size = New System.Drawing.Size(98, 23)
        Me.btnAddAll.TabIndex = 67
        Me.btnAddAll.Text = "Select All"
        Me.btnAddAll.UseVisualStyleBackColor = True
        '
        'ctlSelectTab
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.Controls.Add(Me.gblvColumns)
        Me.Controls.Add(Me.btnSelectAll)
        Me.Controls.Add(Me.gbColumns)
        Me.Controls.Add(Me.btnAddAll)
        Me.Controls.Add(Me.Label11)
        Me.Name = "ctlSelectTab"
        Me.Size = New System.Drawing.Size(549, 504)
        CType(Me.dgvColumns, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.gblvColumns.ResumeLayout(False)
        Me.ToolStripContainer2.ContentPanel.ResumeLayout(False)
        Me.ToolStripContainer2.ContentPanel.PerformLayout()
        Me.ToolStripContainer2.ResumeLayout(False)
        Me.ToolStripContainer2.PerformLayout()
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.gbColumns.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dgvColumns As System.Windows.Forms.DataGridView
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents gblvColumns As System.Windows.Forms.GroupBox
    Friend WithEvents ToolStripContainer2 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnUp As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnDown As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnDel As System.Windows.Forms.ToolStripButton
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnName As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnType As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnLength As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnScale As System.Windows.Forms.ColumnHeader
    Friend WithEvents gbColumns As System.Windows.Forms.GroupBox
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents SelectAllColumnsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents mnuRemove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMoveUp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents mnuMoveDown As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SelectAllColumnsToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnAddAll As System.Windows.Forms.Button

End Class
