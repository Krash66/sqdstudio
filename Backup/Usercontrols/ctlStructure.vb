Public Class ctlStructure
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    
    Dim objThis As New clsStructure
    Dim IsNewObj As Boolean
    Dim lastTreeviewSearchText As String
    Dim lastTreeviewSearchNode As TreeNode
    Dim prevFld As clsField
    Private Split As Integer


    Dim colSkipNodes As New ArrayList

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents chkAddDBD As System.Windows.Forms.CheckBox
    Friend WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseDBDFile As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseCobolFile As System.Windows.Forms.Button
    Friend WithEvents txtCobolFilePath As System.Windows.Forms.TextBox
    Friend WithEvents tvSegments As System.Windows.Forms.TreeView
    Friend WithEvents lblCobolFile As System.Windows.Forms.Label
    Friend WithEvents lblDBD As System.Windows.Forms.Label
    Friend WithEvents txtStructName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents txtSearchField As System.Windows.Forms.TextBox
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents txtStructDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtDBDFilePath As System.Windows.Forms.TextBox
    Friend WithEvents txtFieldDesc As System.Windows.Forms.TextBox
    Friend WithEvents A4 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents AddToSourceMappingTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddToTargetMappingTableToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblConn As System.Windows.Forms.Label
    Friend WithEvents txtConn As System.Windows.Forms.TextBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents scStr As System.Windows.Forms.SplitContainer
    Friend WithEvents scFld As System.Windows.Forms.SplitContainer
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbSeg As System.Windows.Forms.GroupBox
    Friend WithEvents gbFields As System.Windows.Forms.GroupBox
    Friend WithEvents scFldDesc As System.Windows.Forms.SplitContainer
    Friend WithEvents gbAttr As System.Windows.Forms.GroupBox
    Friend WithEvents gbFldDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbSearch As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbConn As System.Windows.Forms.GroupBox
    Friend WithEvents gbCobol As System.Windows.Forms.GroupBox
    Friend WithEvents gbFile As System.Windows.Forms.GroupBox
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents btnViewDBD As System.Windows.Forms.Button
    Friend WithEvents btnViewCobol As System.Windows.Forms.Button
    Friend WithEvents gbProp As System.Windows.Forms.GroupBox
    Friend WithEvents txtLength As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtTablespace As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctlStructure))
        Me.txtConn = New System.Windows.Forms.TextBox
        Me.lblConn = New System.Windows.Forms.Label
        Me.txtStructDesc = New System.Windows.Forms.TextBox
        Me.chkAddDBD = New System.Windows.Forms.CheckBox
        Me.cmdBrowseFile = New System.Windows.Forms.Button
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.lblFile = New System.Windows.Forms.Label
        Me.cmdBrowseDBDFile = New System.Windows.Forms.Button
        Me.txtDBDFilePath = New System.Windows.Forms.TextBox
        Me.cmdBrowseCobolFile = New System.Windows.Forms.Button
        Me.txtCobolFilePath = New System.Windows.Forms.TextBox
        Me.tvSegments = New System.Windows.Forms.TreeView
        Me.lblCobolFile = New System.Windows.Forms.Label
        Me.lblDBD = New System.Windows.Forms.Label
        Me.txtStructName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.A4 = New System.Windows.Forms.Label
        Me.txtFieldDesc = New System.Windows.Forms.TextBox
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.txtSearchField = New System.Windows.Forms.TextBox
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.AddToSourceMappingTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.AddToTargetMappingTableToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.scStr = New System.Windows.Forms.SplitContainer
        Me.gbProp = New System.Windows.Forms.GroupBox
        Me.txtLength = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.gbConn = New System.Windows.Forms.GroupBox
        Me.gbFile = New System.Windows.Forms.GroupBox
        Me.btnView = New System.Windows.Forms.Button
        Me.gbCobol = New System.Windows.Forms.GroupBox
        Me.btnViewDBD = New System.Windows.Forms.Button
        Me.btnViewCobol = New System.Windows.Forms.Button
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbSeg = New System.Windows.Forms.GroupBox
        Me.scFld = New System.Windows.Forms.SplitContainer
        Me.gbFields = New System.Windows.Forms.GroupBox
        Me.gbSearch = New System.Windows.Forms.GroupBox
        Me.scFldDesc = New System.Windows.Forms.SplitContainer
        Me.gbAttr = New System.Windows.Forms.GroupBox
        Me.gbFldDesc = New System.Windows.Forms.GroupBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.txtTablespace = New System.Windows.Forms.TextBox
        Me.ContextMenuStrip1.SuspendLayout()
        Me.scStr.Panel1.SuspendLayout()
        Me.scStr.Panel2.SuspendLayout()
        Me.scStr.SuspendLayout()
        Me.gbProp.SuspendLayout()
        Me.gbConn.SuspendLayout()
        Me.gbFile.SuspendLayout()
        Me.gbCobol.SuspendLayout()
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbSeg.SuspendLayout()
        Me.scFld.Panel1.SuspendLayout()
        Me.scFld.Panel2.SuspendLayout()
        Me.scFld.SuspendLayout()
        Me.gbFields.SuspendLayout()
        Me.gbSearch.SuspendLayout()
        Me.scFldDesc.Panel1.SuspendLayout()
        Me.scFldDesc.Panel2.SuspendLayout()
        Me.scFldDesc.SuspendLayout()
        Me.gbAttr.SuspendLayout()
        Me.gbFldDesc.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtConn
        '
        Me.txtConn.Location = New System.Drawing.Point(85, 19)
        Me.txtConn.Name = "txtConn"
        Me.txtConn.ReadOnly = True
        Me.txtConn.Size = New System.Drawing.Size(369, 20)
        Me.txtConn.TabIndex = 77
        Me.txtConn.TabStop = False
        '
        'lblConn
        '
        Me.lblConn.AutoSize = True
        Me.lblConn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConn.Location = New System.Drawing.Point(6, 22)
        Me.lblConn.Name = "lblConn"
        Me.lblConn.Size = New System.Drawing.Size(71, 13)
        Me.lblConn.TabIndex = 76
        Me.lblConn.Text = "Connection"
        '
        'txtStructDesc
        '
        Me.txtStructDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtStructDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStructDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtStructDesc.MaxLength = 255
        Me.txtStructDesc.Multiline = True
        Me.txtStructDesc.Name = "txtStructDesc"
        Me.txtStructDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStructDesc.Size = New System.Drawing.Size(296, 52)
        Me.txtStructDesc.TabIndex = 10
        '
        'chkAddDBD
        '
        Me.chkAddDBD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddDBD.Location = New System.Drawing.Point(85, 45)
        Me.chkAddDBD.Name = "chkAddDBD"
        Me.chkAddDBD.Size = New System.Drawing.Size(114, 20)
        Me.chkAddDBD.TabIndex = 2
        Me.chkAddDBD.Text = "Add DBD Source"
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.Location = New System.Drawing.Point(372, 18)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 4
        Me.cmdBrowseFile.Text = "..."
        '
        'txtFilePath
        '
        Me.txtFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilePath.Location = New System.Drawing.Point(38, 19)
        Me.txtFilePath.MaxLength = 255
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(328, 20)
        Me.txtFilePath.TabIndex = 3
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFile.Location = New System.Drawing.Point(6, 22)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(26, 14)
        Me.lblFile.TabIndex = 65
        Me.lblFile.Text = "File"
        '
        'cmdBrowseDBDFile
        '
        Me.cmdBrowseDBDFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseDBDFile.Location = New System.Drawing.Point(372, 18)
        Me.cmdBrowseDBDFile.Name = "cmdBrowseDBDFile"
        Me.cmdBrowseDBDFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseDBDFile.TabIndex = 8
        Me.cmdBrowseDBDFile.Text = "..."
        '
        'txtDBDFilePath
        '
        Me.txtDBDFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDBDFilePath.Location = New System.Drawing.Point(85, 19)
        Me.txtDBDFilePath.MaxLength = 255
        Me.txtDBDFilePath.Name = "txtDBDFilePath"
        Me.txtDBDFilePath.Size = New System.Drawing.Size(281, 20)
        Me.txtDBDFilePath.TabIndex = 7
        '
        'cmdBrowseCobolFile
        '
        Me.cmdBrowseCobolFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseCobolFile.Location = New System.Drawing.Point(372, 70)
        Me.cmdBrowseCobolFile.Name = "cmdBrowseCobolFile"
        Me.cmdBrowseCobolFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseCobolFile.TabIndex = 6
        Me.cmdBrowseCobolFile.Text = "..."
        '
        'txtCobolFilePath
        '
        Me.txtCobolFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCobolFilePath.Location = New System.Drawing.Point(85, 71)
        Me.txtCobolFilePath.MaxLength = 255
        Me.txtCobolFilePath.Name = "txtCobolFilePath"
        Me.txtCobolFilePath.Size = New System.Drawing.Size(281, 20)
        Me.txtCobolFilePath.TabIndex = 5
        '
        'tvSegments
        '
        Me.tvSegments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvSegments.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvSegments.Location = New System.Drawing.Point(3, 16)
        Me.tvSegments.Name = "tvSegments"
        Me.tvSegments.Size = New System.Drawing.Size(228, 315)
        Me.tvSegments.TabIndex = 9
        '
        'lblCobolFile
        '
        Me.lblCobolFile.AutoSize = True
        Me.lblCobolFile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCobolFile.Location = New System.Drawing.Point(6, 74)
        Me.lblCobolFile.Name = "lblCobolFile"
        Me.lblCobolFile.Size = New System.Drawing.Size(61, 14)
        Me.lblCobolFile.TabIndex = 58
        Me.lblCobolFile.Text = "Cobol File"
        '
        'lblDBD
        '
        Me.lblDBD.AutoSize = True
        Me.lblDBD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBD.Location = New System.Drawing.Point(6, 22)
        Me.lblDBD.Name = "lblDBD"
        Me.lblDBD.Size = New System.Drawing.Size(73, 14)
        Me.lblDBD.TabIndex = 57
        Me.lblDBD.Text = "IMS DBD File"
        '
        'txtStructName
        '
        Me.txtStructName.Location = New System.Drawing.Point(50, 16)
        Me.txtStructName.MaxLength = 128
        Me.txtStructName.Name = "txtStructName"
        Me.txtStructName.Size = New System.Drawing.Size(258, 20)
        Me.txtStructName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 19)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 55
        Me.Label3.Text = "Name"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label6.Location = New System.Drawing.Point(30, 553)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(139, 13)
        Me.Label6.TabIndex = 90
        Me.Label6.Text = "Has a Field Description"
        '
        'A4
        '
        Me.A4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.A4.BackColor = System.Drawing.Color.White
        Me.A4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.A4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A4.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.A4.Location = New System.Drawing.Point(8, 552)
        Me.A4.Name = "A4"
        Me.A4.Size = New System.Drawing.Size(16, 16)
        Me.A4.TabIndex = 89
        Me.A4.Text = "A"
        Me.A4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFieldDesc
        '
        Me.txtFieldDesc.AcceptsReturn = True
        Me.txtFieldDesc.AllowDrop = True
        Me.txtFieldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtFieldDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFieldDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtFieldDesc.Multiline = True
        Me.txtFieldDesc.Name = "txtFieldDesc"
        Me.txtFieldDesc.Size = New System.Drawing.Size(478, 32)
        Me.txtFieldDesc.TabIndex = 15
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(630, 547)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 22
        Me.cmdHelp.Text = "&Help"
        '
        'cmdSearch
        '
        Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.Location = New System.Drawing.Point(171, 18)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(30, 20)
        Me.cmdSearch.TabIndex = 12
        Me.cmdSearch.Text = "Go"
        '
        'txtSearchField
        '
        Me.txtSearchField.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearchField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSearchField.Location = New System.Drawing.Point(6, 19)
        Me.txtSearchField.MaxLength = 128
        Me.txtSearchField.Name = "txtSearchField"
        Me.txtSearchField.Size = New System.Drawing.Size(159, 20)
        Me.txtSearchField.TabIndex = 11
        '
        'tvFields
        '
        Me.tvFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvFields.ContextMenuStrip = Me.ContextMenuStrip1
        Me.tvFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvFields.HideSelection = False
        Me.tvFields.HotTracking = True
        Me.tvFields.ImageIndex = 1
        Me.tvFields.ImageList = Me.ImageList1
        Me.tvFields.Location = New System.Drawing.Point(3, 19)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.SelectedImageIndex = 0
        Me.tvFields.Size = New System.Drawing.Size(213, 109)
        Me.tvFields.StateImageList = Me.ImageList1
        Me.tvFields.TabIndex = 13
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AddToSourceMappingTableToolStripMenuItem, Me.AddToTargetMappingTableToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(226, 48)
        '
        'AddToSourceMappingTableToolStripMenuItem
        '
        Me.AddToSourceMappingTableToolStripMenuItem.Name = "AddToSourceMappingTableToolStripMenuItem"
        Me.AddToSourceMappingTableToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.AddToSourceMappingTableToolStripMenuItem.Text = "Add to Source Mapping Table"
        '
        'AddToTargetMappingTableToolStripMenuItem
        '
        Me.AddToTargetMappingTableToolStripMenuItem.Name = "AddToTargetMappingTableToolStripMenuItem"
        Me.AddToTargetMappingTableToolStripMenuItem.Size = New System.Drawing.Size(225, 22)
        Me.AddToTargetMappingTableToolStripMenuItem.Text = "Add to Target Mapping Table"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "105.ico")
        Me.ImageList1.Images.SetKeyName(1, "Field List.ico")
        '
        'lvFieldAttrs
        '
        Me.lvFieldAttrs.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.lvFieldAttrs.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvFieldAttrs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvFieldAttrs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvFieldAttrs.FullRowSelect = True
        Me.lvFieldAttrs.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.lvFieldAttrs.HideSelection = False
        Me.lvFieldAttrs.Location = New System.Drawing.Point(3, 16)
        Me.lvFieldAttrs.Name = "lvFieldAttrs"
        Me.lvFieldAttrs.Size = New System.Drawing.Size(478, 111)
        Me.lvFieldAttrs.TabIndex = 14
        Me.lvFieldAttrs.UseCompatibleStateImageBehavior = False
        Me.lvFieldAttrs.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Attribute"
        Me.ColumnHeader1.Width = 232
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Value"
        Me.ColumnHeader2.Width = 215
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(472, 547)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 20
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(552, 547)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 21
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 533)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(696, 7)
        Me.GroupBox1.TabIndex = 38
        Me.GroupBox1.TabStop = False
        '
        'scStr
        '
        Me.scStr.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scStr.Location = New System.Drawing.Point(3, 3)
        Me.scStr.Name = "scStr"
        Me.scStr.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scStr.Panel1
        '
        Me.scStr.Panel1.Controls.Add(Me.GroupBox2)
        Me.scStr.Panel1.Controls.Add(Me.gbProp)
        Me.scStr.Panel1.Controls.Add(Me.gbConn)
        Me.scStr.Panel1.Controls.Add(Me.gbFile)
        Me.scStr.Panel1.Controls.Add(Me.gbCobol)
        Me.scStr.Panel1.Controls.Add(Me.gbName)
        Me.scStr.Panel1.Controls.Add(Me.gbSeg)
        '
        'scStr.Panel2
        '
        Me.scStr.Panel2.Controls.Add(Me.scFld)
        Me.scStr.Size = New System.Drawing.Size(706, 528)
        Me.scStr.SplitterDistance = 339
        Me.scStr.TabIndex = 39
        '
        'gbProp
        '
        Me.gbProp.Controls.Add(Me.txtLength)
        Me.gbProp.Controls.Add(Me.Label1)
        Me.gbProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProp.ForeColor = System.Drawing.Color.White
        Me.gbProp.Location = New System.Drawing.Point(326, 3)
        Me.gbProp.Name = "gbProp"
        Me.gbProp.Size = New System.Drawing.Size(137, 63)
        Me.gbProp.TabIndex = 82
        Me.gbProp.TabStop = False
        Me.gbProp.Text = "Record Properties"
        '
        'txtLength
        '
        Me.txtLength.Location = New System.Drawing.Point(6, 35)
        Me.txtLength.Name = "txtLength"
        Me.txtLength.Size = New System.Drawing.Size(125, 20)
        Me.txtLength.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(3, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(91, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Record Length"
        '
        'gbConn
        '
        Me.gbConn.Controls.Add(Me.lblConn)
        Me.gbConn.Controls.Add(Me.txtConn)
        Me.gbConn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbConn.ForeColor = System.Drawing.Color.White
        Me.gbConn.Location = New System.Drawing.Point(3, 185)
        Me.gbConn.MinimumSize = New System.Drawing.Size(432, 46)
        Me.gbConn.Name = "gbConn"
        Me.gbConn.Size = New System.Drawing.Size(460, 46)
        Me.gbConn.TabIndex = 80
        Me.gbConn.TabStop = False
        Me.gbConn.Text = "DML Connection"
        '
        'gbFile
        '
        Me.gbFile.Controls.Add(Me.btnView)
        Me.gbFile.Controls.Add(Me.lblFile)
        Me.gbFile.Controls.Add(Me.txtFilePath)
        Me.gbFile.Controls.Add(Me.cmdBrowseFile)
        Me.gbFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFile.ForeColor = System.Drawing.Color.White
        Me.gbFile.Location = New System.Drawing.Point(3, 132)
        Me.gbFile.MinimumSize = New System.Drawing.Size(432, 47)
        Me.gbFile.Name = "gbFile"
        Me.gbFile.Size = New System.Drawing.Size(460, 47)
        Me.gbFile.TabIndex = 81
        Me.gbFile.TabStop = False
        Me.gbFile.Text = "Description File"
        '
        'btnView
        '
        Me.btnView.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btnView.Location = New System.Drawing.Point(403, 17)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(51, 23)
        Me.btnView.TabIndex = 66
        Me.btnView.Text = "View"
        Me.btnView.UseVisualStyleBackColor = False
        '
        'gbCobol
        '
        Me.gbCobol.Controls.Add(Me.btnViewDBD)
        Me.gbCobol.Controls.Add(Me.btnViewCobol)
        Me.gbCobol.Controls.Add(Me.cmdBrowseCobolFile)
        Me.gbCobol.Controls.Add(Me.txtCobolFilePath)
        Me.gbCobol.Controls.Add(Me.lblCobolFile)
        Me.gbCobol.Controls.Add(Me.chkAddDBD)
        Me.gbCobol.Controls.Add(Me.cmdBrowseDBDFile)
        Me.gbCobol.Controls.Add(Me.lblDBD)
        Me.gbCobol.Controls.Add(Me.txtDBDFilePath)
        Me.gbCobol.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbCobol.ForeColor = System.Drawing.Color.White
        Me.gbCobol.Location = New System.Drawing.Point(3, 237)
        Me.gbCobol.MinimumSize = New System.Drawing.Size(432, 100)
        Me.gbCobol.Name = "gbCobol"
        Me.gbCobol.Size = New System.Drawing.Size(460, 100)
        Me.gbCobol.TabIndex = 79
        Me.gbCobol.TabStop = False
        Me.gbCobol.Text = "Cobol File"
        '
        'btnViewDBD
        '
        Me.btnViewDBD.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btnViewDBD.Location = New System.Drawing.Point(403, 17)
        Me.btnViewDBD.Name = "btnViewDBD"
        Me.btnViewDBD.Size = New System.Drawing.Size(51, 23)
        Me.btnViewDBD.TabIndex = 60
        Me.btnViewDBD.Text = "View"
        Me.btnViewDBD.UseVisualStyleBackColor = False
        '
        'btnViewCobol
        '
        Me.btnViewCobol.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btnViewCobol.Location = New System.Drawing.Point(403, 69)
        Me.btnViewCobol.Name = "btnViewCobol"
        Me.btnViewCobol.Size = New System.Drawing.Size(51, 23)
        Me.btnViewCobol.TabIndex = 59
        Me.btnViewCobol.Text = "View"
        Me.btnViewCobol.UseVisualStyleBackColor = False
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Controls.Add(Me.txtStructName)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.Color.White
        Me.gbName.Location = New System.Drawing.Point(3, 3)
        Me.gbName.MaximumSize = New System.Drawing.Size(317, 124)
        Me.gbName.MinimumSize = New System.Drawing.Size(317, 124)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(317, 124)
        Me.gbName.TabIndex = 0
        Me.gbName.TabStop = False
        Me.gbName.Text = "Description Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtStructDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 45)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(302, 71)
        Me.gbDesc.TabIndex = 78
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbSeg
        '
        Me.gbSeg.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbSeg.Controls.Add(Me.tvSegments)
        Me.gbSeg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSeg.ForeColor = System.Drawing.Color.White
        Me.gbSeg.Location = New System.Drawing.Point(469, 3)
        Me.gbSeg.Name = "gbSeg"
        Me.gbSeg.Size = New System.Drawing.Size(234, 334)
        Me.gbSeg.TabIndex = 0
        Me.gbSeg.TabStop = False
        Me.gbSeg.Text = "Segments"
        '
        'scFld
        '
        Me.scFld.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scFld.Location = New System.Drawing.Point(0, 0)
        Me.scFld.Name = "scFld"
        '
        'scFld.Panel1
        '
        Me.scFld.Panel1.Controls.Add(Me.gbFields)
        '
        'scFld.Panel2
        '
        Me.scFld.Panel2.Controls.Add(Me.scFldDesc)
        Me.scFld.Size = New System.Drawing.Size(706, 185)
        Me.scFld.SplitterDistance = 218
        Me.scFld.TabIndex = 38
        '
        'gbFields
        '
        Me.gbFields.Controls.Add(Me.gbSearch)
        Me.gbFields.Controls.Add(Me.tvFields)
        Me.gbFields.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFields.ForeColor = System.Drawing.Color.White
        Me.gbFields.Location = New System.Drawing.Point(0, 0)
        Me.gbFields.Name = "gbFields"
        Me.gbFields.Size = New System.Drawing.Size(218, 185)
        Me.gbFields.TabIndex = 0
        Me.gbFields.TabStop = False
        Me.gbFields.Text = "Fields"
        '
        'gbSearch
        '
        Me.gbSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbSearch.Controls.Add(Me.txtSearchField)
        Me.gbSearch.Controls.Add(Me.cmdSearch)
        Me.gbSearch.ForeColor = System.Drawing.Color.White
        Me.gbSearch.Location = New System.Drawing.Point(6, 134)
        Me.gbSearch.Name = "gbSearch"
        Me.gbSearch.Size = New System.Drawing.Size(207, 46)
        Me.gbSearch.TabIndex = 36
        Me.gbSearch.TabStop = False
        Me.gbSearch.Text = "Search"
        '
        'scFldDesc
        '
        Me.scFldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scFldDesc.Location = New System.Drawing.Point(0, 0)
        Me.scFldDesc.Name = "scFldDesc"
        Me.scFldDesc.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scFldDesc.Panel1
        '
        Me.scFldDesc.Panel1.Controls.Add(Me.gbAttr)
        '
        'scFldDesc.Panel2
        '
        Me.scFldDesc.Panel2.Controls.Add(Me.gbFldDesc)
        Me.scFldDesc.Size = New System.Drawing.Size(484, 185)
        Me.scFldDesc.SplitterDistance = 130
        Me.scFldDesc.TabIndex = 0
        '
        'gbAttr
        '
        Me.gbAttr.Controls.Add(Me.lvFieldAttrs)
        Me.gbAttr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAttr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbAttr.ForeColor = System.Drawing.Color.White
        Me.gbAttr.Location = New System.Drawing.Point(0, 0)
        Me.gbAttr.Name = "gbAttr"
        Me.gbAttr.Size = New System.Drawing.Size(484, 130)
        Me.gbAttr.TabIndex = 0
        Me.gbAttr.TabStop = False
        Me.gbAttr.Text = "Field Attributes"
        '
        'gbFldDesc
        '
        Me.gbFldDesc.Controls.Add(Me.txtFieldDesc)
        Me.gbFldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFldDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFldDesc.ForeColor = System.Drawing.Color.White
        Me.gbFldDesc.Location = New System.Drawing.Point(0, 0)
        Me.gbFldDesc.Name = "gbFldDesc"
        Me.gbFldDesc.Size = New System.Drawing.Size(484, 51)
        Me.gbFldDesc.TabIndex = 0
        Me.gbFldDesc.TabStop = False
        Me.gbFldDesc.Text = "Field Description"
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtTablespace)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(326, 72)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(137, 55)
        Me.GroupBox2.TabIndex = 83
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Tablespace"
        '
        'txtTablespace
        '
        Me.txtTablespace.Location = New System.Drawing.Point(7, 20)
        Me.txtTablespace.Name = "txtTablespace"
        Me.txtTablespace.Size = New System.Drawing.Size(124, 20)
        Me.txtTablespace.TabIndex = 0
        '
        'ctlStructure
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.scStr)
        Me.Controls.Add(Me.A4)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdHelp)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlStructure"
        Me.Size = New System.Drawing.Size(712, 579)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.scStr.Panel1.ResumeLayout(False)
        Me.scStr.Panel2.ResumeLayout(False)
        Me.scStr.ResumeLayout(False)
        Me.gbProp.ResumeLayout(False)
        Me.gbProp.PerformLayout()
        Me.gbConn.ResumeLayout(False)
        Me.gbConn.PerformLayout()
        Me.gbFile.ResumeLayout(False)
        Me.gbFile.PerformLayout()
        Me.gbCobol.ResumeLayout(False)
        Me.gbCobol.PerformLayout()
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbSeg.ResumeLayout(False)
        Me.scFld.Panel1.ResumeLayout(False)
        Me.scFld.Panel2.ResumeLayout(False)
        Me.scFld.ResumeLayout(False)
        Me.gbFields.ResumeLayout(False)
        Me.gbSearch.ResumeLayout(False)
        Me.gbSearch.PerformLayout()
        Me.scFldDesc.Panel1.ResumeLayout(False)
        Me.scFldDesc.Panel2.ResumeLayout(False)
        Me.scFldDesc.ResumeLayout(False)
        Me.gbAttr.ResumeLayout(False)
        Me.gbFldDesc.ResumeLayout(False)
        Me.gbFldDesc.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Form Events"

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtStructName.Enabled = True '//added on 7/24 to disable object name editing

        InitControls()

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsStructure

        ClearControls(Me.Controls)
        txtFieldDesc.Text = ""

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    '//modified on 8/15/05
    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddDBD.CheckedChanged, txtCobolFilePath.TextChanged, txtStructName.TextChanged, txtDBDFilePath.TextChanged, txtStructDesc.TextChanged, txtFilePath.TextChanged, txtTablespace.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub OnFldDescChng(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFieldDesc.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        cmdSave.Enabled = True
        prevFld.FieldDescModified = True

        OnChange(Me, New EventArgs)

    End Sub

    Friend Function ShowHideControls(ByVal StructType As enumStructure) As Boolean
        Dim bIsCOBOLorIMS As Boolean
        Dim bIsDML As Boolean

        bIsCOBOLorIMS = ((StructType = enumStructure.STRUCT_COBOL) Or _
                         (StructType = enumStructure.STRUCT_IMS) Or _
                         (StructType = enumStructure.STRUCT_COBOL_IMS))

        bIsDML = (StructType = enumStructure.STRUCT_REL_DML Or StructType = enumStructure.STRUCT_REL_DML_FILE)

        '/// set GroupBoxes and splitter
        gbCobol.Location() = gbFile.Location
        gbFile.Visible = Not bIsCOBOLorIMS
        gbSeg.Visible = bIsCOBOLorIMS
        gbCobol.Visible = bIsCOBOLorIMS
        gbConn.Visible = bIsDML
        If bIsCOBOLorIMS = True Or bIsDML = True Then
            Split = 235
            gbSeg.Height = Split - 5
        Else
            Split = 183
        End If
        scStr.SplitterDistance = Split

        '//For COB, DBD
        lblCobolFile.Visible = bIsCOBOLorIMS
        lblDBD.Visible = bIsCOBOLorIMS
        cmdBrowseCobolFile.Visible = bIsCOBOLorIMS
        cmdBrowseDBDFile.Visible = bIsCOBOLorIMS
        chkAddDBD.Visible = bIsCOBOLorIMS
        txtCobolFilePath.Visible = bIsCOBOLorIMS
        txtDBDFilePath.Visible = bIsCOBOLorIMS
        tvSegments.Visible = bIsCOBOLorIMS

        '//For XML, DDL, DML
        lblFile.Visible = Not bIsCOBOLorIMS
        txtFilePath.Visible = Not bIsCOBOLorIMS
        cmdBrowseFile.Visible = Not bIsCOBOLorIMS
        '/// for DML
        lblConn.Visible = bIsDML
        txtConn.Visible = bIsDML
        btnView.Enabled = Not bIsDML

        '//Dont allow changing field once object is created coz it might conflict other place
        If IsNewObj = False Then
            cmdBrowseCobolFile.Enabled = True
            cmdBrowseDBDFile.Enabled = True
            cmdBrowseFile.Enabled = True
            txtFilePath.Enabled = True
            txtCobolFilePath.Enabled = True
            txtDBDFilePath.Enabled = True
            chkAddDBD.Enabled = True
        End If

        If StructType = enumStructure.STRUCT_REL_DML Then
            cmdBrowseFile.Enabled = False
        End If

        If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL And IsNewObj = True Then
            txtCobolFilePath.Enabled = True
            txtDBDFilePath.Enabled = False
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS And IsNewObj = True Then
            txtCobolFilePath.Enabled = True
            txtDBDFilePath.Enabled = True
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_IMS And IsNewObj = True Then
            txtCobolFilePath.Enabled = False
            txtDBDFilePath.Enabled = True
        End If

        '//show for all
        If IsNewObj = False Then cmdShowHideFieldAttr_Click(Me, New EventArgs)

    End Function

    Function UpdateControls() As Boolean

        'txtDBDFilePath.Enabled = chkAddDBD.Checked
        'cmdBrowseDBDFile.Enabled = chkAddDBD.Checked
        'tvSegments.Enabled = chkAddDBD.Checked

        If IsNewObj = False Then
            txtDBDFilePath.Enabled = chkAddDBD.Checked
            cmdBrowseDBDFile.Enabled = chkAddDBD.Checked
            tvSegments.Enabled = chkAddDBD.Checked
        End If

    End Function

    Function InitControls() As Boolean

        lvFieldAttrs.Columns(0).Width = (lvFieldAttrs.Width / 2) - 12
        lvFieldAttrs.Columns(1).Width = (lvFieldAttrs.Width / 2) - 12

        tvFields.HideSelection = False
        lvFieldAttrs.SmallImageList = imgListSmall

        UpdateControls()

    End Function

    ''' <summary>
    ''' This will load treeview of fields
    ''' </summary>
    ''' <returns>true unless exeption error</returns>
    ''' <remarks>If New Structure then fields will be from XML file If Edit structure then fields wil be from database</remarks>
    Function FillFieldAttr() As Boolean


        Try
            tvFields.BeginUpdate()
            lvFieldAttrs.BeginUpdate()
            tvFields.Nodes.Clear()

            If IsNewObj = True Then
                Dim outXMLPath As String
                outXMLPath = GetOutDumpXML()
                LoadTreeViewFromXmlFile(objThis, outXMLPath, tvFields, True)
                If IO.File.Exists(outXMLPath) Then
                    IO.File.Delete(outXMLPath)
                End If
            Else
                objThis.LoadItems()
                '//This will fill out parent/child fields tree from Array of fields
                AddFieldsToTreeView(objThis, , tvFields)
            End If

            If tvFields.Nodes.Count > 0 Then
                HiLiteFieldDescNodes(tvFields.Nodes(0), True, tvFields)
                HiLiteFieldKeyNodes(tvFields.Nodes(0), True, tvFields)
            End If

            Return True

        Catch ex As Exception
            LogError(ex)
            Return False
        Finally
            tvFields.EndUpdate()
            lvFieldAttrs.EndUpdate()
        End Try

    End Function

    Function ShowFieldAttributes(ByVal cNode As TreeNode) As Boolean

        If cNode Is Nothing Then
            Exit Function
        End If

        lvFieldAttrs.BeginUpdate()
        lvFieldAttrs.GridLines = True
        lvFieldAttrs.Items.Clear()

        lvFieldAttrs.Items.Add(modDeclares.TXT_NCHILDREN).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_NCHILDREN))
        lvFieldAttrs.Items.Add(modDeclares.TXT_LEVEL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LEVEL))
        lvFieldAttrs.Items.Add(modDeclares.TXT_TIMES).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_TIMES))
        lvFieldAttrs.Items.Add(modDeclares.TXT_OCCURS).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OCCURS))
        lvFieldAttrs.Items.Add(modDeclares.TXT_DATATYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATATYPE))
        lvFieldAttrs.Items.Add(modDeclares.TXT_OFFSET).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OFFSET))
        lvFieldAttrs.Items.Add(modDeclares.TXT_LENGTH).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LENGTH))
        lvFieldAttrs.Items.Add(modDeclares.TXT_SCALE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_SCALE))
        lvFieldAttrs.Items.Add(modDeclares.TXT_CANNULL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_CANNULL))
        lvFieldAttrs.Items.Add(modDeclares.TXT_ISKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
        lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY)) '// added by TK and KS 11/6/2006


        'lvFieldAttrs.Items.Add(modDeclares.TXT_ISKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
        'lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY)) '// added by TK and KS 11/6/2006
        lvFieldAttrs.Items.Add(modDeclares.TXT_INITVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INITVAL))
        lvFieldAttrs.Items.Add(modDeclares.TXT_RETYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_RETYPE))

        lvFieldAttrs.Items.Add(modDeclares.TXT_EXTTYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_EXTTYPE))
        lvFieldAttrs.Items.Add(modDeclares.TXT_INVALID).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INVALID))

        lvFieldAttrs.Items.Add(modDeclares.TXT_DATEFORMAT).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATEFORMAT))
        lvFieldAttrs.Items.Add(modDeclares.TXT_LABEL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LABEL))
        lvFieldAttrs.Items.Add(modDeclares.TXT_IDENTVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_IDENTVAL))

        Dim f As New System.Drawing.Font(lvFieldAttrs.Font, FontStyle.Bold)
        Dim i As Integer

        For i = 0 To lvFieldAttrs.Items.Count - 1
            lvFieldAttrs.Items(i).UseItemStyleForSubItems = False
            lvFieldAttrs.Items(i).SubItems(0).Font = f
            lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Gray
            lvFieldAttrs.Items(i).SubItems(1).ForeColor = Color.Gray
            lvFieldAttrs.Items(i).ImageIndex = 30
        Next

        'lblFieldName.Text = cNode.Text
        'lblFieldNameDesc.Text = lblFieldName.Text
        prevFld = cNode.Tag
        IsEventFromCode = True
        txtFieldDesc.Text = prevFld.FieldDesc
        If txtFieldDesc.Text <> "" Then
            cNode.ForeColor = Color.Blue
        Else
            cNode.ForeColor = Color.Black
        End If
        IsEventFromCode = False

        lvFieldAttrs.EndUpdate()

    End Function

    Function FillSegments() As Boolean

        Dim outXMLPath As String
        Dim NodeCol As TreeNodeCollection
        Dim flag As Boolean = False
        Dim filePath As String
        Dim filePre As String
        Dim preIDX As Integer
        Dim pathIDX As Integer

        Try
            If txtDBDFilePath.Text.Contains("%") = True Then
                preIDX = txtDBDFilePath.Text.IndexOf("%")
                pathIDX = txtDBDFilePath.Text.LastIndexOf("%")
                filePre = Strings.Mid(txtDBDFilePath.Text, preIDX + 2, pathIDX - preIDX - 1)
                filePath = Strings.Right(txtDBDFilePath.Text, txtDBDFilePath.Text.Length - pathIDX - 1)
                filePath = System.Environment.GetEnvironmentVariable(filePre) & filePath
            Else
                filePath = txtDBDFilePath.Text
            End If

            If IsNewObj = True Then
                outXMLPath = GetSQDumpXML(filePath, modDeclares.enumStructure.STRUCT_IMS)
                flag = LoadTreeViewFromXmlFile(objThis, outXMLPath, tvSegments)
                If IO.File.Exists(outXMLPath) Then
                    IO.File.Delete(outXMLPath)
                End If
            Else
                If modDeclares.enumStructure.STRUCT_COBOL_IMS Then 'objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or 
                    outXMLPath = GetSQDumpXML(filePath, modDeclares.enumStructure.STRUCT_IMS, objThis.SegmentName)
                    LoadTreeViewFromXmlFile(objThis, outXMLPath, tvSegments)
                    NodeCol = tvSegments.Nodes
                    '/// recurse all nodes in tree to hilite selected segment
                    flag = HiliteSeg(NodeCol)
                End If
            End If
            Return flag

        Catch ex As Exception
            LogError(ex, "ctlStructure FillSegments")
            Return False
        End Try

    End Function

    Function HiliteSeg(ByVal nodecol As TreeNodeCollection) As Boolean

        Dim Col As TreeNodeCollection

        Try
            For Each node As TreeNode In nodecol
                If node.Text = objThis.SegmentName Then
                    node.BackColor = Color.LightBlue
                    node.EnsureVisible()
                    Exit For
                Else
                    If node.Nodes.Count > 0 Then
                        Col = node.Nodes
                        Call HiliteSeg(Col)
                    End If
                End If
            Next
            Return True

        Catch ex As Exception
            LogError(ex, "ctlStructure HiLiteSeg")
            Return False
        End Try

    End Function

    Private Sub On_ctlStructure_resize(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Resize

        scStr.SplitterDistance = Split

    End Sub

    '/// Added by TK 12/07
    Private Sub OnlvAttr_resize(ByVal sender As Object, ByVal e As EventArgs) Handles lvFieldAttrs.Resize

        lvFieldAttrs.Columns(0).Width = (lvFieldAttrs.Width / 2) - 12
        lvFieldAttrs.Columns(1).Width = (lvFieldAttrs.Width / 2) - 12

    End Sub

    Private Sub FillRecordLength()

        Try
            Dim RecLth As Integer = 0

            For Each fld As clsField In objThis.ObjFields
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) <> "GROUPITEM" Then
                    RecLth = RecLth + CType(fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH), Integer)
                End If
            Next

            txtLength.Text = RecLth.ToString

        Catch ex As Exception
            LogError(ex, "FillRecordLength ctlStructure")
        End Try

    End Sub

#End Region

#Region "Click and Help Events"

    Public Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Save()

    End Sub

    Public Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click

        Dim ret As MsgBoxResult

        Try
            If objThis.IsModified = True Then
                ret = MsgBox("Do you want to save change(s) made to the opened form?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel, MsgTitle)
                If ret = MsgBoxResult.Yes Then
                    Save()
                ElseIf ret = MsgBoxResult.No Then
                    objThis.IsModified = False
                ElseIf ret = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Me.Visible = False
            RaiseEvent Closed(Me, objThis)

        Catch ex As Exception
            LogError(ex)
        End Try

    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click

        Dim strFilter As String

        Select Case objThis.StructureType
            Case modDeclares.enumStructure.STRUCT_C
                strFilter = "C Header files (*.h)|*.h|All files (*.*)|*.*"
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                strFilter = "Relational DDL files (*.ddl)|*.ddl|All files (*.*)|*.*"
            Case modDeclares.enumStructure.STRUCT_REL_DML
                strFilter = "Relational DML files (*.dml)|*.dml|All files (*.*)|*.*"
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                strFilter = "XML DTD Files (*.dtd)|*.dtd|All files (*.*)|*.*"
            Case Else
                strFilter = "All files (*.*)|*.*"
        End Select

        dlgOpen.Filter = strFilter
        dlgOpen.FileName = txtFilePath.Text
        dlgOpen.Multiselect = False

        If dlgOpen.ShowDialog() = DialogResult.OK Then
            'txtFilePath.Text = dlgOpen.FileName
            UpdateControls()
        End If

    End Sub

    Private Sub cmdBrowseCobolFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseCobolFile.Click

        dlgOpen.Filter = "Cobol Copybook (*.cob)|*.cob|All files (*.*)|*.*"
        dlgOpen.FileName = txtCobolFilePath.Text
        dlgOpen.Multiselect = False

        If dlgOpen.ShowDialog() = DialogResult.OK Then
            'txtCobolFilePath.Text = dlgOpen.FileName

            If (chkAddDBD.Checked = True And txtDBDFilePath.Text <> "") Then
                objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
            Else
                objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL
            End If
            UpdateControls()
        End If

    End Sub

    Private Sub cmdBrowseDBDFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseDBDFile.Click

        dlgOpen.Filter = "IMS DBD files (*.dbd)|*.dbd|All files (*.*)|*.*"
        dlgOpen.FileName = txtDBDFilePath.Text
        dlgOpen.Multiselect = False

        If dlgOpen.ShowDialog() = DialogResult.OK Then
            'txtDBDFilePath.Text = dlgOpen.FileName

            If (txtCobolFilePath.Text <> "") Then
                objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
            Else
                objThis.StructureType = modDeclares.enumStructure.STRUCT_IMS
            End If
            '//As soon user select *.dbd file display available segments in the file
            Call FillSegments()
            UpdateControls()
        End If

    End Sub

    Private Sub cmdShowHideFieldAttr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If IsNewObj = True Then
            If ValidateSelection() = False Then Exit Sub
        End If

        '//Dont reload everytime once its loaded from database
        If tvFields.GetNodeCount(True) > 0 And IsNewObj = False Then Exit Sub

        tvFields.BackColor = Color.LightBlue
        lvFieldAttrs.BackColor = Color.LightBlue

        'lblFieldName.ForeColor = Color.Red
        'lblFieldName.Text = "********* Loading *********"
        Me.Refresh()

        UpdateControls()

        Call FillFieldAttr()

        tvFields.BackColor = Color.White
        lvFieldAttrs.BackColor = Color.White
        'lblFieldName.Text = ""
        'lblFieldName.ForeColor = Color.LightSkyBlue
        If tvFields.GetNodeCount(True) > 0 Then
            tvFields.SelectedNode = tvFields.Nodes.Item(0)
        End If

    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Dim HelpObj As String

        HelpObj = Me.objThis.ObjTreeNode.Parent.Text

        Select Case HelpObj
            Case "COBOL"
                ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
            Case "CHeader"
                ShowHelp(modHelp.HHId.H_C_Structure)
            Case "COBOLIMS"
                ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
            Case "IMS"
                ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
            Case "DDL"
                ShowHelp(modHelp.HHId.H_Relational_DDL)
            Case "DML"
                ShowHelp(modHelp.HHId.H_Relational_DML)
            Case "XMLDTD"
                ShowHelp(modHelp.HHId.H_XML_DTD)
            Case "Unknown"
                ShowHelp(modHelp.HHId.H_Structures)
            Case Else
                ShowHelp(modHelp.HHId.H_Structures)
        End Select

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles chkAddDBD.KeyDown, lvFieldAttrs.KeyDown, tvFields.KeyDown, tvSegments.KeyDown, txtCobolFilePath.KeyDown, txtDBDFilePath.KeyDown, txtFilePath.KeyDown, txtSearchField.KeyDown, txtStructName.KeyDown, txtStructDesc.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    Public Sub fields_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lvFieldAttrs.KeyDown, tvFields.KeyDown, txtSearchField.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                ShowHelp(modHelp.HHId.H_Structure_Field_Attributes)
        End Select
    End Sub

    Private Sub AddToSourceMappingTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToSourceMappingTableToolStripMenuItem.Click

        If tvFields.SelectedNode IsNot Nothing Then
            Dim frm As New frmMapList
            Dim arrlist As New ArrayList
            arrlist.Clear()
            arrlist.Add(tvFields.SelectedNode.Text)
            Call frm.NewOrOpen(objThis.Project, arrlist)
        End If

    End Sub

    Private Sub AddToTargetMappingTableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddToTargetMappingTableToolStripMenuItem.Click

        If tvFields.SelectedNode IsNot Nothing Then
            Dim frm As New frmMapList
            Dim arrlist As New ArrayList
            arrlist.Clear()
            arrlist.Add(tvFields.SelectedNode.Text)
            Call frm.NewOrOpen(objThis.Project, , arrlist)
        End If

    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        If modGeneral.SelectFirstMatchingNode(tvFields, colSkipNodes, txtSearchField.Text) = False Then
            MsgBox("No matching node found for entered text", MsgBoxStyle.Critical, MsgTitle)
        End If

    End Sub

    Private Sub tvSegments_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvSegments.KeyDown

        If tvSegments.SelectedNode Is Nothing Then Exit Sub

        If e.KeyCode = Keys.Enter Then
            cmdShowHideFieldAttr_Click(Me, New EventArgs)
        End If

        If e.KeyCode = Keys.F1 Then
            Dim HelpObj As String

            HelpObj = Me.objThis.ObjTreeNode.Parent.Text
            If e.KeyCode = Keys.F1 Then
                Select Case HelpObj
                    Case "COBOL"
                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                    Case "CHeader"
                        ShowHelp(modHelp.HHId.H_C_Structure)
                    Case "COBOLIMS"
                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                    Case "IMS"
                        ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                    Case "DDL"
                        ShowHelp(modHelp.HHId.H_Relational_DDL)
                    Case "DML"
                        ShowHelp(modHelp.HHId.H_Relational_DML)
                    Case "XMLDTD"
                        ShowHelp(modHelp.HHId.H_XML_DTD)
                    Case "Unknown"
                        ShowHelp(modHelp.HHId.H_Structures)
                    Case Else
                        ShowHelp(modHelp.HHId.H_Structures)
                End Select
            End If
        End If

    End Sub

#End Region

#Region "Text and Combo Box Events"

    Private Sub chkAddDBD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddDBD.CheckedChanged

        UpdateControls()

    End Sub

    Private Sub txtFieldDesc_lostfocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFieldDesc.LostFocus, txtFieldDesc.MouseLeave, txtFieldDesc.Leave

        If prevFld IsNot Nothing Then
            prevFld.FieldDesc = txtFieldDesc.Text
        End If

    End Sub

    Private Sub txtSearchField_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchField.TextChanged

        modGeneral.SelectFirstMatchingNode(tvFields, txtSearchField.Text)

    End Sub

#End Region

#Region "Treeview and Listview Functions"

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        If prevFld IsNot Nothing Then
            If prevFld.FieldDescModified = True Then
                prevFld.FieldDesc = txtFieldDesc.Text
            End If
        End If
        e.Node.SelectedImageIndex = e.Node.ImageIndex
        HiLiteFieldDescNodes(tvFields.TopNode, True, tvFields)
        HiLiteFieldKeyNodes(tvFields.TopNode, True, tvFields)
        IsEventFromCode = True
        txtFieldDesc.Text = ""
        IsEventFromCode = False
        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub tvFields_Afterclick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvFields.NodeMouseClick, tvFields.NodeMouseDoubleClick

        If prevFld IsNot Nothing Then
            If prevFld.FieldDescModified = True Then
                prevFld.FieldDesc = txtFieldDesc.Text
            End If
        End If
        HiLiteFieldDescNodes(tvFields.TopNode, True, tvFields)
        HiLiteFieldKeyNodes(tvFields.TopNode, True, tvFields)
        IsEventFromCode = True
        txtFieldDesc.Text = ""
        IsEventFromCode = False
        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub tvSegments_AfterSelect(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvSegments.AfterSelect

        cmdShowHideFieldAttr_Click(Me, New EventArgs)

    End Sub

    Private Sub tvSegments_AfterCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvSegments.AfterCheck
        '//Only process this logic if action using mouse or keyboard other wise this event can fire on and on ....
        If IsEventFromCode = False Then
            OnChange(sender, New EventArgs)
            CheckUncheckNodes(e.Node)
        End If

    End Sub

#End Region

#Region "Object Functions"

    Public Function Save() As Boolean

        Dim i As Integer = 0
        Dim objfield As clsField

        Try
            If ValidateSelection() = False Then
                Exit Function
            End If
            '// First Check Validity before Saving
            If ValidateNewName128(txtStructName.Text) = False Then
                Exit Function
            End If
            If objThis.StructureName <> txtStructName.Text Then
                objThis.IsRenamed = RenameStructure(objThis, txtStructName.Text)
            End If
            If objThis.IsRenamed = False Then
                txtStructName.Text = objThis.StructureName
            Else
                objThis.StructureName = txtStructName.Text
            End If

            objThis.StructureDescription = txtStructDesc.Text '//new : 8/15/05
            objThis.Tablespace = txtTablespace.Text  '//new : 1/6/2013

            '//Make all file path case-sensitive : Unix uses case sensitive names so if user enters c:\wrseg.cob and on disk name appears WRSeg.cob then correct it to WRSeg.cob
            If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                objThis.fPath1 = GetCaseSensetivePath(txtCobolFilePath.Text)
                objThis.fPath2 = GetCaseSensetivePath(txtDBDFilePath.Text)
            Else
                objThis.fPath1 = GetCaseSensetivePath(txtFilePath.Text)
            End If
            'If objThis.Connection Is Nothing Then
            '    objThis.Connection.Text = ""
            'End If

            If IsNewObj = True Then
                '//Before we add new structure set structure fields
                LoadFldArrFromXmlFile(GetOutDumpXML, objThis.ObjFields)

                If objThis.AddNew = False Then Exit Function
            Else
                If objThis.IsRenamed = False Then
                    objThis.Save()
                End If
            End If

            For i = 0 To objThis.ObjFields.Count - 1
                objfield = objThis.ObjFields(i)
                If objfield.FieldDescModified = True Then
                    objfield.UpdateFieldDesc()
                End If
            Next

            Save = True

            cmdSave.Enabled = False
            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If
            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Public Function EditObj(ByVal obj As INode) As clsStructure

        Try
            IsNewObj = False
            StartLoad()
            objThis = obj '//Load the form struct object
            objThis.LoadMe()
            objThis.LoadItems()

            Call UpdateFields()

            If objThis.StructureType = enumStructure.STRUCT_REL_DML Or objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                If objThis.Connection IsNot Nothing Then
                    txtConn.Text = objThis.Connection.Text
                End If
            End If

            ShowHideControls(objThis.StructureType)

            FillRecordLength()

            EndLoad()

            EditObj = objThis

        Catch ex As Exception
            LogError(ex, "ctlStructure EditObj")
            EditObj = Nothing
        End Try

    End Function

    '//On Edit call this function
    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        Dim flag As Boolean = False
        Try
            txtStructName.Text = objThis.StructureName
            txtStructDesc.Text = objThis.StructureDescription
            txtTablespace.Text = objThis.Tablespace

            If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Then
                If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Then
                    txtCobolFilePath.Text = objThis.fPath1
                    flag = True
                Else
                    txtDBDFilePath.Enabled = True
                    chkAddDBD.Checked = True
                    'displaying only file name instead of the total path
                    txtCobolFilePath.Text = objThis.fPath1
                    txtDBDFilePath.Text = objThis.fPath2
                    flag = FillSegments()
                End If
            Else
                If objThis.StructureType = enumStructure.STRUCT_REL_DML Then
                    txtFilePath.Text = objThis.fPath1
                Else
                    txtFilePath.Text = objThis.fPath1
                End If
                flag = True
            End If
            Return flag

        Catch ex As Exception
            LogError(ex, "ctlStructure UpdateFields")
            Return False
        End Try

    End Function

    Function TrunkFileName(ByVal filePath As String) As String

        Dim lastSlash As Integer = filePath.LastIndexOf("\")
        Dim i As Integer
        Dim retFilePath As String

        retFilePath = Nothing
        Try
            For i = lastSlash To filePath.Length - 1
                If filePath.Chars(i) = "." Then
                    retFilePath = filePath.Substring(lastSlash)
                End If
            Next
            Return (retFilePath)
        Catch ex As Exception
            Return (retFilePath)
        End Try

    End Function

    Function ValidateSelection() As Boolean

        ValidateSelection = True

        If IsNewObj = False Then Exit Function '//new : 8/15/05

        If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then

            If txtCobolFilePath.Text.Trim = "" Then
                MsgBox("Please specify a cobol file (*.cob) path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(txtCobolFilePath.Text) = False Then
                MsgBox("Can not find the specified file " & txtCobolFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If chkAddDBD.Checked = True Then
                If IO.File.Exists(txtDBDFilePath.Text) = False Then
                    MsgBox("Can not find the specified file " & txtDBDFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                    ValidateSelection = False
                    Exit Function
                End If
                If txtDBDFilePath.Text.Trim = "" Then
                    MsgBox("Please specify a segment file (*.dbd) path.", MsgBoxStyle.Critical, MsgTitle)
                    ValidateSelection = False
                    Exit Function
                End If

                If tvSegments.GetNodeCount(True) > 0 Then
                    If (tvSegments.SelectedNode Is Nothing) Then
                        MsgBox("Please select a segment from the treeview", MsgBoxStyle.Exclamation, MsgTitle)
                        ValidateSelection = False
                        Exit Function
                    End If
                Else
                    MsgBox("No segment found in the selected DBD file", MsgBoxStyle.Critical, MsgTitle)
                    ValidateSelection = False
                    Exit Function
                End If
            End If
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Then
            If txtCobolFilePath.Text.Trim = "" Then
                MsgBox("Please specify a cobol file (*.cob) path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(txtCobolFilePath.Text) = False Then
                MsgBox("Can not find the specified file " & txtCobolFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_IMS Then
            If txtDBDFilePath.Text.Trim = "" Then
                MsgBox("Please specify a segment file (*.dbd) path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(txtDBDFilePath.Text) = False Then
                MsgBox("Can not find the specified file " & txtDBDFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If
        Else
            If txtFilePath.Text.Trim = "" Then
                MsgBox("Please select a valid file.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(txtFilePath.Text) = False Then
                MsgBox("Can not find the specified file " & txtFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If
        End If
    End Function

    Function GetOutDumpXML() As String

        If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Then
            GetOutDumpXML = GetSQDumpXML(txtCobolFilePath.Text, objThis.StructureType)
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
            GetOutDumpXML = GetSQDumpXML(txtDBDFilePath.Text, objThis.StructureType, tvSegments.SelectedNode.Text, Quote(txtCobolFilePath.Text, """"))
        Else
            GetOutDumpXML = GetSQDumpXML(txtFilePath.Text, objThis.StructureType)
        End If

    End Function

#End Region

    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click

        Dim filePath As String
        Dim filePre As String
        Dim preIDX As Integer
        Dim pathIDX As Integer

        Try
            If txtFilePath.Text.Contains("%") = True Then
                preIDX = txtFilePath.Text.IndexOf("%")
                pathIDX = txtFilePath.Text.LastIndexOf("%")
                filePre = Strings.Mid(txtFilePath.Text, preIDX + 2, pathIDX - preIDX - 1)
                filePath = Strings.Right(txtFilePath.Text, txtFilePath.Text.Length - pathIDX - 1)
                filePath = System.Environment.GetEnvironmentVariable(filePre) & filePath
            Else
                filePath = txtFilePath.Text
            End If

            If System.IO.File.Exists(filePath) = True Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(filePath)
            End If

        Catch ex As Exception
            LogError(ex, "ctlStructure btnView_click")
        End Try

    End Sub

    Private Sub btnViewCobol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewCobol.Click

        Dim filePath As String
        Dim filePre As String
        Dim preIDX As Integer
        Dim pathIDX As Integer

        Try
            If txtCobolFilePath.Text.Contains("%") = True Then
                preIDX = txtCobolFilePath.Text.IndexOf("%")
                pathIDX = txtCobolFilePath.Text.LastIndexOf("%")
                filePre = Strings.Mid(txtCobolFilePath.Text, preIDX + 2, pathIDX - preIDX - 1)
                filePath = Strings.Right(txtCobolFilePath.Text, txtCobolFilePath.Text.Length - pathIDX - 1)
                filePath = System.Environment.GetEnvironmentVariable(filePre) & filePath
            Else
                filePath = txtCobolFilePath.Text
            End If

            If System.IO.File.Exists(filePath) = True Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(filePath)
            End If

        Catch ex As Exception
            LogError(ex, "ctlStructure btnViewCobol_click")
        End Try

    End Sub

    Private Sub btnViewDBD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewDBD.Click

        Dim filePath As String
        Dim filePre As String
        Dim preIDX As Integer
        Dim pathIDX As Integer

        Try
            If txtDBDFilePath.Text.Contains("%") = True Then
                preIDX = txtDBDFilePath.Text.IndexOf("%")
                pathIDX = txtDBDFilePath.Text.LastIndexOf("%")
                filePre = Strings.Mid(txtDBDFilePath.Text, preIDX + 2, pathIDX - preIDX - 1)
                filePath = Strings.Right(txtDBDFilePath.Text, txtDBDFilePath.Text.Length - pathIDX - 1)
                filePath = System.Environment.GetEnvironmentVariable(filePre) & filePath
            Else
                filePath = txtDBDFilePath.Text
            End If

            If System.IO.File.Exists(filePath) = True Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(filePath)
            End If

        Catch ex As Exception
            LogError(ex, "ctlStructure btnViewDBD_click")
        End Try
    End Sub

End Class
