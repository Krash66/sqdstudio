Public Class ctlStructureSelection
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)
    
    Dim objThis As New clsStructureSelection
    Dim IsNewObj As Boolean
    Dim lastTreeviewSearchText As String
    Dim lastTreeviewSearchNode As TreeNode
    Dim prevFld As clsField
    Private Const Split As Integer = 128
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
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents txtSearchField As System.Windows.Forms.TextBox
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtSelectionDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtSelectionName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtFieldDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents A4 As System.Windows.Forms.Label
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents scSS As System.Windows.Forms.SplitContainer
    Friend WithEvents gbProp As System.Windows.Forms.GroupBox
    Friend WithEvents gbFldDesc As System.Windows.Forms.GroupBox
    Friend WithEvents scField As System.Windows.Forms.SplitContainer
    Friend WithEvents gbFields As System.Windows.Forms.GroupBox
    Friend WithEvents scFldDesc As System.Windows.Forms.SplitContainer
    Friend WithEvents gbAttr As System.Windows.Forms.GroupBox
    Friend WithEvents gbSearch As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtSelectionDesc = New System.Windows.Forms.TextBox
        Me.txtSelectionName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.A4 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtFieldDesc = New System.Windows.Forms.TextBox
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.txtSearchField = New System.Windows.Forms.TextBox
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.scSS = New System.Windows.Forms.SplitContainer
        Me.gbProp = New System.Windows.Forms.GroupBox
        Me.scField = New System.Windows.Forms.SplitContainer
        Me.gbFields = New System.Windows.Forms.GroupBox
        Me.gbSearch = New System.Windows.Forms.GroupBox
        Me.scFldDesc = New System.Windows.Forms.SplitContainer
        Me.gbAttr = New System.Windows.Forms.GroupBox
        Me.gbFldDesc = New System.Windows.Forms.GroupBox
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.scSS.Panel1.SuspendLayout()
        Me.scSS.Panel2.SuspendLayout()
        Me.scSS.SuspendLayout()
        Me.gbProp.SuspendLayout()
        Me.scField.Panel1.SuspendLayout()
        Me.scField.Panel2.SuspendLayout()
        Me.scField.SuspendLayout()
        Me.gbFields.SuspendLayout()
        Me.gbSearch.SuspendLayout()
        Me.scFldDesc.Panel1.SuspendLayout()
        Me.scFldDesc.Panel2.SuspendLayout()
        Me.scFldDesc.SuspendLayout()
        Me.gbAttr.SuspendLayout()
        Me.gbFldDesc.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSelectionDesc
        '
        Me.txtSelectionDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtSelectionDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSelectionDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtSelectionDesc.Multiline = True
        Me.txtSelectionDesc.Name = "txtSelectionDesc"
        Me.txtSelectionDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSelectionDesc.Size = New System.Drawing.Size(306, 52)
        Me.txtSelectionDesc.TabIndex = 2
        '
        'txtSelectionName
        '
        Me.txtSelectionName.Location = New System.Drawing.Point(50, 19)
        Me.txtSelectionName.MaxLength = 128
        Me.txtSelectionName.Name = "txtSelectionName"
        Me.txtSelectionName.Size = New System.Drawing.Size(265, 20)
        Me.txtSelectionName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 67
        Me.Label3.Text = "Name"
        '
        'A4
        '
        Me.A4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.A4.BackColor = System.Drawing.Color.White
        Me.A4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.A4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A4.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.A4.Location = New System.Drawing.Point(8, 571)
        Me.A4.Name = "A4"
        Me.A4.Size = New System.Drawing.Size(16, 16)
        Me.A4.TabIndex = 88
        Me.A4.Text = "A"
        Me.A4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label6.Location = New System.Drawing.Point(30, 572)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(139, 13)
        Me.Label6.TabIndex = 87
        Me.Label6.Text = "Has a Field Description"
        '
        'txtFieldDesc
        '
        Me.txtFieldDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtFieldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtFieldDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFieldDesc.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.txtFieldDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtFieldDesc.Multiline = True
        Me.txtFieldDesc.Name = "txtFieldDesc"
        Me.txtFieldDesc.Size = New System.Drawing.Size(478, 55)
        Me.txtFieldDesc.TabIndex = 34
        '
        'cmdSearch
        '
        Me.cmdSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSearch.Location = New System.Drawing.Point(197, 19)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(30, 20)
        Me.cmdSearch.TabIndex = 4
        Me.cmdSearch.Text = "Go"
        '
        'txtSearchField
        '
        Me.txtSearchField.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSearchField.Location = New System.Drawing.Point(6, 19)
        Me.txtSearchField.MaxLength = 128
        Me.txtSearchField.Name = "txtSearchField"
        Me.txtSearchField.Size = New System.Drawing.Size(185, 20)
        Me.txtSearchField.TabIndex = 3
        '
        'tvFields
        '
        Me.tvFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvFields.HideSelection = False
        Me.tvFields.HotTracking = True
        Me.tvFields.Location = New System.Drawing.Point(6, 16)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.Size = New System.Drawing.Size(233, 318)
        Me.tvFields.TabIndex = 32
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
        Me.lvFieldAttrs.Size = New System.Drawing.Size(478, 295)
        Me.lvFieldAttrs.TabIndex = 33
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
        Me.ColumnHeader2.Width = 222
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(663, 566)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 42
        Me.cmdHelp.Text = "&Help"
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.Enabled = False
        Me.cmdSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(505, 566)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 40
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(585, 566)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 41
        Me.cmdClose.Text = "&Close"
        Me.cmdClose.UseVisualStyleBackColor = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 550)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(729, 7)
        Me.GroupBox1.TabIndex = 38
        Me.GroupBox1.TabStop = False
        '
        'gbName
        '
        Me.gbName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Controls.Add(Me.txtSelectionName)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.Color.White
        Me.gbName.Location = New System.Drawing.Point(3, 3)
        Me.gbName.MaximumSize = New System.Drawing.Size(324, 122)
        Me.gbName.MinimumSize = New System.Drawing.Size(324, 122)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(324, 122)
        Me.gbName.TabIndex = 43
        Me.gbName.TabStop = False
        Me.gbName.Text = "Description Selection Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtSelectionDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 45)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(312, 71)
        Me.gbDesc.TabIndex = 68
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'scSS
        '
        Me.scSS.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scSS.Location = New System.Drawing.Point(3, 3)
        Me.scSS.Name = "scSS"
        Me.scSS.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scSS.Panel1
        '
        Me.scSS.Panel1.Controls.Add(Me.gbName)
        '
        'scSS.Panel2
        '
        Me.scSS.Panel2.Controls.Add(Me.gbProp)
        Me.scSS.Size = New System.Drawing.Size(739, 541)
        Me.scSS.SplitterDistance = 126
        Me.scSS.TabIndex = 44
        '
        'gbProp
        '
        Me.gbProp.Controls.Add(Me.scField)
        Me.gbProp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProp.ForeColor = System.Drawing.Color.White
        Me.gbProp.Location = New System.Drawing.Point(0, 0)
        Me.gbProp.Name = "gbProp"
        Me.gbProp.Size = New System.Drawing.Size(739, 411)
        Me.gbProp.TabIndex = 0
        Me.gbProp.TabStop = False
        Me.gbProp.Text = "Description Field Properties"
        '
        'scField
        '
        Me.scField.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scField.Location = New System.Drawing.Point(3, 16)
        Me.scField.Name = "scField"
        '
        'scField.Panel1
        '
        Me.scField.Panel1.Controls.Add(Me.gbFields)
        '
        'scField.Panel2
        '
        Me.scField.Panel2.Controls.Add(Me.scFldDesc)
        Me.scField.Size = New System.Drawing.Size(733, 392)
        Me.scField.SplitterDistance = 245
        Me.scField.TabIndex = 61
        '
        'gbFields
        '
        Me.gbFields.Controls.Add(Me.gbSearch)
        Me.gbFields.Controls.Add(Me.tvFields)
        Me.gbFields.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFields.ForeColor = System.Drawing.Color.White
        Me.gbFields.Location = New System.Drawing.Point(0, 0)
        Me.gbFields.Name = "gbFields"
        Me.gbFields.Size = New System.Drawing.Size(245, 392)
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
        Me.gbSearch.Location = New System.Drawing.Point(6, 340)
        Me.gbSearch.Name = "gbSearch"
        Me.gbSearch.Size = New System.Drawing.Size(233, 46)
        Me.gbSearch.TabIndex = 0
        Me.gbSearch.TabStop = False
        Me.gbSearch.Text = "Field Search"
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
        Me.scFldDesc.Size = New System.Drawing.Size(484, 392)
        Me.scFldDesc.SplitterDistance = 314
        Me.scFldDesc.TabIndex = 0
        '
        'gbAttr
        '
        Me.gbAttr.Controls.Add(Me.lvFieldAttrs)
        Me.gbAttr.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAttr.ForeColor = System.Drawing.Color.White
        Me.gbAttr.Location = New System.Drawing.Point(0, 0)
        Me.gbAttr.Name = "gbAttr"
        Me.gbAttr.Size = New System.Drawing.Size(484, 314)
        Me.gbAttr.TabIndex = 0
        Me.gbAttr.TabStop = False
        Me.gbAttr.Text = "Field Attributes"
        '
        'gbFldDesc
        '
        Me.gbFldDesc.Controls.Add(Me.txtFieldDesc)
        Me.gbFldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFldDesc.ForeColor = System.Drawing.Color.White
        Me.gbFldDesc.Location = New System.Drawing.Point(0, 0)
        Me.gbFldDesc.Name = "gbFldDesc"
        Me.gbFldDesc.Size = New System.Drawing.Size(484, 74)
        Me.gbFldDesc.TabIndex = 1
        Me.gbFldDesc.TabStop = False
        Me.gbFldDesc.Text = "Field Description"
        '
        'ctlStructureSelection
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.A4)
        Me.Controls.Add(Me.scSS)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdHelp)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlStructureSelection"
        Me.Size = New System.Drawing.Size(745, 598)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.scSS.Panel1.ResumeLayout(False)
        Me.scSS.Panel2.ResumeLayout(False)
        Me.scSS.ResumeLayout(False)
        Me.gbProp.ResumeLayout(False)
        Me.scField.Panel1.ResumeLayout(False)
        Me.scField.Panel2.ResumeLayout(False)
        Me.scField.ResumeLayout(False)
        Me.gbFields.ResumeLayout(False)
        Me.gbSearch.ResumeLayout(False)
        Me.gbSearch.PerformLayout()
        Me.scFldDesc.Panel1.ResumeLayout(False)
        Me.scFldDesc.Panel2.ResumeLayout(False)
        Me.scFldDesc.ResumeLayout(False)
        Me.gbAttr.ResumeLayout(False)
        Me.gbFldDesc.ResumeLayout(False)
        Me.gbFldDesc.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Form Events"

    Private Sub StartLoad()
        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtSelectionName.Enabled = True '//added on 7/24 to disable object name editing

        InitControls()

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsStructureSelection

        ClearControls(Me.Controls)
    End Sub

    Private Sub EndLoad()
        Me.BringToFront()
        Me.Visible = True

        IsEventFromCode = False

        '//Now load fields and selection
        Call cmdShowHideFieldAttr_Click(Me, New EventArgs)
    End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSelectionDesc.TextChanged, txtSelectionName.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
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

    Function UpdateControls() As Boolean

    End Function

    Function InitControls() As Boolean

        '//Add listview columns
        tvFields.HideSelection = False
        tvFields.CheckBoxes = True
        lvFieldAttrs.SmallImageList = imgListSmall

        scSS.SplitterDistance = Split

        UpdateControls()

    End Function

    '//On Edit call this function
    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtSelectionName.Text = objThis.SelectionName
        txtSelectionDesc.Text = objThis.SelectionDescription

    End Function

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

    Private Sub cmdShowHideFieldAttr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        '//Dont reload and reprocess it once its loaded
        If tvFields.GetNodeCount(True) > 0 Then Exit Sub

        tvFields.BackColor = Color.LightBlue
        lvFieldAttrs.BackColor = Color.LightBlue

        'lblFieldName.ForeColor = Color.Red
        'lblFieldName.Text = "********* Loading *********"
        Me.Refresh()

        UpdateControls()

        '//FirstLoad all field of parent structure and then check only selected for this selection
        Call FillFieldAttr()

        tvFields.BackColor = Color.White
        lvFieldAttrs.BackColor = Color.White
        'lblFieldName.Text = ""
        'lblFieldName.ForeColor = Color.LightSkyBlue

        '//select first itm
        If tvFields.GetNodeCount(True) > 0 Then
            tvFields.SelectedNode = tvFields.Nodes.Item(0)
        End If
    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Dim HelpType As enumStructure = CType(objThis.Parent, clsStructure).StructureType
        Select Case HelpType
            Case enumStructure.STRUCT_COBOL
                ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
            Case enumStructure.STRUCT_C
                ShowHelp(modHelp.HHId.H_C_Structure)
            Case enumStructure.STRUCT_COBOL_IMS
                ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
            Case enumStructure.STRUCT_IMS
                ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
            Case enumStructure.STRUCT_REL_DDL
                ShowHelp(modHelp.HHId.H_Relational_DDL)
            Case enumStructure.STRUCT_REL_DML
                ShowHelp(modHelp.HHId.H_Relational_DML)
            Case enumStructure.STRUCT_XMLDTD
                ShowHelp(modHelp.HHId.H_XML_DTD)
            Case enumStructure.STRUCT_UNKNOWN
                ShowHelp(modHelp.HHId.H_Structures)
            Case Else
                ShowHelp(modHelp.HHId.H_Structures)
        End Select

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtSelectionDesc.KeyDown, txtSelectionName.KeyDown

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

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        If SelectFirstMatchingNode(tvFields, colSkipNodes, txtSearchField.Text) = False Then
            MsgBox("No matching node found for entered text", MsgBoxStyle.Critical, MsgTitle)
        End If

    End Sub

#End Region
   
#Region "Text and Combo Box Events"

    Private Sub txtSearchField_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchField.TextChanged

        SelectFirstMatchingNode(tvFields, txtSearchField.Text)

    End Sub

    Private Sub txtFieldDesc_lostfocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFieldDesc.LostFocus, txtFieldDesc.MouseLeave, txtFieldDesc.Leave

        prevFld.FieldDesc = txtFieldDesc.Text

    End Sub

#End Region
   
#Region "Treeview and Listview Events"

    '//This will load treeview of fields
    Function FillFieldAttr() As Boolean
        Try
            tvFields.BeginUpdate()
            lvFieldAttrs.BeginUpdate()
            tvFields.Nodes.Clear()

            objThis.ObjStructure.LoadItems()
            '//Load all fields from structure
            AddFieldsToTreeView(objThis.ObjStructure, , tvFields)

            '//Now load selected fields and check if object is already created 
            '//and user is editing it
            If IsNewObj = False Then objThis.LoadMe()
            CheckSelectedFields()
            HiLiteFieldDescNodes(tvFields.Nodes(0), True, tvFields)
            HiLiteFieldFKeyNodes(tvFields.Nodes(0), True, tvFields)

        Catch ex As Exception
            LogError(ex)
        Finally
            tvFields.EndUpdate()
            lvFieldAttrs.EndUpdate()
        End Try

    End Function

    Function CheckSelectedFields() As Boolean

        Dim i As Integer
        Dim cnt As Integer = tvFields.GetNodeCount(True)

        If cnt <= 0 Then
            cmdSave.Enabled = False
            Exit Function
        End If
        Try
            tvFields.BeginUpdate()
            '//Looop all selected field and check that field in treeview
            For i = 0 To objThis.ObjSelectionFields.Count - 1
                SelectFirstMatchingNode(tvFields, "", True, CType(objThis.ObjSelectionFields(i), clsField).Key)
            Next

        Catch ex As Exception
            LogError(ex)
        Finally
            tvFields.EndUpdate()
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

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        If prevFld IsNot Nothing Then
            If prevFld.FieldDescModified = True Then
                prevFld.FieldDesc = txtFieldDesc.Text
            End If
        End If
        HiLiteFieldDescNodes(tvFields.TopNode, True, tvFields)
        'HiLiteFieldFKeyNodes(tvFields.TopNode, True, tvFields)
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
        'HiLiteFieldFKeyNodes(tvFields.TopNode, True, tvFields)
        IsEventFromCode = True
        txtFieldDesc.Text = ""
        IsEventFromCode = False
        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub tvFields_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterCheck
        '//Only process this logic if action using mouse or keyboard other wise this event can fire on and on ....
        If IsEventFromCode = False Then
            OnChange(sender, e)
            CheckUncheckNodes(e.Node)
        End If

    End Sub

    Private Sub OnlvAttr_resize(ByVal sender As Object, ByVal e As EventArgs) Handles lvFieldAttrs.Resize

        lvFieldAttrs.Columns(0).Width = (lvFieldAttrs.Width / 2) - 12
        lvFieldAttrs.Columns(1).Width = (lvFieldAttrs.Width / 2) - 12

    End Sub

#End Region

#Region "Object Functions"

    Public Function Save() As Boolean

        Dim i As Integer = 0
        Dim objfield As clsField

        Try
            Me.Cursor = Cursors.WaitCursor
            '// First Check Validity before Saving
            If ValidateNewName128(txtSelectionName.Text) = False Then
                Save = False
                Me.Cursor = Cursors.Default
                Exit Function
            End If

            '//save current selected fields before we do add/save
            LoadFldArrFromTreeview(objThis.ObjSelectionFields, tvFields, True)

            If objThis.SelectionName <> txtSelectionName.Text Then
                objThis.IsRenamed = RenameStructureSelection(objThis, txtSelectionName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtSelectionName.Text = objThis.SelectionName
            Else
                objThis.SelectionName = txtSelectionName.Text
            End If

            objThis.SelectionDescription = txtSelectionDesc.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    Save = False
                    Me.Cursor = Cursors.Default
                    Exit Function
                End If
            Else
                If objThis.Save() = False Then
                    Save = False
                    Me.Cursor = Cursors.Default
                    Exit Function
                End If
            End If

            For i = 0 To objThis.ObjStructure.ObjFields.Count - 1
                objfield = objThis.ObjStructure.ObjFields(i)
                If objfield.FieldDescModified = True Then
                    objfield.UpdateFieldDesc()
                End If
            Next
            Me.Cursor = Cursors.Default

            Save = True

            cmdSave.Enabled = False
            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If
            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex, "ctlSS Save")
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Public Function EditObj(ByVal obj As INode) As clsStructureSelection

        Try
            Me.Cursor = Cursors.WaitCursor

            IsNewObj = False
            StartLoad()
            objThis = obj '//Load the form struct object
            objThis.LoadMe()

            UpdateFields()
            EndLoad()
            Me.Cursor = Cursors.Default

            EditObj = objThis

        Catch ex As Exception
            LogError(ex, "ctlStructureSelection EditObj")
            EditObj = Nothing
            Me.Cursor = Cursors.Default
        End Try

    End Function

#End Region

End Class
