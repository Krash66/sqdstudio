Public Class frmDatastore
    Inherits SQDStudio.frmBlank

    Public objThis As New clsDatastore
    Dim IsNewObj As Boolean
    Dim WithEvents sl As MyListener
    Dim IsLU As Boolean

    Private Const SplitSrc As Integer = 389
    Private Const SplitTgt As Integer = 308
    Private Const SplitSrcNoDel As Integer = 332
    Private Const SplitTgtNoDel As Integer = 251

    '/// Group box sizes plus 6 for gap between
    Private Const szProp As Integer = 102
    Private Const szExtProp As Integer = 74
    Private Const szSrc As Integer = 131
    Private Const szTgt As Integer = 50
    Private Const szDel As Integer = 63
    Private Const szFld As Integer = 151 'Actual size
    Dim KeyForRel As Boolean = True

    Private Pnt As Point
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList

    Private lvItem As ListViewItem

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cmbTextQualifier As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents cmbRowDelimiter As System.Windows.Forms.ComboBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbColDelimiter As System.Windows.Forms.ComboBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtDatastoreDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmbCharacterCode As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbAccessMethod As System.Windows.Forms.ComboBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtPhysicalSource As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents tvDatastoreStructures As System.Windows.Forms.TreeView
    Friend WithEvents txtDatastoreName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtException As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents chkSkipChangeCheck As System.Windows.Forms.CheckBox
    Friend WithEvents chkIMSPathData As System.Windows.Forms.CheckBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents cmbOperationType As System.Windows.Forms.ComboBox
    Friend WithEvents txtPortOrMQMgr As System.Windows.Forms.TextBox
    Friend WithEvents lblPortorMQMgr As System.Windows.Forms.Label
    Friend WithEvents cmbListViewCombo As System.Windows.Forms.ComboBox
    Friend WithEvents txtListViewText As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtUOW As System.Windows.Forms.TextBox
    Friend WithEvents lblUOW As System.Windows.Forms.Label
    Friend WithEvents cmbDatastoreType As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtPoll As System.Windows.Forms.TextBox
    Friend WithEvents lblPoll As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents rbLookup As System.Windows.Forms.RadioButton
    Friend WithEvents gbProp As System.Windows.Forms.GroupBox
    Friend WithEvents gbExtProps As System.Windows.Forms.GroupBox
    Friend WithEvents lblRestart As System.Windows.Forms.Label
    Friend WithEvents txtRestart As System.Windows.Forms.TextBox
    Friend WithEvents gbTarget As System.Windows.Forms.GroupBox
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cbIfNullChar As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents cbIfNullNum As System.Windows.Forms.ComboBox
    Friend WithEvents lblInValid As System.Windows.Forms.Label
    Friend WithEvents cbInvalidChar As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents cbInvalidNum As System.Windows.Forms.ComboBox
    Friend WithEvents lblExtType As System.Windows.Forms.Label
    Friend WithEvents cbIfSpaceChar As System.Windows.Forms.ComboBox
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents cbIfSpaceNum As System.Windows.Forms.ComboBox
    Friend WithEvents gbStruct As System.Windows.Forms.GroupBox
    Friend WithEvents gbDel As System.Windows.Forms.GroupBox
    Friend WithEvents gbField As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents gbFldname As System.Windows.Forms.GroupBox
    Friend WithEvents gbAtt As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDatastore))
        Me.txtListViewText = New System.Windows.Forms.TextBox
        Me.cmbListViewCombo = New System.Windows.Forms.ComboBox
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.rbLookup = New System.Windows.Forms.RadioButton
        Me.lblPoll = New System.Windows.Forms.Label
        Me.txtPoll = New System.Windows.Forms.TextBox
        Me.cmbDatastoreType = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtUOW = New System.Windows.Forms.TextBox
        Me.lblUOW = New System.Windows.Forms.Label
        Me.cmbOperationType = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtPortOrMQMgr = New System.Windows.Forms.TextBox
        Me.lblPortorMQMgr = New System.Windows.Forms.Label
        Me.txtException = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.chkSkipChangeCheck = New System.Windows.Forms.CheckBox
        Me.chkIMSPathData = New System.Windows.Forms.CheckBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtDatastoreDesc = New System.Windows.Forms.TextBox
        Me.cmbCharacterCode = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmbAccessMethod = New System.Windows.Forms.ComboBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPhysicalSource = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.tvDatastoreStructures = New System.Windows.Forms.TreeView
        Me.txtDatastoreName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.cmbTextQualifier = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmbRowDelimiter = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbColDelimiter = New System.Windows.Forms.ComboBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.gbProp = New System.Windows.Forms.GroupBox
        Me.gbExtProps = New System.Windows.Forms.GroupBox
        Me.txtRestart = New System.Windows.Forms.TextBox
        Me.lblRestart = New System.Windows.Forms.Label
        Me.gbTarget = New System.Windows.Forms.GroupBox
        Me.gbSource = New System.Windows.Forms.GroupBox
        Me.cbIfSpaceNum = New System.Windows.Forms.ComboBox
        Me.Label20 = New System.Windows.Forms.Label
        Me.cbIfSpaceChar = New System.Windows.Forms.ComboBox
        Me.lblExtType = New System.Windows.Forms.Label
        Me.cbInvalidNum = New System.Windows.Forms.ComboBox
        Me.Label18 = New System.Windows.Forms.Label
        Me.cbInvalidChar = New System.Windows.Forms.ComboBox
        Me.lblInValid = New System.Windows.Forms.Label
        Me.cbIfNullNum = New System.Windows.Forms.ComboBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.cbIfNullChar = New System.Windows.Forms.ComboBox
        Me.Label17 = New System.Windows.Forms.Label
        Me.gbStruct = New System.Windows.Forms.GroupBox
        Me.gbDel = New System.Windows.Forms.GroupBox
        Me.gbField = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.gbFldname = New System.Windows.Forms.GroupBox
        Me.gbAtt = New System.Windows.Forms.GroupBox
        Me.Panel1.SuspendLayout()
        Me.gbProp.SuspendLayout()
        Me.gbExtProps.SuspendLayout()
        Me.gbTarget.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.gbStruct.SuspendLayout()
        Me.gbDel.SuspendLayout()
        Me.gbField.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.gbFldname.SuspendLayout()
        Me.gbAtt.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Size = New System.Drawing.Size(704, 76)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(0, 609)
        Me.GroupBox1.Size = New System.Drawing.Size(712, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(432, 625)
        Me.cmdOk.TabIndex = 21
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(520, 625)
        Me.cmdCancel.TabIndex = 22
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(608, 625)
        Me.cmdHelp.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.Size = New System.Drawing.Size(187, 16)
        Me.Label1.Text = "Datastore definition"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(72, 21)
        Me.Label2.Size = New System.Drawing.Size(620, 40)
        Me.Label2.Text = resources.GetString("Label2.Text")
        '
        'txtListViewText
        '
        Me.txtListViewText.BackColor = System.Drawing.Color.Moccasin
        Me.txtListViewText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtListViewText.Font = New System.Drawing.Font("Tahoma", 7.0!)
        Me.txtListViewText.Location = New System.Drawing.Point(135, 54)
        Me.txtListViewText.MaxLength = 128
        Me.txtListViewText.Name = "txtListViewText"
        Me.txtListViewText.Size = New System.Drawing.Size(96, 19)
        Me.txtListViewText.TabIndex = 40
        Me.txtListViewText.Visible = False
        Me.txtListViewText.WordWrap = False
        '
        'cmbListViewCombo
        '
        Me.cmbListViewCombo.BackColor = System.Drawing.Color.Moccasin
        Me.cmbListViewCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbListViewCombo.Font = New System.Drawing.Font("Tahoma", 7.0!)
        Me.cmbListViewCombo.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmbListViewCombo.ItemHeight = 11
        Me.cmbListViewCombo.Location = New System.Drawing.Point(324, 54)
        Me.cmbListViewCombo.Name = "cmbListViewCombo"
        Me.cmbListViewCombo.Size = New System.Drawing.Size(128, 19)
        Me.cmbListViewCombo.TabIndex = 39
        Me.cmbListViewCombo.Visible = False
        '
        'tvFields
        '
        Me.tvFields.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvFields.ImageIndex = 0
        Me.tvFields.ImageList = Me.ImageList1
        Me.tvFields.Location = New System.Drawing.Point(3, 16)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.SelectedImageIndex = 0
        Me.tvFields.Size = New System.Drawing.Size(208, 113)
        Me.tvFields.TabIndex = 34
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
        Me.lvFieldAttrs.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvFieldAttrs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvFieldAttrs.Location = New System.Drawing.Point(3, 16)
        Me.lvFieldAttrs.Name = "lvFieldAttrs"
        Me.lvFieldAttrs.Size = New System.Drawing.Size(459, 113)
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
        Me.ColumnHeader2.Width = 200
        '
        'rbLookup
        '
        Me.rbLookup.AutoCheck = False
        Me.rbLookup.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLookup.ForeColor = System.Drawing.SystemColors.WindowText
        Me.rbLookup.Location = New System.Drawing.Point(334, 261)
        Me.rbLookup.Name = "rbLookup"
        Me.rbLookup.Size = New System.Drawing.Size(104, 39)
        Me.rbLookup.TabIndex = 140
        Me.rbLookup.TabStop = True
        Me.rbLookup.Text = "LOOKUP DATASTORE"
        Me.rbLookup.UseVisualStyleBackColor = True
        Me.rbLookup.Visible = False
        '
        'lblPoll
        '
        Me.lblPoll.AutoSize = True
        Me.lblPoll.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPoll.Location = New System.Drawing.Point(319, 42)
        Me.lblPoll.Name = "lblPoll"
        Me.lblPoll.Size = New System.Drawing.Size(28, 13)
        Me.lblPoll.TabIndex = 139
        Me.lblPoll.Text = "Poll"
        Me.lblPoll.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtPoll
        '
        Me.txtPoll.Location = New System.Drawing.Point(353, 39)
        Me.txtPoll.Name = "txtPoll"
        Me.txtPoll.Size = New System.Drawing.Size(75, 20)
        Me.txtPoll.TabIndex = 7
        Me.txtPoll.Text = "5"
        '
        'cmbDatastoreType
        '
        Me.cmbDatastoreType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDatastoreType.Location = New System.Drawing.Point(311, 16)
        Me.cmbDatastoreType.Name = "cmbDatastoreType"
        Me.cmbDatastoreType.Size = New System.Drawing.Size(118, 21)
        Me.cmbDatastoreType.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.cmbDatastoreType, "Type of Datastore" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Usually your Database type")
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(262, 19)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(34, 14)
        Me.Label15.TabIndex = 136
        Me.Label15.Text = "Type"
        '
        'txtUOW
        '
        Me.txtUOW.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUOW.Location = New System.Drawing.Point(282, 13)
        Me.txtUOW.MaxLength = 128
        Me.txtUOW.Name = "txtUOW"
        Me.txtUOW.Size = New System.Drawing.Size(146, 20)
        Me.txtUOW.TabIndex = 16
        '
        'lblUOW
        '
        Me.lblUOW.AutoSize = True
        Me.lblUOW.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUOW.Location = New System.Drawing.Point(244, 16)
        Me.lblUOW.Name = "lblUOW"
        Me.lblUOW.Size = New System.Drawing.Size(32, 14)
        Me.lblUOW.TabIndex = 133
        Me.lblUOW.Text = "UOW"
        Me.lblUOW.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbOperationType
        '
        Me.cmbOperationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOperationType.Location = New System.Drawing.Point(103, 13)
        Me.cmbOperationType.Name = "cmbOperationType"
        Me.cmbOperationType.Size = New System.Drawing.Size(216, 21)
        Me.cmbOperationType.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cmbOperationType, "Database operation to perform" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10))
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(6, 16)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(91, 16)
        Me.Label16.TabIndex = 125
        Me.Label16.Text = "Operation Type"
        '
        'txtPortOrMQMgr
        '
        Me.txtPortOrMQMgr.Location = New System.Drawing.Point(311, 68)
        Me.txtPortOrMQMgr.MaxLength = 128
        Me.txtPortOrMQMgr.Name = "txtPortOrMQMgr"
        Me.txtPortOrMQMgr.Size = New System.Drawing.Size(118, 20)
        Me.txtPortOrMQMgr.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.txtPortOrMQMgr, "Connection Port Number")
        '
        'lblPortorMQMgr
        '
        Me.lblPortorMQMgr.AutoSize = True
        Me.lblPortorMQMgr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPortorMQMgr.Location = New System.Drawing.Point(262, 71)
        Me.lblPortorMQMgr.Name = "lblPortorMQMgr"
        Me.lblPortorMQMgr.Size = New System.Drawing.Size(30, 14)
        Me.lblPortorMQMgr.TabIndex = 123
        Me.lblPortorMQMgr.Text = "Port"
        '
        'txtException
        '
        Me.txtException.Location = New System.Drawing.Point(68, 68)
        Me.txtException.MaxLength = 128
        Me.txtException.Name = "txtException"
        Me.txtException.Size = New System.Drawing.Size(188, 20)
        Me.txtException.TabIndex = 10
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 71)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 14)
        Me.Label7.TabIndex = 121
        Me.Label7.Text = "Exception"
        '
        'chkSkipChangeCheck
        '
        Me.chkSkipChangeCheck.AutoSize = True
        Me.chkSkipChangeCheck.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkSkipChangeCheck.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSkipChangeCheck.Location = New System.Drawing.Point(105, 15)
        Me.chkSkipChangeCheck.Name = "chkSkipChangeCheck"
        Me.chkSkipChangeCheck.Size = New System.Drawing.Size(133, 18)
        Me.chkSkipChangeCheck.TabIndex = 12
        Me.chkSkipChangeCheck.Text = "Skip Change Check"
        '
        'chkIMSPathData
        '
        Me.chkIMSPathData.AutoSize = True
        Me.chkIMSPathData.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkIMSPathData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIMSPathData.Location = New System.Drawing.Point(6, 15)
        Me.chkIMSPathData.Name = "chkIMSPathData"
        Me.chkIMSPathData.Size = New System.Drawing.Size(93, 18)
        Me.chkIMSPathData.TabIndex = 13
        Me.chkIMSPathData.Text = "IMSPathData"
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 211)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(72, 35)
        Me.Label8.TabIndex = 116
        Me.Label8.Text = "Datastore Description"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDatastoreDesc
        '
        Me.txtDatastoreDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatastoreDesc.Location = New System.Drawing.Point(77, 210)
        Me.txtDatastoreDesc.MaxLength = 255
        Me.txtDatastoreDesc.Multiline = True
        Me.txtDatastoreDesc.Name = "txtDatastoreDesc"
        Me.txtDatastoreDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDatastoreDesc.Size = New System.Drawing.Size(165, 35)
        Me.txtDatastoreDesc.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtDatastoreDesc, "User defined description " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "for this Datastore")
        '
        'cmbCharacterCode
        '
        Me.cmbCharacterCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCharacterCode.Location = New System.Drawing.Point(105, 13)
        Me.cmbCharacterCode.Name = "cmbCharacterCode"
        Me.cmbCharacterCode.Size = New System.Drawing.Size(213, 21)
        Me.cmbCharacterCode.TabIndex = 9
        Me.ToolTip1.SetToolTip(Me.cmbCharacterCode, "How Data is Encoded")
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 14)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(93, 19)
        Me.Label6.TabIndex = 113
        Me.Label6.Text = "Character Code"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbAccessMethod
        '
        Me.cmbAccessMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAccessMethod.Location = New System.Drawing.Point(311, 42)
        Me.cmbAccessMethod.Name = "cmbAccessMethod"
        Me.cmbAccessMethod.Size = New System.Drawing.Size(118, 21)
        Me.cmbAccessMethod.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.cmbAccessMethod, "How You Communicate" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "With the Database")
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(262, 45)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 14)
        Me.Label5.TabIndex = 111
        Me.Label5.Text = "Access"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPhysicalSource
        '
        Me.txtPhysicalSource.Location = New System.Drawing.Point(68, 16)
        Me.txtPhysicalSource.MaxLength = 128
        Me.txtPhysicalSource.Name = "txtPhysicalSource"
        Me.txtPhysicalSource.Size = New System.Drawing.Size(188, 20)
        Me.txtPhysicalSource.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtPhysicalSource, "How Database is named" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "On your System")
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(6, 13)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(54, 28)
        Me.Label4.TabIndex = 108
        Me.Label4.Text = "Physical Name"
        '
        'tvDatastoreStructures
        '
        Me.tvDatastoreStructures.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvDatastoreStructures.CheckBoxes = True
        Me.tvDatastoreStructures.HideSelection = False
        Me.tvDatastoreStructures.Location = New System.Drawing.Point(6, 13)
        Me.tvDatastoreStructures.Name = "tvDatastoreStructures"
        Me.tvDatastoreStructures.Size = New System.Drawing.Size(236, 190)
        Me.tvDatastoreStructures.TabIndex = 106
        Me.ToolTip1.SetToolTip(Me.tvDatastoreStructures, "Available Structures")
        '
        'txtDatastoreName
        '
        Me.txtDatastoreName.Location = New System.Drawing.Point(68, 42)
        Me.txtDatastoreName.MaxLength = 20
        Me.txtDatastoreName.Name = "txtDatastoreName"
        Me.txtDatastoreName.Size = New System.Drawing.Size(188, 20)
        Me.txtDatastoreName.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtDatastoreName, "Name Your Datastore")
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 45)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 14)
        Me.Label3.TabIndex = 104
        Me.Label3.Text = "Alias"
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.Blue
        Me.Label12.Location = New System.Drawing.Point(6, 36)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(465, 14)
        Me.Label12.TabIndex = 44
        Me.Label12.Text = "Note : If you enter hex value then it must start with 0x (e.g. hex value for semi" & _
            "colon (;) is 0x3B)"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(444, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 16)
        Me.Label11.TabIndex = 43
        Me.Label11.Text = "Text Qualifier"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbTextQualifier
        '
        Me.cmbTextQualifier.Location = New System.Drawing.Point(530, 14)
        Me.cmbTextQualifier.Name = "cmbTextQualifier"
        Me.cmbTextQualifier.Size = New System.Drawing.Size(112, 21)
        Me.cmbTextQualifier.TabIndex = 19
        Me.ToolTip1.SetToolTip(Me.cmbTextQualifier, "Character(s) that" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "qualifies your Data")
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(105, 16)
        Me.Label10.TabIndex = 41
        Me.Label10.Text = "Column Delimiter"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cmbRowDelimiter
        '
        Me.cmbRowDelimiter.Location = New System.Drawing.Point(322, 14)
        Me.cmbRowDelimiter.Name = "cmbRowDelimiter"
        Me.cmbRowDelimiter.Size = New System.Drawing.Size(104, 21)
        Me.cmbRowDelimiter.TabIndex = 18
        Me.ToolTip1.SetToolTip(Me.cmbRowDelimiter, "Character that separates" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "your rows")
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(222, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(94, 16)
        Me.Label9.TabIndex = 39
        Me.Label9.Text = "Row Delimiter"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbColDelimiter
        '
        Me.cmbColDelimiter.Location = New System.Drawing.Point(117, 14)
        Me.cmbColDelimiter.Name = "cmbColDelimiter"
        Me.cmbColDelimiter.Size = New System.Drawing.Size(100, 21)
        Me.cmbColDelimiter.TabIndex = 17
        Me.ToolTip1.SetToolTip(Me.cmbColDelimiter, "Character that separates" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "your columns")
        '
        'gbProp
        '
        Me.gbProp.Controls.Add(Me.Label4)
        Me.gbProp.Controls.Add(Me.txtPhysicalSource)
        Me.gbProp.Controls.Add(Me.Label3)
        Me.gbProp.Controls.Add(Me.cmbDatastoreType)
        Me.gbProp.Controls.Add(Me.txtDatastoreName)
        Me.gbProp.Controls.Add(Me.Label15)
        Me.gbProp.Controls.Add(Me.Label7)
        Me.gbProp.Controls.Add(Me.txtPortOrMQMgr)
        Me.gbProp.Controls.Add(Me.txtException)
        Me.gbProp.Controls.Add(Me.lblPortorMQMgr)
        Me.gbProp.Controls.Add(Me.Label5)
        Me.gbProp.Controls.Add(Me.cmbAccessMethod)
        Me.gbProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbProp.Location = New System.Drawing.Point(3, 82)
        Me.gbProp.Name = "gbProp"
        Me.gbProp.Size = New System.Drawing.Size(435, 96)
        Me.gbProp.TabIndex = 118
        Me.gbProp.TabStop = False
        Me.gbProp.Text = "Datastore Properties"
        '
        'gbExtProps
        '
        Me.gbExtProps.Controls.Add(Me.txtRestart)
        Me.gbExtProps.Controls.Add(Me.lblRestart)
        Me.gbExtProps.Controls.Add(Me.chkIMSPathData)
        Me.gbExtProps.Controls.Add(Me.txtPoll)
        Me.gbExtProps.Controls.Add(Me.lblPoll)
        Me.gbExtProps.Controls.Add(Me.chkSkipChangeCheck)
        Me.gbExtProps.Controls.Add(Me.txtUOW)
        Me.gbExtProps.Controls.Add(Me.lblUOW)
        Me.gbExtProps.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbExtProps.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbExtProps.Location = New System.Drawing.Point(3, 184)
        Me.gbExtProps.Name = "gbExtProps"
        Me.gbExtProps.Size = New System.Drawing.Size(435, 68)
        Me.gbExtProps.TabIndex = 119
        Me.gbExtProps.TabStop = False
        Me.gbExtProps.Text = "Extended Properties"
        '
        'txtRestart
        '
        Me.txtRestart.Location = New System.Drawing.Point(60, 39)
        Me.txtRestart.Name = "txtRestart"
        Me.txtRestart.Size = New System.Drawing.Size(253, 20)
        Me.txtRestart.TabIndex = 141
        '
        'lblRestart
        '
        Me.lblRestart.AutoSize = True
        Me.lblRestart.Location = New System.Drawing.Point(6, 42)
        Me.lblRestart.Name = "lblRestart"
        Me.lblRestart.Size = New System.Drawing.Size(48, 13)
        Me.lblRestart.TabIndex = 140
        Me.lblRestart.Text = "Restart"
        '
        'gbTarget
        '
        Me.gbTarget.Controls.Add(Me.Label16)
        Me.gbTarget.Controls.Add(Me.cmbOperationType)
        Me.gbTarget.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbTarget.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbTarget.Location = New System.Drawing.Point(365, 339)
        Me.gbTarget.Name = "gbTarget"
        Me.gbTarget.Size = New System.Drawing.Size(327, 44)
        Me.gbTarget.TabIndex = 141
        Me.gbTarget.TabStop = False
        Me.gbTarget.Text = "Target Properties"
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.cbIfSpaceNum)
        Me.gbSource.Controls.Add(Me.Label20)
        Me.gbSource.Controls.Add(Me.cbIfSpaceChar)
        Me.gbSource.Controls.Add(Me.lblExtType)
        Me.gbSource.Controls.Add(Me.cbInvalidNum)
        Me.gbSource.Controls.Add(Me.Label18)
        Me.gbSource.Controls.Add(Me.cbInvalidChar)
        Me.gbSource.Controls.Add(Me.lblInValid)
        Me.gbSource.Controls.Add(Me.cbIfNullNum)
        Me.gbSource.Controls.Add(Me.Label19)
        Me.gbSource.Controls.Add(Me.cbIfNullChar)
        Me.gbSource.Controls.Add(Me.Label17)
        Me.gbSource.Controls.Add(Me.cmbCharacterCode)
        Me.gbSource.Controls.Add(Me.Label6)
        Me.gbSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSource.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbSource.Location = New System.Drawing.Point(3, 258)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(435, 125)
        Me.gbSource.TabIndex = 142
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Source Datastore Field Properties"
        '
        'cbIfSpaceNum
        '
        Me.cbIfSpaceNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfSpaceNum.FormattingEnabled = True
        Me.cbIfSpaceNum.Location = New System.Drawing.Point(322, 94)
        Me.cbIfSpaceNum.Name = "cbIfSpaceNum"
        Me.cbIfSpaceNum.Size = New System.Drawing.Size(107, 21)
        Me.cbIfSpaceNum.TabIndex = 125
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Location = New System.Drawing.Point(249, 97)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(69, 13)
        Me.Label20.TabIndex = 124
        Me.Label20.Text = "for All Num"
        '
        'cbIfSpaceChar
        '
        Me.cbIfSpaceChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfSpaceChar.FormattingEnabled = True
        Me.cbIfSpaceChar.Location = New System.Drawing.Point(131, 94)
        Me.cbIfSpaceChar.MaxDropDownItems = 20
        Me.cbIfSpaceChar.Name = "cbIfSpaceChar"
        Me.cbIfSpaceChar.Size = New System.Drawing.Size(107, 21)
        Me.cbIfSpaceChar.TabIndex = 123
        '
        'lblExtType
        '
        Me.lblExtType.AutoSize = True
        Me.lblExtType.Location = New System.Drawing.Point(6, 97)
        Me.lblExtType.Name = "lblExtType"
        Me.lblExtType.Size = New System.Drawing.Size(122, 13)
        Me.lblExtType.TabIndex = 122
        Me.lblExtType.Text = "If Space for All Char"
        '
        'cbInvalidNum
        '
        Me.cbInvalidNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbInvalidNum.FormattingEnabled = True
        Me.cbInvalidNum.Location = New System.Drawing.Point(322, 67)
        Me.cbInvalidNum.Name = "cbInvalidNum"
        Me.cbInvalidNum.Size = New System.Drawing.Size(106, 21)
        Me.cbInvalidNum.TabIndex = 121
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Location = New System.Drawing.Point(249, 70)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(69, 13)
        Me.Label18.TabIndex = 120
        Me.Label18.Text = "for All Num"
        '
        'cbInvalidChar
        '
        Me.cbInvalidChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbInvalidChar.FormattingEnabled = True
        Me.cbInvalidChar.Location = New System.Drawing.Point(131, 67)
        Me.cbInvalidChar.MaxDropDownItems = 20
        Me.cbInvalidChar.Name = "cbInvalidChar"
        Me.cbInvalidChar.Size = New System.Drawing.Size(107, 21)
        Me.cbInvalidChar.TabIndex = 119
        '
        'lblInValid
        '
        Me.lblInValid.AutoSize = True
        Me.lblInValid.Location = New System.Drawing.Point(6, 70)
        Me.lblInValid.Name = "lblInValid"
        Me.lblInValid.Size = New System.Drawing.Size(112, 13)
        Me.lblInValid.TabIndex = 118
        Me.lblInValid.Text = "Invalid for All Char"
        '
        'cbIfNullNum
        '
        Me.cbIfNullNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfNullNum.FormattingEnabled = True
        Me.cbIfNullNum.Location = New System.Drawing.Point(322, 40)
        Me.cbIfNullNum.Name = "cbIfNullNum"
        Me.cbIfNullNum.Size = New System.Drawing.Size(107, 21)
        Me.cbIfNullNum.TabIndex = 117
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(249, 45)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(69, 13)
        Me.Label19.TabIndex = 116
        Me.Label19.Text = "for All Num"
        '
        'cbIfNullChar
        '
        Me.cbIfNullChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfNullChar.FormattingEnabled = True
        Me.cbIfNullChar.Location = New System.Drawing.Point(131, 40)
        Me.cbIfNullChar.Name = "cbIfNullChar"
        Me.cbIfNullChar.Size = New System.Drawing.Size(107, 21)
        Me.cbIfNullChar.TabIndex = 115
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 43)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(108, 13)
        Me.Label17.TabIndex = 114
        Me.Label17.Text = "If Null for All Char"
        '
        'gbStruct
        '
        Me.gbStruct.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbStruct.Controls.Add(Me.txtDatastoreDesc)
        Me.gbStruct.Controls.Add(Me.Label8)
        Me.gbStruct.Controls.Add(Me.tvDatastoreStructures)
        Me.gbStruct.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStruct.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbStruct.Location = New System.Drawing.Point(444, 82)
        Me.gbStruct.Name = "gbStruct"
        Me.gbStruct.Size = New System.Drawing.Size(248, 251)
        Me.gbStruct.TabIndex = 143
        Me.gbStruct.TabStop = False
        Me.gbStruct.Text = "Descriptions"
        '
        'gbDel
        '
        Me.gbDel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDel.Controls.Add(Me.cmbTextQualifier)
        Me.gbDel.Controls.Add(Me.Label11)
        Me.gbDel.Controls.Add(Me.Label12)
        Me.gbDel.Controls.Add(Me.Label10)
        Me.gbDel.Controls.Add(Me.cmbRowDelimiter)
        Me.gbDel.Controls.Add(Me.cmbColDelimiter)
        Me.gbDel.Controls.Add(Me.Label9)
        Me.gbDel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDel.Location = New System.Drawing.Point(3, 389)
        Me.gbDel.Name = "gbDel"
        Me.gbDel.Size = New System.Drawing.Size(689, 57)
        Me.gbDel.TabIndex = 144
        Me.gbDel.TabStop = False
        Me.gbDel.Text = "File Properties"
        '
        'gbField
        '
        Me.gbField.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbField.Controls.Add(Me.SplitContainer1)
        Me.gbField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbField.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbField.Location = New System.Drawing.Point(3, 452)
        Me.gbField.Name = "gbField"
        Me.gbField.Size = New System.Drawing.Size(689, 151)
        Me.gbField.TabIndex = 145
        Me.gbField.TabStop = False
        Me.gbField.Text = "Field Properties"
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.gbFldname)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.gbAtt)
        Me.SplitContainer1.Size = New System.Drawing.Size(683, 132)
        Me.SplitContainer1.SplitterDistance = 214
        Me.SplitContainer1.TabIndex = 0
        '
        'gbFldname
        '
        Me.gbFldname.Controls.Add(Me.tvFields)
        Me.gbFldname.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFldname.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbFldname.Location = New System.Drawing.Point(0, 0)
        Me.gbFldname.Name = "gbFldname"
        Me.gbFldname.Size = New System.Drawing.Size(214, 132)
        Me.gbFldname.TabIndex = 0
        Me.gbFldname.TabStop = False
        Me.gbFldname.Text = "Field Names"
        '
        'gbAtt
        '
        Me.gbAtt.Controls.Add(Me.cmbListViewCombo)
        Me.gbAtt.Controls.Add(Me.txtListViewText)
        Me.gbAtt.Controls.Add(Me.lvFieldAttrs)
        Me.gbAtt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAtt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbAtt.Location = New System.Drawing.Point(0, 0)
        Me.gbAtt.Name = "gbAtt"
        Me.gbAtt.Size = New System.Drawing.Size(465, 132)
        Me.gbAtt.TabIndex = 0
        Me.gbAtt.TabStop = False
        Me.gbAtt.Text = "Field Attributes"
        '
        'frmDatastore
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Nothing
        Me.ClientSize = New System.Drawing.Size(704, 663)
        Me.Controls.Add(Me.gbProp)
        Me.Controls.Add(Me.gbExtProps)
        Me.Controls.Add(Me.gbSource)
        Me.Controls.Add(Me.gbTarget)
        Me.Controls.Add(Me.gbField)
        Me.Controls.Add(Me.gbStruct)
        Me.Controls.Add(Me.gbDel)
        Me.Controls.Add(Me.rbLookup)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(712, 528)
        Me.Name = "frmDatastore"
        Me.Text = "Datastore Properties"
        Me.Controls.SetChildIndex(Me.rbLookup, 0)
        Me.Controls.SetChildIndex(Me.gbDel, 0)
        Me.Controls.SetChildIndex(Me.gbStruct, 0)
        Me.Controls.SetChildIndex(Me.gbField, 0)
        Me.Controls.SetChildIndex(Me.gbTarget, 0)
        Me.Controls.SetChildIndex(Me.gbSource, 0)
        Me.Controls.SetChildIndex(Me.gbExtProps, 0)
        Me.Controls.SetChildIndex(Me.gbProp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbProp.ResumeLayout(False)
        Me.gbProp.PerformLayout()
        Me.gbExtProps.ResumeLayout(False)
        Me.gbExtProps.PerformLayout()
        Me.gbTarget.ResumeLayout(False)
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.gbStruct.ResumeLayout(False)
        Me.gbStruct.PerformLayout()
        Me.gbDel.ResumeLayout(False)
        Me.gbDel.PerformLayout()
        Me.gbField.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.gbFldname.ResumeLayout(False)
        Me.gbAtt.ResumeLayout(False)
        Me.gbAtt.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form events"

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbListViewCombo.SelectedIndexChanged, txtListViewText.TextChanged, chkIMSPathData.CheckedChanged, chkSkipChangeCheck.CheckedChanged, cmbAccessMethod.SelectedIndexChanged, cmbCharacterCode.SelectedIndexChanged, cmbDatastoreType.SelectedIndexChanged, cmbOperationType.SelectedIndexChanged, txtDatastoreDesc.TextChanged, txtException.TextChanged, txtPhysicalSource.TextChanged, txtPortOrMQMgr.TextChanged, txtUOW.TextChanged, txtPoll.TextChanged '// modified by TK 1/2011
        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)
    End Sub

    Private Sub StartLoad()
        IsEventFromCode = True
        objThis.IsModified = False
        txtDatastoreName.Select()

        '#If CONFIG = "ETI" Then
        '                Me.Text = "ETI CDC Studio V3 " & Application.ProductVersion
        '                Me.PictureBox1.ImageLocation = "C:\Documents and Settings\Tom\My Documents\Visual Studio 2005\Projects\sqdstudio\images\ETI_Logo_Installation_Graphic\ETI-installer-right.bmp"
        '#Else
        '        Me.Text = "SQData Studio V3 " & Application.ProductVersion
        '        Me.PictureBox1.ImageLocation = "C:\Documents and Settings\Tom\My Documents\Visual Studio 2005\Projects\sqdstudio\images\FormTop\sq_skyblue.jpg"
        '        Me.Icon = Drawing.Icon.ExtractAssociatedIcon("C:\Documents and Settings\Tom\My Documents\Visual Studio 2005\Projects\sqdstudio\images\icons\SQ.ico")
        '#End If

    End Sub

    Private Sub SetDefaultName()

        Try
            Dim EngObj As clsEngine = Nothing
            Dim EnvObj As clsEnvironment = Nothing
            Dim NewTgtName As String = ""
            Dim NewSrcName As String = ""
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            EnvObj = objThis.Environment
            If objThis.Engine IsNot Nothing Then
                EngObj = objThis.Engine
            End If

            While GoodName = False
                GoodName = True

                NewName = "U_DS" & Count.ToString
                If Count = 1 Then
                    If objThis.IsLookUp = True Then
                        NewSrcName = "LOOKUP" & Count.ToString
                    Else
                        NewSrcName = "CDCIN"
                    End If
                Else
                    If objThis.IsLookUp = True Then
                        NewSrcName = "LOOKUP" & Count.ToString
                    Else
                        NewSrcName = "CDCIN" & Count.ToString
                    End If
                End If
                NewTgtName = objThis.DsDirection & "_" & "DS" & Count.ToString

                If objThis.Engine IsNot Nothing Then
                    If objThis.DsDirection = "S" Then
                        For Each TestDS As clsDatastore In EngObj.Sources
                            If TestDS.DatastoreName = NewSrcName Then
                                GoodName = False
                                Exit For
                            End If
                        Next
                        'Else
                        '    For Each TestDS As clsDatastore In EngObj.Targets
                        '        If TestDS.DatastoreName = NewTgtName Then
                        '            GoodName = False
                        '            Exit For
                        '        End If
                        '    Next
                    End If
                End If

                'For Each TestDS As clsDatastore In EnvObj.Datastores
                '    If TestDS.DatastoreName = NewName Or TestDS.DatastoreName = NewSrcName Or _
                '    TestDS.DatastoreName = NewTgtName Then
                '        GoodName = False
                '        Exit For
                '    End If
                'Next

                Count += 1
            End While

            If objThis.DsDirection = "S" Then
                txtDatastoreName.Text = NewSrcName
            ElseIf objThis.DsDirection = "T" Then
                txtDatastoreName.Text = NewTgtName
            Else
                txtDatastoreName.Text = NewName
            End If


        Catch ex As Exception
            LogError(ex, "frmDS SetDefaultName")
        End Try

    End Sub

    Private Sub EndLoad()
        IsEventFromCode = False
        'cmdShowHideFieldAttr.Text = "Show Fields/Attributes"
    End Sub

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        sl = New MyListener(lvFieldAttrs)
    End Sub

    Private Sub sl_MyScroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles sl.MyScroll
        cmbListViewCombo.Visible = False
        txtListViewText.Visible = False
    End Sub

    Private Sub frmDatastore_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        'pnlFields.SuspendLayout()
        'pnlFields.Location = pnlSelect.Location
        'pnlFields.Size = pnlSelect.Size
        'pnlFields.ResumeLayout()
    End Sub

#End Region

#Region "click events"

    Public Overrides Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        If objThis.IsModified = False Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
            Exit Sub
        End If

        If MsgBox("Do you really want to discard the changes?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
            Me.Close()
        Else
            Me.DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidateSelection() = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            Else
                DialogResult = Windows.Forms.DialogResult.OK
            End If

            If objThis.ValidateNewObject(txtDatastoreName.Text) = False Or ValidateNewName128(txtDatastoreName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            If txtPhysicalSource.Text.Trim = "" Then
                MsgBox("'Physical Name' Must not be left blank", MsgBoxStyle.OkOnly, "Missing Physical Name")
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If


            '//save structure selections too 
            Call SaveCurrentSelection()

            If objThis.DsDirection = DS_DIRECTION_TARGET Or objThis.IsLookUp = True Then
                If objThis.DatastoreType = enumDatastore.DS_RELATIONAL Then
                    KeyForRel = HasKeyForRel()
                End If
            End If
            'If KeyForRel = False Then
            '    Exit Try
            'End If

            If IsNewObj = True Then
                '//Before we add new Datastore set Datastore fields
                'LoadFldArrFromXmlFile(GetOutDumpXML, objThis.ObjFields)

                If objThis.AddNew = True Then
                    'If objThis.Engine IsNot Nothing Then
                    '    Dim objenv As clsDatastore = CType(objThis.Clone(objThis.Environment), clsDatastore)
                    '    objenv.Engine = Nothing
                    '    objenv.DsDirection = ""
                    '    If objenv.AddNew = False Then
                    '        DialogResult = Windows.Forms.DialogResult.Retry
                    '        Exit Sub
                    '    End If
                    'End If
                    'MsgBox("saved successfully", MsgBoxStyle.Information)
                Else
                    'MsgBox("Error during save operation", MsgBoxStyle.Critical)
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex, "frmDatastore cmdOK_Click")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdShowHideFieldAttr_Click()

        'If IsNewObj = True Then
        '    If ValidateSelection() = False Then Exit Sub
        'End If

        'pnlFields.Visible = Not pnlFields.Visible
        'pnlSelect.Visible = Not pnlSelect.Visible

        Call UpdateControls()

        '//Dont reload everytime once its loaded from database
        'If tvFields.GetNodeCount(True) > 0 And IsNewObj = False Then Exit Sub

        tvFields.BackColor = Color.LightBlue
        lvFieldAttrs.BackColor = Color.LightBlue

        'lblFieldName.ForeColor = Color.Red
        'lblFieldName.Text = "********* Loading *********"
        Me.Refresh()

        'If pnlFields.Visible = True Then
        Call FillFieldAttr()
        'End If
        tvFields.BackColor = Color.White
        lvFieldAttrs.BackColor = Color.White
        'lblFieldName.Text = ""
        'lblFieldName.ForeColor = Color.Blue
        If tvFields.GetNodeCount(True) > 0 Then
            tvFields.SelectedNode = tvFields.Nodes.Item(0)
        End If
    End Sub

    'Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If SelectFirstMatchingNode(tvFields, txtSearchField.Text) Is Nothing Then
    '        MsgBox("No matching node found for entered text", MsgBoxStyle.Critical, MsgTitle)
    '    End If

    'End Sub

    Private Sub tvFields_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvFields.Click

        txtListViewText.Visible = False
        cmbListViewCombo.Visible = False

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Dim HelpObj As String

        HelpObj = Me.objThis.DsDirection

        If HelpObj = "T" Then
            ShowHelp(modHelp.HHId.H_Add_Targets)
        Else
            ShowHelp(modHelp.HHId.H_Add_Sources)
        End If

        ShowHelp(modHelp.HHId.H_Datastore_Name)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) ' Handles MyBase.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click(sender, e)
            Case Keys.Enter
                cmdOk_Click(sender, e)
            Case Keys.F1
                cmdHelp_Click(sender, e)
        End Select
    End Sub

#End Region

#Region "Form controls Events"

    Private Sub cmbListViewCombo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbListViewCombo.SelectedIndexChanged

        If IsEventFromCode = True Then Exit Sub
        Try

            If Not (tvFields.SelectedNode Is Nothing) Then
                Dim lstItm As Mylist
                lstItm = cmbListViewCombo.SelectedItem
                CType(tvFields.SelectedNode.Tag, clsField).SetSingleFieldAttr(lvItem.Index, lstItm.ItemData)
                lvItem.SubItems(1).Text = cmbListViewCombo.Text
                CType(tvFields.SelectedNode.Tag, clsField).IsModified = True
            End If
        Catch ex As Exception
            LogError(ex)
        End Try

    End Sub

    Private Sub txtListViewText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtListViewText.TextChanged

        If IsEventFromCode = True Then Exit Sub

        If Not (tvFields.SelectedNode Is Nothing) Then
            CType(tvFields.SelectedNode.Tag, clsField).SetSingleFieldAttr(lvItem.Index, txtListViewText.Text)
            lvItem.SubItems(1).Text = txtListViewText.Text
            CType(tvFields.SelectedNode.Tag, clsField).IsModified = True
        End If

    End Sub

    Private Sub tvDatastoreStructures_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvDatastoreStructures.KeyDown

        If tvDatastoreStructures.SelectedNode Is Nothing Then Exit Sub

        If e.KeyCode = Keys.Enter Then
            cmdShowHideFieldAttr_Click()
        End If
        If e.KeyCode = Keys.F1 Then
            Dim HelpObj As String

            HelpObj = Me.objThis.DsDirection

            If HelpObj = "T" Then
                ShowHelp(modHelp.HHId.H_Add_Targets)
            Else
                ShowHelp(modHelp.HHId.H_Add_Sources)
            End If
            ShowHelp(modHelp.HHId.H_Datastore_Name)
        End If

    End Sub

    '//added by npatel on 9/6/05
    Private Sub cmbDatastoreType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDatastoreType.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.DatastoreType = CType(cmbDatastoreType.Items(cmbDatastoreType.SelectedIndex), Mylist).ItemData

        ShowHideControls(objThis.DatastoreType)

    End Sub

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        e.Node.SelectedImageIndex = e.Node.ImageIndex
        ShowFieldAttributes(e.Node)

    End Sub

    'Private Sub txtSearchField_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    SelectFirstMatchingNode(tvFields, txtSearchField.Text)

    'End Sub

    Private Sub tvDatastoreStructures_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvDatastoreStructures.AfterCheck

        '//Couple of things to do when user check/uncheck any node
        Try
            If IsEventFromCode = True Then Exit Try
            Dim NodeText As String = ""

            If e.Node.Checked = True Then
                If (e.Node.Parent Is Nothing) Then
                    If objThis.DsDirection = DS_DIRECTION_TARGET Then 'objThis.DatastoreType = enumDatastore.DS_RELATIONAL And 
                        e.Node.Checked = False
                        MsgBox("Only one selection per target allowed")
                        Exit Try
                    Else
                        '//if we come here means parent node is checked. 
                        '// Check all children, then uncheck parent (folder) node
                        For Each nd As TreeNode In e.Node.Nodes
                            nd.Checked = True
                            NodeText = nd.Text
                        Next
                        If objThis.DsDirection = DS_DIRECTION_TARGET Then
                            txtDatastoreName.Text = "T_" & NodeText
                        End If
                        e.Node.Checked = False
                    End If
                Else
                    If objThis.DsDirection = DS_DIRECTION_TARGET Then 'objThis.DatastoreType = enumDatastore.DS_RELATIONAL And 
                        IsEventFromCode = True
                        For Each nd As TreeNode In tvDatastoreStructures.Nodes
                            nd.Checked = False
                            For Each chnd As TreeNode In nd.Nodes
                                chnd.Checked = False
                            Next
                        Next
                        e.Node.Checked = True
                        txtDatastoreName.Text = "T_" & e.Node.Text
                        IsEventFromCode = False
                    End If
                End If
            End If

            tvDatastoreStructures.SelectedNode = e.Node
            'tvDatastoreStructures_AfterSelect(sender, e)

        Catch ex As Exception
            LogError(ex, "frmDatastore tvDatastoreStructures_AfterCheck")
        End Try

    End Sub

    Private Sub cmbAccessMethod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAccessMethod.SelectedIndexChanged

        SetPortOrMQ()
        SetPoll()

    End Sub

    Private Sub tvDatastoreStructures_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvDatastoreStructures.AfterSelect

        If CType(tvDatastoreStructures.SelectedNode.Tag, INode).Type = NODE_FO_STRUCT Then Exit Sub
        'If CType(tvDatastoreStructures.SelectedNode.Tag, INode).Type = NODE_FO_STRUCT Then
        'cmdShowHideFieldAttr.Enabled = False
        Call cmdShowHideFieldAttr_Click()
        'Else
        'cmdShowHideFieldAttr.Enabled = True
        'End If

    End Sub

    Private Sub lvFieldAttrs_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvFieldAttrs.MouseUp

        'Get the item on the row that is clicked.
        lvItem = Me.lvFieldAttrs.GetItemAt(e.X, e.Y)

        ' Make sure that an item is clicked.
        If Not (lvItem Is Nothing) Then
            'If lvItem.Index > modDeclares.enumFieldAttributes.ATTR_IDENTVAL Then
            cmbListViewCombo.Visible = False
            txtListViewText.Visible = False
            Exit Sub
        End If

        '' Get the bounds of the item that is clicked.
        'Dim ClickedItem As Rectangle = lvItem.Bounds

        'ClickedItem.X = lvFieldAttrs.Columns(0).Width
        'ClickedItem.Width = lvFieldAttrs.Columns(1).Width

        '' Verify that the column is completely scrolled off to the left.
        'If ((ClickedItem.Left + Me.lvFieldAttrs.Columns(1).Width) < 0) Then
        '    ' If the cell is out of view to the left, do nothing.
        '    Return
        '    ' Verify that the column is partially scrolled off to the left.
        'ElseIf (ClickedItem.Left < 0) Then
        '    ' Determine if column extends beyond right side of ListView.
        '    If ((ClickedItem.Left + Me.lvFieldAttrs.Columns(0).Width) > Me.lvFieldAttrs.Width) Then
        '        ' Set width of column to match width of ListView.
        '        ClickedItem.Width = Me.lvFieldAttrs.Width
        '        ClickedItem.X = 0
        '    End If

        'ElseIf (Me.lvFieldAttrs.Columns(1).Width > Me.lvFieldAttrs.Width) Then
        '    ClickedItem.Width = Me.lvFieldAttrs.Width
        'Else
        '    ClickedItem.Width = Me.lvFieldAttrs.Columns(1).Width
        '    ClickedItem.X = 2 + Me.lvFieldAttrs.Columns(0).Width
        'End If

        '' Adjust the top to account for the location of the ListView.
        'ClickedItem.Y += Me.lvFieldAttrs.Top
        'ClickedItem.X += Me.lvFieldAttrs.Left

        'If lvItem.Index = modDeclares.enumFieldAttributes.ATTR_ISKEY Or _
        '    lvItem.Index = modDeclares.enumFieldAttributes.ATTR_RETYPE Or _
        '    lvItem.Index = modDeclares.enumFieldAttributes.ATTR_EXTTYPE Or _
        '    lvItem.Index = modDeclares.enumFieldAttributes.ATTR_INVALID Or _
        '    lvItem.Index = modDeclares.enumFieldAttributes.ATTR_FKEY Then

        '    ' Assign calculated bounds to the ComboBox.
        '    Me.cmbListViewCombo.Bounds = ClickedItem

        '    ShowLVCombo(lvItem)
        'Else
        '    ' Assign calculated bounds to the TextBox.
        '    Me.txtListViewText.Bounds = ClickedItem

        '    ShowLVText(lvItem)
        'End If
        'End If

    End Sub

#End Region

    '//cNode is current environment structure root node
    Friend Function NewObj(ByVal cNode As TreeNode, Optional ByVal DatastoreType As enumDatastore = modDeclares.enumDatastore.DS_UNKNOWN, Optional ByVal IsLook As Boolean = False) As clsDatastore

        IsNewObj = True

        StartLoad()

        objThis.IsLookUp = IsLook
        'If objThis.IsLookUp = True Then
        '    gbProp.Text = "Lookup Datastore Properties"
        '    Me.Text = "Lookup Datastore Properties"
        'Else
        '    gbProp.Text = "Datastore Properties"
        '    Me.Text = "Datastore Properties"
        'End If

        SetDefaultName()

        '// added by KS and TK 11/6/2006
        If (txtPoll.Text.Trim = "") Then
            txtPoll.Text = "5"
        End If

        objThis.DatastoreType = DatastoreType

        'SetDefaultName()

        InitControls()

        ShowHideControls(DatastoreType)

        FillStructure(cNode)

        txtDatastoreName.SelectAll()

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = objThis
            Case Windows.Forms.DialogResult.Retry
                tvDatastoreStructures.ExpandAll()
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Function GetCharDesc(ByVal strChar As String) As String
        Select Case strChar
            Case vbCrLf
                GetCharDesc = "{CR}{LF}"
            Case vbCr
                GetCharDesc = "{CR}"
            Case vbLf
                GetCharDesc = "{LF}"
            Case vbTab
                GetCharDesc = "Tab"
            Case "|"
                GetCharDesc = "Verticalbar"
            Case ","
                GetCharDesc = "Comma"
            Case "'"
                GetCharDesc = "Single Quote"
            Case """"
                GetCharDesc = "Double Quote"
            Case ";"
                GetCharDesc = "Semicolon"
            Case Else
                GetCharDesc = strChar
        End Select
    End Function

    '//write form controls values to object
    Function SaveCurrentSelection() As Boolean

        Try
            '//save form data to the object before saving object to database
            objThis.DatastoreName = txtDatastoreName.Text
            objThis.DatastoreDescription = txtDatastoreDesc.Text

            '//only store if target
            If objThis.DsDirection = DS_DIRECTION_TARGET Then
                objThis.OperationType = CType(cmbOperationType.SelectedItem, Mylist).ItemData
            End If

            If cmbCharacterCode.Enabled = True Then
                objThis.DsCharacterCode = CType(cmbCharacterCode.SelectedItem, Mylist).ItemData
            Else
                objThis.DsCharacterCode = ""
            End If



            '//npatel 10/7/05
            If cmbAccessMethod.Enabled = True Then
                Dim Itm As Mylist
                Itm = cmbAccessMethod.SelectedItem
                If (Itm Is Nothing) = False Then
                    If Itm.ItemData = DS_ACCESSMETHOD_IP Then
                        objThis.DsPort = txtPortOrMQMgr.Text
                    ElseIf Itm.ItemData = DS_ACCESSMETHOD_MQSERIES Then
                        objThis.DsQueMgr = txtPortOrMQMgr.Text
                    Else
                        objThis.DsPort = txtPortOrMQMgr.Text
                    End If
                End If

                objThis.DsAccessMethod = CType(cmbAccessMethod.SelectedItem, Mylist).ItemData
            Else
                objThis.DsAccessMethod = ""
                objThis.DsQueMgr = ""
                objThis.DsPort = ""
            End If

            objThis.DsPhysicalSource = txtPhysicalSource.Text
            objThis.ExceptionDatastore = txtException.Text
            objThis.DsUOW = txtUOW.Text
            objThis.Poll = txtPoll.Text
            objThis.Restart = txtRestart.Text


            'objThis.IsCmmtKey = IIf(chkCmmtKey.Checked, "1", "0")
            objThis.IsIMSPathData = IIf(chkIMSPathData.Checked, "1", "0")
            'objThis.IsOrdered = IIf(chkOrdered.Checked, "1", "0")
            objThis.IsSkipChangeCheck = IIf(chkSkipChangeCheck.Checked, "1", "0")
            'objThis.IsBeforeImage = IIf(chkBeforeImage.Checked, "1", "0") '//new by npatel on 8/10/05

            If objThis.DatastoreType = modDeclares.enumDatastore.DS_DELIMITED Then
                If (cmbColDelimiter.SelectedItem Is Nothing) = True Then
                    objThis.ColumnDelimiter = ""
                Else
                    objThis.ColumnDelimiter = CType(cmbColDelimiter.SelectedItem, Mylist).ItemData
                End If

                If (cmbRowDelimiter.SelectedItem Is Nothing) = True Then
                    objThis.RowDelimiter = ""
                Else
                    objThis.RowDelimiter = CType(cmbRowDelimiter.SelectedItem, Mylist).ItemData
                End If

                If (cmbTextQualifier.SelectedItem Is Nothing) = True Then
                    objThis.TextQualifier = ""
                Else
                    objThis.TextQualifier = CType(cmbTextQualifier.SelectedItem, Mylist).ItemData
                End If
            End If

            '//now save structure selection

            '//before we clear old selection save it to another array so later on while saving this object we can compare for any new addition in selection list
            objThis.OldObjSelections.Clear() '//new : 8/15/05
            objThis.OldObjSelections = objThis.ObjSelections.Clone

            '//clear previous selection and store new selection
            objThis.ObjSelections.Clear()

            '[STR TYPE1]
            '   + [Str-A]
            '       + [Str-A-Subset1]
            '       + [Str-A-Subset2]
            '[STR TYPE2]
            '   + [Str-B]
            '       + [Str-B-Subset1]
            '       + [Str-B-Subset2]

            '//now save structure selection

            '/// completely rewritten 1/8/2007 for ds selections by TK

            Dim fNode, pNode, cNode As TreeNode
            Dim tempObj As New clsDSSelection
            Dim tempsel As New clsDSSelection
            Dim i As Integer = 0
            Dim flag As Boolean = False

            '//for each folder node
            For Each fNode In tvDatastoreStructures.Nodes
                '//each structure under folder
                For Each pNode In fNode.Nodes
                    '//check for structures if selected then we dont consider child selection otherwise scan for children selection
                    If pNode.Checked = True Then
                        tempsel = objThis.CloneSSeltoDSSel(pNode.Tag)
                        For i = 0 To objThis.OldObjSelections.Count - 1
                            tempObj = objThis.OldObjSelections(i)
                            tempObj.LoadMe()
                            '// if the DSSelection exists already in old DSselections, 
                            '//then add it to current DS selections
                            If tempObj.SelectionName = tempsel.SelectionName Then
                                objThis.ObjSelections.Add(tempObj)
                                flag = True
                                Exit For
                            End If
                        Next
                        If flag = False Then
                            objThis.ObjSelections.Add(tempsel)
                        End If
                        flag = False
                    Else
                        '//each selection Of Selection
                        For Each cNode In pNode.Nodes
                            If cNode.Checked = True Then
                                tempsel = objThis.CloneSSeltoDSSel(cNode.Tag)
                                For i = 0 To objThis.OldObjSelections.Count - 1
                                    tempObj = objThis.OldObjSelections(i)
                                    tempObj.LoadMe()
                                    If tempObj.SelectionName = tempsel.SelectionName Then
                                        objThis.ObjSelections.Add(tempObj)
                                        flag = True
                                        Exit For
                                    End If
                                Next
                                If flag = False Then
                                    objThis.ObjSelections.Add(tempsel)
                                End If
                                flag = False
                            End If
                        Next
                    End If
                Next
            Next
            '/// once all datastore selections are saved from form, make sure all parents are set
            objThis.SetDSselParents()

            Return True

        Catch ex As Exception
            LogError(ex, "frmDatastore SaveCurrentSelection")
            Return False
        End Try

    End Function

    Friend Function ShowHideControls(ByVal DatastoreType As enumDatastore) As Boolean

        '//Show/Hide Src/Tgt and Delimited DS option  (TK 12/18/2009)
        Pnt.X = 3
        Pnt.Y = 82

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then  'Or objThis.Engine Is Nothing 
            If DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                'scDS.SplitterDistance = SplitSrc
                Pnt.Y = Pnt.Y + szProp + szExtProp + szSrc
                gbDel.Location = Pnt
                Pnt.Y = Pnt.Y + szDel
                gbField.Location = Pnt
            Else
                gbDel.Visible = False
                'scDS.SplitterDistance = SplitSrcNoDel
                Pnt.Y = Pnt.Y + szProp + szExtProp + szSrc
                gbField.Location = Pnt
                gbField.Height = szFld + szDel
            End If
            gbSource.Visible = True
            gbTarget.Visible = False
            gbStruct.Height = SplitTgt - 5
            gbStruct.Update()
            gbStruct.Refresh()
            cmbCharacterCode.Enabled = True
            txtException.Enabled = True
            If objThis.IsLookUp = True Then
                Label1.Text = "Lookup Datastore Definition"
                gbProp.Text = "Lookup Datastore Properties"
                Me.Text = "Lookup Datastore Properties"
            Else
                Label1.Text = "Source Definition"
                gbProp.Text = "Source Properties"
                Me.Text = "Source Datastore"
            End If
        Else
            If DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                'scDS.SplitterDistance = SplitTgt
                Pnt.Y = Pnt.Y + szProp + szExtProp + szTgt
                gbDel.Location = Pnt
                Pnt.Y = Pnt.Y + szDel
                gbField.Location = Pnt
                gbField.Height = szFld + (szSrc - szTgt)
            Else
                gbDel.Visible = False
                'scDS.SplitterDistance = SplitTgtNoDel
                Pnt.Y = Pnt.Y + szProp + szExtProp + szTgt
                gbField.Location = Pnt
                gbField.Height = szFld + szDel + (szSrc - szTgt)
            End If
            gbSource.Visible = False
            gbTarget.Visible = True
            gbTarget.Location = gbSource.Location
            gbStruct.Height = SplitTgt - 81 - 5
            gbStruct.Update()
            gbStruct.Refresh()
            cmbCharacterCode.Enabled = False
            txtException.Enabled = True
            rbLookup.Visible = False
            Label1.Text = "Target Definition"
            gbProp.Text = "Target Properties"
            Me.Text = "Target Datastore"
        End If

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            gbExtProps.Enabled = True
        Else
            gbExtProps.Enabled = False
            txtPoll.Text = ""
        End If

        SetAccessCombo(DatastoreType)

        'Trigger based cdc
        '//AccessMethod is Enable for all except Relational,IMS,VSAM, Trigger based
        cmbAccessMethod.Enabled = Not (DatastoreType = enumDatastore.DS_RELATIONAL Or _
        DatastoreType = enumDatastore.DS_VSAM Or _
        DatastoreType = enumDatastore.DS_IMSDB Or _
        DatastoreType = enumDatastore.DS_INCLUDE)

        lblUOW.Enabled = (DatastoreType = modDeclares.enumDatastore.DS_DB2CDC) 'or _
        '                    DatastoreType = modDeclares.enumDatastore.DS_TRBCDC Or _
        '                    DatastoreType = modDeclares.enumDatastore.DS_IMSCDC)

        txtUOW.Enabled = (DatastoreType = modDeclares.enumDatastore.DS_DB2CDC) 'Or _
        '                    DatastoreType = modDeclares.enumDatastore.DS_TRBCDC Or _
        '                    DatastoreType = modDeclares.enumDatastore.DS_IMSCDC)

        SetPortOrMQ()

        'Set Character Code by 
        SetCharCodeCombo(DatastoreType)

        If objThis.Engine Is Nothing Then
            cmbOperationType.Enabled = True
        Else
            cmbOperationType.Enabled = (objThis.DsDirection = DS_DIRECTION_TARGET)
        End If

        SetOperationTypeCombo()

        '//Only IBM Event : new added by npatel on 8/10/05
        'chkBeforeImage.Enabled = (objThis.DatastoreType = modDeclares.enumDatastore.DS_IBMEVENT)

        '//only for IMS and IMSCDC
        chkIMSPathData.Enabled = DatastoreType = modDeclares.enumDatastore.DS_IMSDB 'Or _
        'objThis.DatastoreType = modDeclares.enumDatastore.DS_IMSCDC Or _
        'objThis.DatastoreType = enumDatastore.DS_IMSLE Or _
        'objThis.DatastoreType = enumDatastore.DS_IMSLEBATCH)
        '//only for relational 
        'chkCmmtKey.Enabled = (objThis.DatastoreType = modDeclares.enumDatastore.DS_RELATIONAL)


        '//only for CDC
        chkSkipChangeCheck.Enabled = (DatastoreType = modDeclares.enumDatastore.DS_UTSCDC Or _
                                  DatastoreType = modDeclares.enumDatastore.DS_VSAMCDC Or _
                                  DatastoreType = enumDatastore.DS_SUBVAR Or _
                                  DatastoreType = enumDatastore.DS_ORACLECDC Or _
                                  DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                                  DatastoreType = enumDatastore.DS_IMSCDC Or _
                                  DatastoreType = modDeclares.enumDatastore.DS_DB2CDC)

        SetPoll()
        'cmdShowHideFieldAttr.Enabled = False

        '//added by npatel on 9/6/05
        SetListItemByValue(cmbDatastoreType, objThis.DatastoreType, False)

        If objThis.DsDirection = DS_DIRECTION_SOURCE Or objThis.Engine Is Nothing Then
            gbSource.Enabled = True
        End If

    End Function

    Sub SetPortOrMQ()

        txtPortOrMQMgr.Enabled = ((cmbAccessMethod.Enabled = True) And (cmbAccessMethod.Text = "IP"))
        lblPortorMQMgr.Enabled = ((cmbAccessMethod.Enabled = True) And (cmbAccessMethod.Text = "IP"))

    End Sub

    Sub SetPoll()

        txtPoll.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
                          objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
                          objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
                          objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                          objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
                          objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
                          cmbAccessMethod.Text = "VSAM CDCStore") Or _
                          objThis.DatastoreType = enumDatastore.DS_UTSCDC

        lblPoll.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
                            objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
                            cmbAccessMethod.Text = "VSAM CDCStore") Or _
                            objThis.DatastoreType = enumDatastore.DS_UTSCDC

        txtRestart.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
                            objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
                            cmbAccessMethod.Text = "VSAM CDCStore") Or _
                            objThis.DatastoreType = enumDatastore.DS_UTSCDC

        lblRestart.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
                            objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                            objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
                            objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
                            cmbAccessMethod.Text = "VSAM CDCStore") Or _
                            objThis.DatastoreType = enumDatastore.DS_UTSCDC

    End Sub

    Sub SetAccessCombo(ByVal DStype As enumDatastore)

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            Select Case DStype
                Case enumDatastore.DS_BINARY, enumDatastore.DS_TEXT, enumDatastore.DS_DELIMITED, enumDatastore.DS_HSSUNLOAD, _
                enumDatastore.DS_XML, enumDatastore.DS_DB2LOAD, enumDatastore.DS_IBMEVENT
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("IP", DS_ACCESSMETHOD_IP))
                    cmbAccessMethod.SelectedIndex = 0

                Case enumDatastore.DS_DB2CDC, enumDatastore.DS_UTSCDC, enumDatastore.DS_IMSCDC, enumDatastore.DS_IMSCDCLE, _
                enumDatastore.DS_ORACLECDC, enumDatastore.DS_VSAMCDC, enumDatastore.DS_SUBVAR
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("IP", DS_ACCESSMETHOD_IP))
                    cmbAccessMethod.Items.Add(New Mylist("VSAM CDCStore", DS_ACCESSMETHOD_VSAM))
                    cmbAccessMethod.Items.Add(New Mylist("SQD CDCStore", DS_ACCESSMETHOD_SQDCDC))
                    cmbAccessMethod.SelectedIndex = 0
            End Select
        Else
            cmbAccessMethod.Items.Clear()
            cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
            cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
            cmbAccessMethod.Items.Add(New Mylist("IP", DS_ACCESSMETHOD_IP))
            cmbAccessMethod.SelectedIndex = 0
        End If

        SetListItemByValue(cmbAccessMethod, objThis.DsAccessMethod, False)

    End Sub

    Sub SetCharCodeCombo(ByVal DatastoreType As enumDatastore)

        Try
            cmbCharacterCode.Items.Clear()

            If DatastoreType = enumDatastore.DS_HSSUNLOAD Then
                cmbCharacterCode.Items.Add(New Mylist("EBCDIC", DS_CHARACTERCODE_EBCDIC))

                If IsNewObj = True Then
                    cmbCharacterCode.SelectedIndex = 0
                End If

            ElseIf DatastoreType = enumDatastore.DS_DB2LOAD Or _
                DatastoreType = enumDatastore.DS_IMSDB Or _
                DatastoreType = enumDatastore.DS_IMSCDC Or _
                DatastoreType = enumDatastore.DS_DB2CDC Or _
                DatastoreType = enumDatastore.DS_VSAMCDC Or _
                DatastoreType = enumDatastore.DS_ORACLECDC Or _
                DatastoreType = enumDatastore.DS_UTSCDC Then

                cmbCharacterCode.Items.Add(New Mylist("System", DS_CHARACTERCODE_SYSTEM))
                cmbCharacterCode.Items.Add(New Mylist("ASCII", DS_CHARACTERCODE_ASCII))
                cmbCharacterCode.Items.Add(New Mylist("EBCDIC", DS_CHARACTERCODE_EBCDIC))

                If IsNewObj = True Then
                    cmbCharacterCode.SelectedIndex = 2 '//EBCDIC
                End If

            Else
                cmbCharacterCode.Items.Add(New Mylist("System", DS_CHARACTERCODE_SYSTEM))
                cmbCharacterCode.Items.Add(New Mylist("ASCII", DS_CHARACTERCODE_ASCII))
                cmbCharacterCode.Items.Add(New Mylist("EBCDIC", DS_CHARACTERCODE_EBCDIC))

                If IsNewObj = True Then
                    cmbCharacterCode.SelectedIndex = 0 '//System (Default)
                End If

            End If

            SetListItemByValue(cmbCharacterCode, objThis.DsCharacterCode, False)

        Catch ex As Exception
            LogError(ex, "frmDatastore SetCharCodeCombo")
        End Try

    End Sub

    Sub SetOperationTypeCombo()

        Select Case objThis.DatastoreType

            Case enumDatastore.DS_DB2CDC, enumDatastore.DS_UTSCDC, enumDatastore.DS_IMSCDC, enumDatastore.DS_IMSCDCLE, _
            enumDatastore.DS_ORACLECDC, enumDatastore.DS_VSAMCDC, enumDatastore.DS_SUBVAR, enumDatastore.DS_RELATIONAL, _
            enumDatastore.DS_IMSDB
                cmbOperationType.Items.Clear()
                cmbOperationType.Items.Add(New Mylist("", ""))
                cmbOperationType.Items.Add(New Mylist("Insert", DS_OPERATION_INSERT))
                cmbOperationType.Items.Add(New Mylist("Update", DS_OPERATION_UPDATE))
                cmbOperationType.Items.Add(New Mylist("Delete", DS_OPERATION_DELETE))
                cmbOperationType.Items.Add(New Mylist("Change", DS_OPERATION_CHANGE))
                cmbOperationType.Items.Add(New Mylist("Modify", DS_OPERATION_MODIFY))
                cmbOperationType.Items.Add(New Mylist("Replace", DS_OPERATION_REPLACE))

            Case Else
                cmbOperationType.Items.Clear()
                cmbOperationType.Items.Add(New Mylist("", ""))
                cmbOperationType.Items.Add(New Mylist("Insert", DS_OPERATION_INSERT))
                'cmbOperationType.Items.Add(New Mylist("Update", DS_OPERATION_UPDATE))
                'cmbOperationType.Items.Add(New Mylist("Delete", DS_OPERATION_DELETE))
                'cmbOperationType.Items.Add(New Mylist("Change", DS_OPERATION_CHANGE))
                'cmbOperationType.Items.Add(New Mylist("Modify", DS_OPERATION_MODIFY))
                cmbOperationType.Items.Add(New Mylist("Replace", DS_OPERATION_REPLACE))
        End Select

        If IsNewObj = True Then cmbOperationType.SelectedIndex = 0
        SetListItemByValue(cmbOperationType, objThis.OperationType, False)

    End Sub

    Private Sub cbIfSpaceChar_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIfSpaceChar.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.ExtTypeChar = CType(cbIfSpaceChar.Items(cbIfSpaceChar.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Private Sub cbIfSpaceNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIfSpaceNum.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.ExtTypeNum = CType(cbIfSpaceNum.Items(cbIfSpaceNum.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Private Sub cbIfNullChar_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIfNullChar.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.IfNullChar = CType(cbIfNullChar.Items(cbIfNullChar.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Private Sub cbIfNullNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIfNullNum.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.IfNullNum = CType(cbIfNullNum.Items(cbIfNullNum.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Private Sub cbInvalidChar_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbInvalidChar.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.InValidChar = CType(cbInvalidChar.Items(cbInvalidChar.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Private Sub cbInvalidNum_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbInvalidNum.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.InValidNum = CType(cbInvalidNum.Items(cbInvalidNum.SelectedIndex), Mylist).ItemData
        OnChange(Me, New EventArgs)

    End Sub

    Function UpdateControls() As Boolean

        'If pnlSelect.Visible = True Then
        '    'cmdShowHideFieldAttr.Text = "Show Fields/Attributes"
        'Else
        '    'cmdShowHideFieldAttr.Text = "Hide Fields/Attributes"
        'End If

    End Function

    Function InitControls() As Boolean

        txtDatastoreName.Focus()

        Me.rbLookup.Checked = objThis.IsLookUp

        Me.tvDatastoreStructures.ImageList = modDeclares.imgListSmall
        tvDatastoreStructures.Font = New Font("Arial", 8, FontStyle.Bold)

        '//Add listview columns
        tvFields.HideSelection = False

        lvFieldAttrs.SmallImageList = imgListSmall

        'lvFieldAttrs.SuspendLayout()
        'lvFieldAttrs.View = View.Details
        'lvFieldAttrs.Columns.Add("Attribute", 100, HorizontalAlignment.Left)
        'lvFieldAttrs.Columns.Add("Value", 150, HorizontalAlignment.Left)
        'lvFieldAttrs.Columns.Add("", 0, HorizontalAlignment.Left)
        lvFieldAttrs.FullRowSelect = True

        'lvFieldAttrs.ResumeLayout()

        'pnlSelect.Visible = True
        'pnlFields.Visible = True
        'cmdShowHideFieldAttr.Text = "Show Fields/Attributes"

        'pnlFields.Size = pnlSelect.Size
        'pnlFields.Location = pnlSelect.Location


        'cmbOperationType.Items.Add(New Mylist("Insert", DS_OPERATION_INSERT))
        ''cmbOperationType.Items.Add(New Mylist("Update", DS_OPERATION_UPDATE))
        ''cmbOperationType.Items.Add(New Mylist("Delete", DS_OPERATION_DELETE))
        ''cmbOperationType.Items.Add(New Mylist("Change", DS_OPERATION_CHANGE))
        ''cmbOperationType.Items.Add(New Mylist("Modify", DS_OPERATION_MODIFY))
        'cmbOperationType.Items.Add(New Mylist("Replace", DS_OPERATION_REPLACE))

        'If IsNewObj = True Then cmbOperationType.SelectedIndex = 0

        '//Modified by TK  12/2009
        cmbDatastoreType.Items.Clear()
        cmbDatastoreType.Items.Add(New Mylist("Binary", enumDatastore.DS_BINARY))
        cmbDatastoreType.Items.Add(New Mylist("Text", enumDatastore.DS_TEXT))
        cmbDatastoreType.Items.Add(New Mylist("Delimited", enumDatastore.DS_DELIMITED))
        cmbDatastoreType.Items.Add(New Mylist("XML", enumDatastore.DS_XML))
        cmbDatastoreType.Items.Add(New Mylist("Relational", enumDatastore.DS_RELATIONAL))
        cmbDatastoreType.Items.Add(New Mylist("DB2LOAD", enumDatastore.DS_DB2LOAD))
        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            cmbDatastoreType.Items.Add(New Mylist("HSSUNLOAD", enumDatastore.DS_HSSUNLOAD))
        End If
        cmbDatastoreType.Items.Add(New Mylist("IMSDB", enumDatastore.DS_IMSDB))
        cmbDatastoreType.Items.Add(New Mylist("VSAM", enumDatastore.DS_VSAM))
        cmbDatastoreType.Items.Add(New Mylist("IMSCDC", enumDatastore.DS_IMSCDC))
        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            cmbDatastoreType.Items.Add(New Mylist("IMSCDC LE", enumDatastore.DS_IMSCDCLE))
        End If
        cmbDatastoreType.Items.Add(New Mylist("DB2CDC", enumDatastore.DS_DB2CDC))
        cmbDatastoreType.Items.Add(New Mylist("VSAMCDC", enumDatastore.DS_VSAMCDC))
        cmbDatastoreType.Items.Add(New Mylist("IBM Event", enumDatastore.DS_IBMEVENT))
        cmbDatastoreType.Items.Add(New Mylist("OracleCDC", enumDatastore.DS_ORACLECDC))
        cmbDatastoreType.Items.Add(New Mylist("SQD CDC", enumDatastore.DS_UTSCDC))
        '
        'cmbDatastoreType.Items.Add(New Mylist("XML CDC", enumDatastore.DS_XMLCDC))
        'cmbDatastoreType.Items.Add(New Mylist("Trigger based CDC", enumDatastore.DS_TRBCDC))
        'cmbDatastoreType.Items.Add(New Mylist("IMS LE Batch", enumDatastore.DS_IMSLEBATCH))
        cmbDatastoreType.Items.Add(New Mylist("SubVariable", enumDatastore.DS_SUBVAR))

        '//
        'cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
        'cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
        'cmbAccessMethod.Items.Add(New Mylist("IP", DS_ACCESSMETHOD_IP))
        'cmbAccessMethod.Items.Add(New Mylist("VSAM CDCStore", DS_ACCESSMETHOD_VSAM))
        'If IsNewObj = True Then cmbAccessMethod.SelectedIndex = 0

        '//


        '//
        setColCombo()
        setRowCombo()
        setTextCombo()



        SetCombos()

        UpdateControls()

    End Function

    Function ValidateSelection() As Boolean

        'ValidateSelection = True
        'If tvDatastoreStructures.GetNodeCount(True) <= 0 Then
        '    Exit Function
        'End If

        If tvDatastoreStructures.GetNodeCount(True) <= 0 Then
            ValidateSelection = False
            Exit Function
        End If
        ValidateSelection = True

    End Function

    '//This will load treeview of fields
    '//If New Datastore then fields will be from Structure
    '//If Edit Datastore then fields wil be from database
    Function FillFieldAttr() As Boolean
        Dim objSel As clsStructureSelection
        Dim dsSel As clsDSSelection
        Dim isfound As Boolean = False
        Dim i As Integer
        'Dim fldNode As TreeNode

        Try
            tvFields.BeginUpdate()
            lvFieldAttrs.BeginUpdate()
            tvFields.Nodes.Clear()

            txtListViewText.Visible = False
            cmbListViewCombo.Visible = False

            '//Load all fields from structure or datastore
            objSel = CType(tvDatastoreStructures.SelectedNode.Tag, clsStructureSelection)
            For i = 0 To objThis.ObjSelections.Count - 1
                dsSel = objThis.ObjSelections(i)
                If objSel Is dsSel.ObjSelection Then
                    isfound = True
                    dsSel.LoadMe()
                    AddFieldsToTreeView(dsSel, , tvFields)
                    Exit For
                End If
            Next

            If isfound = False Then
                objSel.LoadMe()
                AddFieldsToTreeView(objSel, , tvFields, True)
            End If
            HiLiteFieldDescNodes(tvFields.Nodes(0), True, tvFields)
            HiLiteFieldFKeyNodes(tvFields.Nodes(0), True, tvFields)
            HiLiteFieldKeyNodes(tvFields.Nodes(0), True, tvFields)

        Catch ex As Exception
            LogError(ex)
        Finally
            tvFields.EndUpdate()
            lvFieldAttrs.EndUpdate()
        End Try

    End Function

    Function ShowFieldAttributes(ByVal cNode As TreeNode) As Boolean

        Try
            If cNode Is Nothing Then
                Exit Function
            End If

            lvFieldAttrs.BeginUpdate()

            lvFieldAttrs.GridLines = True
            lvFieldAttrs.Items.Clear()

            lvFieldAttrs.Items.Add(modDeclares.TXT_ISKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
            lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY))
            lvFieldAttrs.Items.Add(modDeclares.TXT_INITVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INITVAL))
            lvFieldAttrs.Items.Add(modDeclares.TXT_RETYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_RETYPE))

            lvFieldAttrs.Items.Add(modDeclares.TXT_EXTTYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_EXTTYPE))
            lvFieldAttrs.Items.Add(modDeclares.TXT_INVALID).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INVALID))

            lvFieldAttrs.Items.Add(modDeclares.TXT_DATEFORMAT).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATEFORMAT))
            lvFieldAttrs.Items.Add(modDeclares.TXT_LABEL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LABEL))

            lvFieldAttrs.Items.Add(modDeclares.TXT_IDENTVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_IDENTVAL))
            lvFieldAttrs.Items.Add(modDeclares.TXT_NCHILDREN).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_NCHILDREN))
            lvFieldAttrs.Items.Add(modDeclares.TXT_LEVEL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LEVEL))
            lvFieldAttrs.Items.Add(modDeclares.TXT_TIMES).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_TIMES))
            lvFieldAttrs.Items.Add(modDeclares.TXT_OCCURS).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OCCURS))
            lvFieldAttrs.Items.Add(modDeclares.TXT_DATATYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATATYPE))
            lvFieldAttrs.Items.Add(modDeclares.TXT_OFFSET).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_OFFSET))
            lvFieldAttrs.Items.Add(modDeclares.TXT_LENGTH).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LENGTH))
            lvFieldAttrs.Items.Add(modDeclares.TXT_SCALE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_SCALE))
            lvFieldAttrs.Items.Add(modDeclares.TXT_CANNULL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_CANNULL))

            Dim f As New System.Drawing.Font(lvFieldAttrs.Font, FontStyle.Bold)
            Dim i As Integer

            For i = 0 To lvFieldAttrs.Items.Count - 1
                lvFieldAttrs.Items(i).UseItemStyleForSubItems = False
                lvFieldAttrs.Items(i).SubItems(0).Font = f

                'If i < 9 Then
                '    lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Black
                '    lvFieldAttrs.Items(i).ImageIndex = 25
                'Else
                lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).SubItems(1).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).ImageIndex = 30
                'End If

            Next
            'lblFieldName.Text = cNode.Text

            lvFieldAttrs.EndUpdate()

            Return True

        Catch ex As Exception
            LogError(ex, "frmDatastore ShowFieldAttribute")
            Return False
        End Try

    End Function

    Function FillStructure(ByVal cNode As TreeNode, Optional ByVal IsDSstructure As Boolean = False) As Boolean

        Try
            Dim cStructFolder As TreeNode
            Dim activenode As INode = Nothing
            Dim NodeType As String = CType(cNode.Tag, INode).Type

            If IsDSstructure = True Then
                activenode = CType(cNode.Tag, INode)
            End If

            Select Case NodeType
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    cNode = CType(cNode.Tag, clsDatastore).Engine.ObjSystem.Environment.ObjTreeNode
                Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                    cNode = CType(cNode.Tag, clsDSSelection).ObjDatastore.Engine.ObjSystem.Environment.ObjTreeNode
                Case NODE_FO_SOURCEDATASTORE, NODE_FO_TARGETDATASTORE
                    cNode = CType(cNode.Tag, clsFolderNode).Parent.ObjTreeNode
                    cNode = CType(cNode.Tag, clsEngine).ObjSystem.Environment.ObjTreeNode
                Case NODE_FO_DATASTORE
                    cNode = CType(cNode.Tag, clsFolderNode).Parent.ObjTreeNode
                Case DS_UNKNOWN, DS_BINARY, DS_TEXT, DS_DELIMITED, DS_XML, DS_RELATIONAL, DS_DB2LOAD, DS_HSSUNLOAD, DS_IMSDB, DS_VSAM, DS_IMSCDCLE, DS_DB2CDC, DS_VSAMCDC, DS_IBMEVENT, DS_ORACLECDC, DS_GENERICCDC, DS_SUBVAR, DS_INCLUDE, DS_IMSCDC ', DS_IMSLEBATCH, DS_XMLCDC, DS_TRBCDC,
                    cNode = CType(cNode.Tag, clsFolderNode).Parent.ObjTreeNode
                    'cNode = CType(cNode.Tag, clsFolderNode).Parent.ObjTreeNode
                Case Else
                    Return False
                    Exit Function
            End Select

            cStructFolder = cNode.Nodes(1)

            AddChildrenNoFolder(cStructFolder.Clone, tvDatastoreStructures.Nodes)
            tvDatastoreStructures.CollapseAll()


            '//Now load selected fields and check if object is already created and user is editing it
            If IsNewObj = False Then
                objThis.LoadItems(True)
                CheckSelected()
                If IsDSstructure = True Then
                    tvDatastoreStructures.SelectedNode = SelectFirstMatchingNode(tvDatastoreStructures, activenode.Text)
                End If
            End If

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Sub AddChildrenNoFolder(ByVal cNodeSource As TreeNode, ByVal cNodeTarget As TreeNodeCollection)

        Dim Snd As TreeNode
        Dim newNode As TreeNode
        Dim obj As INode

        For Each Snd In cNodeSource.Nodes

            obj = CType(Snd.Tag, INode) '.Clone(objThis)
            '//only insert structure and selection skip folders
            '//now only add applicable entries in treeview
            '//e.g. for Binary DS all structure but for XML DS only DTD is valid structure so only add DTD structures and subselection
            '//If not valid structure or selection then no need to find their children coz they will be the same type
            If IsValidStructureForDS(obj) = True Then
                newNode = cNodeTarget.Add(Snd.Text)
                If obj.Type = NODE_STRUCT = True Then
                    newNode.Tag = CType(obj, clsStructure).SysAllSelection
                Else
                    newNode.Tag = obj
                End If
                newNode.ImageIndex = Snd.ImageIndex
                newNode.SelectedImageIndex = Snd.SelectedImageIndex
                '// recurse for children
                AddChildrenNoFolder(Snd, newNode.Nodes)
            End If
        Next

    End Sub

    Function CheckSelected() As Boolean
        '//TODO
        Dim i As Integer
        Dim cnt As Integer = tvDatastoreStructures.GetNodeCount(True)

        If cnt <= 0 Then Exit Function
        Try
            tvDatastoreStructures.BeginUpdate()
            '//Looop all selected field and check that field in treeview
            For i = 0 To objThis.ObjSelections.Count - 1
                SelectFirstMatchingNode(tvDatastoreStructures, "", True, CType(objThis.ObjSelections(i), clsDSSelection).ObjSelection.Key)
            Next
        Catch ex As Exception
            LogError(ex)
        Finally
            tvDatastoreStructures.EndUpdate()
        End Try

    End Function

    Sub ShowLVCombo(ByVal lvItem As ListViewItem)

        txtListViewText.Visible = False

        IsEventFromCode = True
        SetListViewCombo(lvItem.Index)
        IsEventFromCode = False

        ' Set default text for ComboBox to match the item that is clicked.
        IsEventFromCode = True
        Me.cmbListViewCombo.Text = lvItem.SubItems(1).Text
        IsEventFromCode = False

        ' Display the ComboBox, and make sure that it is on top with focus.
        Me.cmbListViewCombo.Visible = True
        Me.cmbListViewCombo.BringToFront()
        Me.cmbListViewCombo.Focus()

    End Sub

    Sub ShowLVText(ByVal lvItem As ListViewItem)

        cmbListViewCombo.Visible = False
        ' Set default text for ComboBox to match the item that is clicked.

        IsEventFromCode = True
        Me.txtListViewText.Text = lvItem.SubItems(1).Text
        IsEventFromCode = False

        ' Display the ComboBox, and make sure that it is on top with focus.
        Me.txtListViewText.Visible = True
        Me.txtListViewText.BringToFront()
        Me.txtListViewText.Focus()

    End Sub

    Sub SetListViewCombo(ByVal Attribute As enumFieldAttributes)

        Dim i As Integer
        Dim j As Integer
        Dim x As Integer
        Dim k As Integer
        Dim selName As String
        Dim isKey As String
        Dim keyName As String

        IsEventFromCode = True
        cmbListViewCombo.Items.Clear()
        IsEventFromCode = False

        IsEventFromCode = True

        Select Case Attribute
            Case modDeclares.enumFieldAttributes.ATTR_FKEY
                '// ADDED by TK December 2006

                Dim objSel As clsDSSelection
                Dim tempstring As String = ""
                Dim tempobj As clsDSSelection
                Dim childSel As clsDSSelection = Nothing
                Dim templist As New ArrayList  '/exclusion list of selections
                Dim StrSelName As String
                Dim fld As clsField
                Dim flag As Boolean = False

                '/// initialize combo list with blank
                cmbListViewCombo.Items.Add(New Mylist("", ""))

                '// first see if the selection is checked
                If tvDatastoreStructures.SelectedNode.Checked = True Then
                    tempstring = tvDatastoreStructures.SelectedNode.Text
                    '/// add the current node in the exclusion list
                    templist.Add(tempstring)

                    '/// first see if the selected structure is already a datastore selection
                    If (Not tempstring = "") Then
                        For k = 0 To objThis.ObjSelections.Count - 1
                            StrSelName = objThis.ObjSelections(k).selectionname
                            If tempstring = StrSelName Then
                                childSel = objThis.ObjSelections(k)
                                Exit For
                            End If
                        Next
                    End If
                    '// now make a list of all dsselection names that are children of the 
                    '// DSselection and all it's children's children, etc...
                    If Not childSel Is Nothing Then
recurse:                For x = 0 To childSel.ObjDSSelections.Count - 1
                            tempobj = childSel.ObjDSSelections(x)
                            templist.Add(tempobj.SelectionName)
                            If tempobj.ObjDSSelections.Count > 0 Then
                                childSel = tempobj
                                GoTo recurse
                            End If
                        Next
                    End If
                    '/// now go through the all the datastore selections and add all
                    '/// key fields of all Non-child selections of this selection
                    For i = 0 To objThis.ObjSelections.Count - 1
                        objSel = objThis.ObjSelections(i)
                        selName = objSel.SelectionName
                        If Not templist.Contains(selName) Then
                            For j = 1 To objSel.DSSelectionFields.Count
                                fld = CType(objSel.DSSelectionFields(j - 1), clsField)
                                isKey = fld.GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY)
                                If (isKey = "Yes") Or (isKey = "1") Then
                                    keyName = selName & "." & fld.FieldName
                                    cmbListViewCombo.Items.Add(New Mylist(keyName, keyName))
                                End If
                            Next
                        End If
                    Next
                End If


            Case modDeclares.enumFieldAttributes.ATTR_ISKEY
                cmbListViewCombo.Items.Add(New Mylist("Yes", "YES"))
                cmbListViewCombo.Items.Add(New Mylist("No", "NO"))

            Case modDeclares.enumFieldAttributes.ATTR_EXTTYPE
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("ALPHA", "ALPHA"))
                cmbListViewCombo.Items.Add(New Mylist("ALNUM", "ALNUM"))
                cmbListViewCombo.Items.Add(New Mylist("CIYYMMDD", "CIYYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("DDMMYY", "DDMMYY"))
                cmbListViewCombo.Items.Add(New Mylist("DDMMYYHHMM", "DDMMYYHHMM"))
                cmbListViewCombo.Items.Add(New Mylist("DDMMYYHHMMSS", "DDMMYYHHMMSS"))
                cmbListViewCombo.Items.Add(New Mylist("EURDATE", "EURDATE"))
                cmbListViewCombo.Items.Add(New Mylist("HHMM", "HHMM"))
                cmbListViewCombo.Items.Add(New Mylist("HHMMSS", "HHMMSS"))
                cmbListViewCombo.Items.Add(New Mylist("ISODATE", "ISODATE"))
                cmbListViewCombo.Items.Add(New Mylist("MMDDYY", "MMDDYY"))
                cmbListViewCombo.Items.Add(New Mylist("MMDDYYHHMM", "MMDDYYHHMM"))
                cmbListViewCombo.Items.Add(New Mylist("MMDDYYHHMMSS", "MMDDYYHHMMSS"))
                cmbListViewCombo.Items.Add(New Mylist("USADATE", "USADATE"))
                cmbListViewCombo.Items.Add(New Mylist("YYDDD", "YYDDD"))
                cmbListViewCombo.Items.Add(New Mylist("YYDDMM", "YYDDMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYMMDD", "YYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("YYMMDDHHMM", "YYMMDDHHMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYMMDDHHMMSS", "YYMMDDHHMMSS"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYDDD", "YYYYDDD"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYDDMM", "YYYYDDMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYMMDD", "YYYYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYMMDDHHMM", "YYYYMMDDHHMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYMMDDHHMMSS", "YYYYMMDDHHMMSS"))

            Case modDeclares.enumFieldAttributes.ATTR_RETYPE
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("BINARY", "BINARY"))
                cmbListViewCombo.Items.Add(New Mylist("CHAR", "CHAR"))
                cmbListViewCombo.Items.Add(New Mylist("DATE", "DATE"))
                cmbListViewCombo.Items.Add(New Mylist("DECIMAL", "DECIMAL"))
                cmbListViewCombo.Items.Add(New Mylist("INTEGER", "DECIMAL"))
                cmbListViewCombo.Items.Add(New Mylist("SMALLINT", "SMALLINT"))
                cmbListViewCombo.Items.Add(New Mylist("TEXTNUM", "TEXTNUM"))
                cmbListViewCombo.Items.Add(New Mylist("TIME", "TIME"))
                cmbListViewCombo.Items.Add(New Mylist("TIMESTAMP", "TIMESTAMP"))
                cmbListViewCombo.Items.Add(New Mylist("VARCHAR", "VARCHAR"))
                cmbListViewCombo.Items.Add(New Mylist("ZONE", "ZONE"))
                cmbListViewCombo.Items.Add(New Mylist("XMLATTC", "XMLATTC"))
                cmbListViewCombo.Items.Add(New Mylist("XMLATTPC", "XMLATTPC"))
                cmbListViewCombo.Items.Add(New Mylist("XMLCDATA", "XMLCDATA"))
                cmbListViewCombo.Items.Add(New Mylist("XMLPCDATA", "XMLPCDATA"))

            Case modDeclares.enumFieldAttributes.ATTR_INVALID
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("SETNULL", "SETNULL"))
                cmbListViewCombo.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
                cmbListViewCombo.Items.Add(New Mylist("SETZERO", "SETZERO"))
                cmbListViewCombo.Items.Add(New Mylist("SETEXP", "SETEXP"))

        End Select
        IsEventFromCode = False
    End Sub

    Sub SetCombos()

        Try
            'Set ExtType Combo
            cbIfSpaceChar.Items.Clear()
            cbIfSpaceChar.Items.Add(New Mylist("", ""))
            cbIfSpaceChar.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbIfSpaceChar.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbIfSpaceChar.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbIfSpaceChar.Items.Add(New Mylist("SETEXP", "SETEXP"))
            'cbIfSpace.Items.Add(New Mylist("DDMMYYHHMM", "DDMMYYHHMM"))
            'cbIfSpace.Items.Add(New Mylist("DDMMYYHHMMSS", "DDMMYYHHMMSS"))
            'cbIfSpace.Items.Add(New Mylist("EURDATE", "EURDATE"))
            'cbIfSpace.Items.Add(New Mylist("HHMM", "HHMM"))
            'cbIfSpace.Items.Add(New Mylist("HHMMSS", "HHMMSS"))
            'cbIfSpace.Items.Add(New Mylist("ISODATE", "ISODATE"))
            'cbIfSpace.Items.Add(New Mylist("MMDDYY", "MMDDYY"))
            'cbIfSpace.Items.Add(New Mylist("MMDDYYHHMM", "MMDDYYHHMM"))
            'cbIfSpace.Items.Add(New Mylist("MMDDYYHHMMSS", "MMDDYYHHMMSS"))
            'cbIfSpace.Items.Add(New Mylist("USADATE", "USADATE"))
            'cbIfSpace.Items.Add(New Mylist("YYDDD", "YYDDD"))
            'cbIfSpace.Items.Add(New Mylist("YYDDMM", "YYDDMM"))
            'cbIfSpace.Items.Add(New Mylist("YYMMDD", "YYMMDD"))
            'cbIfSpace.Items.Add(New Mylist("YYMMDDHHMM", "YYMMDDHHMM"))
            'cbIfSpace.Items.Add(New Mylist("YYMMDDHHMMSS", "YYMMDDHHMMSS"))
            'cbIfSpace.Items.Add(New Mylist("YYYYDDD", "YYYYDDD"))
            'cbIfSpace.Items.Add(New Mylist("YYYYDDMM", "YYYYDDMM"))
            'cbIfSpace.Items.Add(New Mylist("YYYYMMDD", "YYYYMMDD"))
            'cbIfSpace.Items.Add(New Mylist("YYYYMMDDHHMM", "YYYYMMDDHHMM"))
            'cbIfSpace.Items.Add(New Mylist("YYYYMMDDHHMMSS", "YYYYMMDDHHMMSS"))

            'Set ExtAll Combo
            cbIfSpaceNum.Items.Clear()
            cbIfSpaceNum.Items.Add(New Mylist("", ""))
            cbIfSpaceNum.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbIfSpaceNum.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbIfSpaceNum.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbIfSpaceNum.Items.Add(New Mylist("SETEXP", "SETEXP"))

            'Set IfNull Combo
            cbIfNullChar.Items.Clear()
            cbIfNullChar.Items.Add(New Mylist("", ""))
            cbIfNullChar.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbIfNullChar.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbIfNullChar.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbIfNullChar.Items.Add(New Mylist("SETEXP", "SETEXP"))

            'Set IfnAll Combo
            cbIfNullNum.Items.Clear()
            cbIfNullNum.Items.Add(New Mylist("", ""))
            cbIfNullNum.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbIfNullNum.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbIfNullNum.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbIfNullNum.Items.Add(New Mylist("SETEXP", "SETEXP"))

            'Set Invalid Combo
            cbInvalidChar.Items.Clear()
            cbInvalidChar.Items.Add(New Mylist("", ""))
            cbInvalidChar.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbInvalidChar.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbInvalidChar.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbInvalidChar.Items.Add(New Mylist("SETEXP", "SETEXP"))
            'cbInvalid.Items.Add(New Mylist("0001-01-01", "0001-01-01"))
            'cbInvalid.Items.Add(New Mylist("00.00.00", "00.00.00"))

            'Set InvAll Combo
            cbInvalidNum.Items.Clear()
            cbInvalidNum.Items.Add(New Mylist("", ""))
            cbInvalidNum.Items.Add(New Mylist("SETNULL", "SETNULL"))
            cbInvalidNum.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
            cbInvalidNum.Items.Add(New Mylist("SETZERO", "SETZERO"))
            cbInvalidNum.Items.Add(New Mylist("SETEXP", "SETEXP"))

        Catch ex As Exception
            LogError(ex, "frmDatastore SetCombos")
        End Try

    End Sub

    '//This function check structure type and selected datastore type and if applicatble structure then returns true so we can add into the treeview
    Private Function IsValidStructureForDS(ByVal obj As INode) As Boolean

        IsValidStructureForDS = True '//Change : lets allow all type of structures to be selected

    End Function

    '// Added 12/09 to see if any fields in relational targets are key fields.
    '// If it finds a key, it returns True, otherwise returns false
    Function HasKeyForRel() As Boolean

        Try
            For Each sel As clsDSSelection In objThis.ObjSelections
                For Each fld As clsField In sel.DSSelectionFields
                    If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                        HasKeyForRel = True
                        Exit Try
                    End If
                Next
            Next

            MsgBox("At least one Key Field must be set in any Relational Target Datastore." & Chr(13) & _
            "No key fields have been set in this Datastore. Please set at least one key field.", _
            MsgBoxStyle.Exclamation, "Please Set a key field")

            HasKeyForRel = False

        Catch ex As Exception
            LogError(ex, "ctlDatastore HasKeyForRel")
            Return False
        End Try

    End Function

    Private Sub cmbColDelimiter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbColDelimiter.SelectedIndexChanged

        Try
            If IsEventFromCode = True Then Exit Sub

            objThis.ColumnDelimiter = CType(cmbColDelimiter.Items(cmbColDelimiter.SelectedIndex), Mylist).ItemData
            setRowCombo(cmbColDelimiter.SelectedItem, cmbTextQualifier.SelectedItem)
            SetListItemByValue(cmbRowDelimiter, objThis.RowDelimiter, False)
            setTextCombo(cmbRowDelimiter.SelectedItem, cmbColDelimiter.SelectedItem)
            SetListItemByValue(cmbTextQualifier, objThis.TextQualifier, False)

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbColDelimiter_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub cmbRowDelimiter_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRowDelimiter.SelectedIndexChanged

        Try
            If IsEventFromCode = True Then Exit Sub

            objThis.RowDelimiter = CType(cmbRowDelimiter.Items(cmbRowDelimiter.SelectedIndex), Mylist).ItemData
            setColCombo(cmbRowDelimiter.SelectedItem, cmbTextQualifier.SelectedItem)
            SetListItemByValue(cmbColDelimiter, objThis.ColumnDelimiter, False)
            setTextCombo(cmbRowDelimiter.SelectedItem, cmbColDelimiter.SelectedItem)
            SetListItemByValue(cmbTextQualifier, objThis.TextQualifier, False)

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbRowDelimiter_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub cmbTextQualifier_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTextQualifier.SelectedIndexChanged

        Try
            If IsEventFromCode = True Then Exit Sub

            objThis.TextQualifier = CType(cmbTextQualifier.Items(cmbTextQualifier.SelectedIndex), Mylist).ItemData
            setRowCombo(cmbColDelimiter.SelectedItem, cmbTextQualifier.SelectedItem)
            SetListItemByValue(cmbRowDelimiter, objThis.RowDelimiter, False)
            setColCombo(cmbRowDelimiter.SelectedItem, cmbTextQualifier.SelectedItem)
            SetListItemByValue(cmbColDelimiter, objThis.ColumnDelimiter, False)

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbTextQualifier_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub setColCombo(Optional ByVal RowVal As Mylist = Nothing, Optional ByVal TextVal As Mylist = Nothing)

        Try
            Dim rowtxt As String
            Dim texttxt As String

            If RowVal IsNot Nothing Then
                rowtxt = RowVal.Name
            Else
                rowtxt = ""
            End If
            If TextVal IsNot Nothing Then
                texttxt = TextVal.Name
            Else
                texttxt = ""
            End If

            cmbColDelimiter.BeginUpdate()

            cmbColDelimiter.Items.Clear()

            cmbColDelimiter.Items.Add(New Mylist("", DS_DELIMITER_NONE))

            If rowtxt <> "Comma" And texttxt <> "Comma" Then
                cmbColDelimiter.Items.Add(New Mylist("Comma", DS_DELIMITER_COMMA))
            End If
            If rowtxt <> "Tab" And texttxt <> "Tab" Then
                cmbColDelimiter.Items.Add(New Mylist("Tab", DS_DELIMITER_TAB))
            End If
            If rowtxt <> "Semicolon" And texttxt <> "Semicolon" Then
                cmbColDelimiter.Items.Add(New Mylist("Semicolon", DS_DELIMITER_SEMICOLON))
            End If
            If rowtxt <> "Verticalbar" And texttxt <> "Verticalbar" Then
                cmbColDelimiter.Items.Add(New Mylist("Verticalbar", DS_DELIMITER_VERTICALBAR))
            End If
            If rowtxt <> "{CR}{LF}" And texttxt <> "{CR}{LF}" Then
                cmbColDelimiter.Items.Add(New Mylist("{CR}{LF}", DS_DELIMITER_CRLF))
            End If
            If rowtxt <> "{CR}" And texttxt <> "{CR}" Then
                cmbColDelimiter.Items.Add(New Mylist("{CR}", DS_DELIMITER_CR))
            End If
            If rowtxt <> "{LF}" And texttxt <> "{LF}" Then
                cmbColDelimiter.Items.Add(New Mylist("{LF}", DS_DELIMITER_LF))
            End If
            If rowtxt <> "Tilde{~}" And texttxt <> "Tilde{~}" Then
                cmbColDelimiter.Items.Add(New Mylist("Tilde{~}", DS_DELIMITER_TILDE))
            End If

            cmbColDelimiter.EndUpdate()

        Catch ex As Exception
            LogError(ex, "ctlDatastore setColCombo")
        End Try

    End Sub

    Private Sub setRowCombo(Optional ByVal ColVal As Mylist = Nothing, Optional ByVal TextVal As Mylist = Nothing)

        Try
            Dim coltxt As String
            Dim texttxt As String

            If ColVal IsNot Nothing Then
                coltxt = ColVal.Name
            Else
                coltxt = ""
            End If
            If TextVal IsNot Nothing Then
                texttxt = TextVal.Name
            Else
                texttxt = ""
            End If

            cmbRowDelimiter.BeginUpdate()

            cmbRowDelimiter.Items.Clear()

            cmbRowDelimiter.Items.Add(New Mylist("", DS_DELIMITER_NONE))

            If coltxt <> "{CR}{LF}" And texttxt <> "{CR}{LF}" Then
                cmbRowDelimiter.Items.Add(New Mylist("{CR}{LF}", DS_DELIMITER_CRLF))
            End If
            If coltxt <> "{CR}" And texttxt <> "{CR}" Then
                cmbRowDelimiter.Items.Add(New Mylist("{CR}", DS_DELIMITER_CR))
            End If
            If coltxt <> "{LF}" And texttxt <> "{LF}" Then
                cmbRowDelimiter.Items.Add(New Mylist("{LF}", DS_DELIMITER_LF))
            End If
            If coltxt <> "Semicolon" And texttxt <> "Semicolon" Then
                cmbRowDelimiter.Items.Add(New Mylist("Semicolon", DS_DELIMITER_SEMICOLON))
            End If
            If coltxt <> "Verticalbar" And texttxt <> "Verticalbar" Then
                cmbRowDelimiter.Items.Add(New Mylist("Verticalbar", DS_DELIMITER_VERTICALBAR))
            End If

            cmbRowDelimiter.EndUpdate()

        Catch ex As Exception
            LogError(ex, "ctlDatastore setRowCombo")
        End Try

    End Sub

    Private Sub setTextCombo(Optional ByVal RowVal As Mylist = Nothing, Optional ByVal ColVal As Mylist = Nothing)

        Try
            Dim coltxt As String
            Dim rowtxt As String

            If ColVal IsNot Nothing Then
                coltxt = ColVal.Name
            Else
                coltxt = ""
            End If
            If RowVal IsNot Nothing Then
                rowtxt = RowVal.Name
            Else
                rowtxt = ""
            End If

            cmbTextQualifier.BeginUpdate()

            cmbTextQualifier.Items.Clear()

            cmbTextQualifier.Items.Add(New Mylist("", DS_DELIMITER_NONE))

            If coltxt <> "Double Quote" And rowtxt <> "Double Quote" Then
                cmbTextQualifier.Items.Add(New Mylist("Double Quote", DS_DELIMITER_DQUOTE))
            End If
            If coltxt <> "Single Quote" And rowtxt <> "Single Quote" Then
                cmbTextQualifier.Items.Add(New Mylist("Single Quote", DS_DELIMITER_SQUOTE))
            End If
            If coltxt <> "{CR}{LF}" And rowtxt <> "{CR}{LF}" Then
                cmbTextQualifier.Items.Add(New Mylist("{CR}{LF}", DS_DELIMITER_CRLF))
            End If
            If coltxt <> "{CR}" And rowtxt <> "{CR}" Then
                cmbTextQualifier.Items.Add(New Mylist("{CR}", DS_DELIMITER_CR))
            End If
            If coltxt <> "{LF}" And rowtxt <> "{LF}" Then
                cmbTextQualifier.Items.Add(New Mylist("{LF}", DS_DELIMITER_LF))
            End If
            If coltxt <> "Semicolon" And rowtxt <> "Semicolon" Then
                cmbTextQualifier.Items.Add(New Mylist("Semicolon", DS_DELIMITER_SEMICOLON))
            End If
            If coltxt <> "Verticalbar" And rowtxt <> "Verticalbar" Then
                cmbTextQualifier.Items.Add(New Mylist("Verticalbar", DS_DELIMITER_VERTICALBAR))
            End If

            cmbTextQualifier.EndUpdate()

        Catch ex As Exception
            LogError(ex, "ctlDatastore setTextCombo")
        End Try

    End Sub

End Class
