Public Class ctlDatastore
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    Dim WithEvents sl As MyListener

    Public objThis As New clsDatastore

    Dim IsNewObj As Boolean
    Dim prevfld As clsField
    Dim KeyForRel As Boolean = True
    Dim check4key As Boolean = False

    Private Const SplitSrc As Integer = 361
    Private Const SplitTgt As Integer = 284
    Private Const SplitSrcNoDel As Integer = 298
    Private Const SplitTgtNoDel As Integer = 221
    Private Pnt As Point
    Friend WithEvents lblHostName As System.Windows.Forms.Label
    Friend WithEvents txtHostName As System.Windows.Forms.TextBox
    Friend WithEvents gbMultiDesc As System.Windows.Forms.GroupBox
    Friend WithEvents scDescriptions As System.Windows.Forms.SplitContainer
    Friend WithEvents btnMTD As System.Windows.Forms.Button
    Friend WithEvents txtMTD As System.Windows.Forms.TextBox
    Friend WithEvents cbUseFile As System.Windows.Forms.CheckBox
    Friend WithEvents btnSelDesc As System.Windows.Forms.Button

    Private lvItem As ListViewItem

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
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents cmbDatastoreType As System.Windows.Forms.ComboBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents txtUOW As System.Windows.Forms.TextBox
    Friend WithEvents lblUOW As System.Windows.Forms.Label
    Friend WithEvents cmbOperationType As System.Windows.Forms.ComboBox
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents txtPortOrMQMgr As System.Windows.Forms.TextBox
    Friend WithEvents lblPortorMQMgr As System.Windows.Forms.Label
    Friend WithEvents txtException As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents chkSkipChangeCheck As System.Windows.Forms.CheckBox
    Friend WithEvents chkIMSPathData As System.Windows.Forms.CheckBox
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtPhysicalSource As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents tvDatastoreStructures As System.Windows.Forms.TreeView
    Friend WithEvents txtDatastoreName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtFieldDesc As System.Windows.Forms.TextBox
    Friend WithEvents A4 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents cmbListViewCombo As System.Windows.Forms.ComboBox
    Friend WithEvents txtListViewText As System.Windows.Forms.TextBox
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents gbProp As System.Windows.Forms.GroupBox
    Friend WithEvents gbStruct As System.Windows.Forms.GroupBox
    Friend WithEvents gbExtProps As System.Windows.Forms.GroupBox
    Friend WithEvents scDS As System.Windows.Forms.SplitContainer
    Friend WithEvents gbField As System.Windows.Forms.GroupBox
    Friend WithEvents gbDel As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer5 As System.Windows.Forms.SplitContainer
    Friend WithEvents gbFldname As System.Windows.Forms.GroupBox
    Friend WithEvents gbAtt As System.Windows.Forms.GroupBox
    Friend WithEvents SplitContainer6 As System.Windows.Forms.SplitContainer
    Friend WithEvents gbFldDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents cbInvalidChar As System.Windows.Forms.ComboBox
    Friend WithEvents cbIfSpaceChar As System.Windows.Forms.ComboBox
    Friend WithEvents lblInValid As System.Windows.Forms.Label
    Friend WithEvents lblExtType As System.Windows.Forms.Label
    Friend WithEvents cbIfSpaceNum As System.Windows.Forms.ComboBox
    Friend WithEvents cbIfNullChar As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents cbInvalidNum As System.Windows.Forms.ComboBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents cbIfNullNum As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents rbLU As System.Windows.Forms.RadioButton
    Friend WithEvents gbTarget As System.Windows.Forms.GroupBox
    Friend WithEvents cbKeyChng As System.Windows.Forms.CheckBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ctlDatastore))
        Me.cmbDatastoreType = New System.Windows.Forms.ComboBox
        Me.Label15 = New System.Windows.Forms.Label
        Me.txtUOW = New System.Windows.Forms.TextBox
        Me.cmbOperationType = New System.Windows.Forms.ComboBox
        Me.Label16 = New System.Windows.Forms.Label
        Me.txtPortOrMQMgr = New System.Windows.Forms.TextBox
        Me.lblPortorMQMgr = New System.Windows.Forms.Label
        Me.txtException = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.chkSkipChangeCheck = New System.Windows.Forms.CheckBox
        Me.chkIMSPathData = New System.Windows.Forms.CheckBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.cmbTextQualifier = New System.Windows.Forms.ComboBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.cmbRowDelimiter = New System.Windows.Forms.ComboBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbColDelimiter = New System.Windows.Forms.ComboBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtDatastoreDesc = New System.Windows.Forms.TextBox
        Me.cmbCharacterCode = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmbAccessMethod = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtPhysicalSource = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.tvDatastoreStructures = New System.Windows.Forms.TreeView
        Me.txtDatastoreName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblUOW = New System.Windows.Forms.Label
        Me.cmbListViewCombo = New System.Windows.Forms.ComboBox
        Me.txtListViewText = New System.Windows.Forms.TextBox
        Me.Label23 = New System.Windows.Forms.Label
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Label22 = New System.Windows.Forms.Label
        Me.Label21 = New System.Windows.Forms.Label
        Me.Label20 = New System.Windows.Forms.Label
        Me.Label18 = New System.Windows.Forms.Label
        Me.A4 = New System.Windows.Forms.Label
        Me.txtFieldDesc = New System.Windows.Forms.TextBox
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.gbProp = New System.Windows.Forms.GroupBox
        Me.lblHostName = New System.Windows.Forms.Label
        Me.txtHostName = New System.Windows.Forms.TextBox
        Me.rbLU = New System.Windows.Forms.RadioButton
        Me.gbExtProps = New System.Windows.Forms.GroupBox
        Me.gbSource = New System.Windows.Forms.GroupBox
        Me.cbIfNullNum = New System.Windows.Forms.ComboBox
        Me.Label19 = New System.Windows.Forms.Label
        Me.Label17 = New System.Windows.Forms.Label
        Me.cbInvalidNum = New System.Windows.Forms.ComboBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.cbIfSpaceNum = New System.Windows.Forms.ComboBox
        Me.cbIfNullChar = New System.Windows.Forms.ComboBox
        Me.lblInValid = New System.Windows.Forms.Label
        Me.lblExtType = New System.Windows.Forms.Label
        Me.cbInvalidChar = New System.Windows.Forms.ComboBox
        Me.cbIfSpaceChar = New System.Windows.Forms.ComboBox
        Me.gbStruct = New System.Windows.Forms.GroupBox
        Me.scDS = New System.Windows.Forms.SplitContainer
        Me.scDescriptions = New System.Windows.Forms.SplitContainer
        Me.gbMultiDesc = New System.Windows.Forms.GroupBox
        Me.btnSelDesc = New System.Windows.Forms.Button
        Me.cbUseFile = New System.Windows.Forms.CheckBox
        Me.btnMTD = New System.Windows.Forms.Button
        Me.txtMTD = New System.Windows.Forms.TextBox
        Me.gbTarget = New System.Windows.Forms.GroupBox
        Me.cbKeyChng = New System.Windows.Forms.CheckBox
        Me.gbDel = New System.Windows.Forms.GroupBox
        Me.gbField = New System.Windows.Forms.GroupBox
        Me.SplitContainer5 = New System.Windows.Forms.SplitContainer
        Me.gbFldname = New System.Windows.Forms.GroupBox
        Me.gbAtt = New System.Windows.Forms.GroupBox
        Me.SplitContainer6 = New System.Windows.Forms.SplitContainer
        Me.gbFldDesc = New System.Windows.Forms.GroupBox
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.gbProp.SuspendLayout()
        Me.gbExtProps.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.gbStruct.SuspendLayout()
        Me.scDS.Panel1.SuspendLayout()
        Me.scDS.Panel2.SuspendLayout()
        Me.scDS.SuspendLayout()
        Me.scDescriptions.Panel1.SuspendLayout()
        Me.scDescriptions.Panel2.SuspendLayout()
        Me.scDescriptions.SuspendLayout()
        Me.gbMultiDesc.SuspendLayout()
        Me.gbTarget.SuspendLayout()
        Me.gbDel.SuspendLayout()
        Me.gbField.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.Panel2.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        Me.gbFldname.SuspendLayout()
        Me.gbAtt.SuspendLayout()
        Me.SplitContainer6.Panel1.SuspendLayout()
        Me.SplitContainer6.Panel2.SuspendLayout()
        Me.SplitContainer6.SuspendLayout()
        Me.gbFldDesc.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbDatastoreType
        '
        Me.cmbDatastoreType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDatastoreType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbDatastoreType.Location = New System.Drawing.Point(301, 14)
        Me.cmbDatastoreType.Name = "cmbDatastoreType"
        Me.cmbDatastoreType.Size = New System.Drawing.Size(117, 21)
        Me.cmbDatastoreType.TabIndex = 5
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(251, 18)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(34, 14)
        Me.Label15.TabIndex = 0
        Me.Label15.Text = "Type"
        '
        'txtUOW
        '
        Me.txtUOW.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUOW.Location = New System.Drawing.Point(284, 13)
        Me.txtUOW.MaxLength = 128
        Me.txtUOW.Name = "txtUOW"
        Me.txtUOW.Size = New System.Drawing.Size(134, 20)
        Me.txtUOW.TabIndex = 16
        '
        'cmbOperationType
        '
        Me.cmbOperationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOperationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbOperationType.Location = New System.Drawing.Point(103, 14)
        Me.cmbOperationType.Name = "cmbOperationType"
        Me.cmbOperationType.Size = New System.Drawing.Size(117, 21)
        Me.cmbOperationType.TabIndex = 6
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(6, 18)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(91, 14)
        Me.Label16.TabIndex = 0
        Me.Label16.Text = "Operation Type"
        '
        'txtPortOrMQMgr
        '
        Me.txtPortOrMQMgr.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPortOrMQMgr.Location = New System.Drawing.Point(301, 67)
        Me.txtPortOrMQMgr.MaxLength = 128
        Me.txtPortOrMQMgr.Name = "txtPortOrMQMgr"
        Me.txtPortOrMQMgr.Size = New System.Drawing.Size(117, 20)
        Me.txtPortOrMQMgr.TabIndex = 4
        '
        'lblPortorMQMgr
        '
        Me.lblPortorMQMgr.AutoSize = True
        Me.lblPortorMQMgr.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPortorMQMgr.Location = New System.Drawing.Point(251, 70)
        Me.lblPortorMQMgr.Name = "lblPortorMQMgr"
        Me.lblPortorMQMgr.Size = New System.Drawing.Size(30, 14)
        Me.lblPortorMQMgr.TabIndex = 0
        Me.lblPortorMQMgr.Text = "Port"
        '
        'txtException
        '
        Me.txtException.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtException.Location = New System.Drawing.Point(72, 37)
        Me.txtException.MaxLength = 128
        Me.txtException.Name = "txtException"
        Me.txtException.Size = New System.Drawing.Size(173, 20)
        Me.txtException.TabIndex = 3
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 40)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(60, 14)
        Me.Label7.TabIndex = 0
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
        Me.chkSkipChangeCheck.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkIMSPathData
        '
        Me.chkIMSPathData.AutoSize = True
        Me.chkIMSPathData.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkIMSPathData.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIMSPathData.Location = New System.Drawing.Point(6, 15)
        Me.chkIMSPathData.Name = "chkIMSPathData"
        Me.chkIMSPathData.Size = New System.Drawing.Size(93, 18)
        Me.chkIMSPathData.TabIndex = 17
        Me.chkIMSPathData.Text = "IMSPathData"
        Me.chkIMSPathData.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.Label12.Location = New System.Drawing.Point(6, 37)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(465, 14)
        Me.Label12.TabIndex = 0
        Me.Label12.Text = "Note : If you enter hex value then it must start with 0x (e.g. hex value for semi" & _
            "colon (;) is 0x3B)"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(436, 16)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(80, 14)
        Me.Label11.TabIndex = 0
        Me.Label11.Text = "Text Qualifier"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbTextQualifier
        '
        Me.cmbTextQualifier.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbTextQualifier.Location = New System.Drawing.Point(522, 13)
        Me.cmbTextQualifier.Name = "cmbTextQualifier"
        Me.cmbTextQualifier.Size = New System.Drawing.Size(111, 21)
        Me.cmbTextQualifier.TabIndex = 20
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 16)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(103, 14)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Column Delimiter"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbRowDelimiter
        '
        Me.cmbRowDelimiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbRowDelimiter.Location = New System.Drawing.Point(334, 13)
        Me.cmbRowDelimiter.Name = "cmbRowDelimiter"
        Me.cmbRowDelimiter.Size = New System.Drawing.Size(89, 21)
        Me.cmbRowDelimiter.TabIndex = 19
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(229, 16)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(99, 14)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Record Delimiter"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbColDelimiter
        '
        Me.cmbColDelimiter.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbColDelimiter.Location = New System.Drawing.Point(115, 13)
        Me.cmbColDelimiter.Name = "cmbColDelimiter"
        Me.cmbColDelimiter.Size = New System.Drawing.Size(92, 21)
        Me.cmbColDelimiter.TabIndex = 18
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 172)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(70, 28)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Datastore Description"
        '
        'txtDatastoreDesc
        '
        Me.txtDatastoreDesc.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDatastoreDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatastoreDesc.Location = New System.Drawing.Point(82, 169)
        Me.txtDatastoreDesc.MaxLength = 255
        Me.txtDatastoreDesc.Multiline = True
        Me.txtDatastoreDesc.Name = "txtDatastoreDesc"
        Me.txtDatastoreDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDatastoreDesc.Size = New System.Drawing.Size(307, 36)
        Me.txtDatastoreDesc.TabIndex = 9
        '
        'cmbCharacterCode
        '
        Me.cmbCharacterCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCharacterCode.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbCharacterCode.Location = New System.Drawing.Point(115, 13)
        Me.cmbCharacterCode.Name = "cmbCharacterCode"
        Me.cmbCharacterCode.Size = New System.Drawing.Size(163, 21)
        Me.cmbCharacterCode.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 18)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(93, 14)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Character Code"
        '
        'cmbAccessMethod
        '
        Me.cmbAccessMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAccessMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbAccessMethod.Location = New System.Drawing.Point(301, 41)
        Me.cmbAccessMethod.Name = "cmbAccessMethod"
        Me.cmbAccessMethod.Size = New System.Drawing.Size(117, 21)
        Me.cmbAccessMethod.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(251, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(48, 14)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Access"
        '
        'txtPhysicalSource
        '
        Me.txtPhysicalSource.AcceptsTab = True
        Me.txtPhysicalSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhysicalSource.Location = New System.Drawing.Point(60, 15)
        Me.txtPhysicalSource.MaxLength = 128
        Me.txtPhysicalSource.Name = "txtPhysicalSource"
        Me.txtPhysicalSource.Size = New System.Drawing.Size(185, 20)
        Me.txtPhysicalSource.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 13)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(58, 28)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Physical Name"
        '
        'tvDatastoreStructures
        '
        Me.tvDatastoreStructures.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvDatastoreStructures.CheckBoxes = True
        Me.tvDatastoreStructures.HideSelection = False
        Me.tvDatastoreStructures.Location = New System.Drawing.Point(6, 16)
        Me.tvDatastoreStructures.Name = "tvDatastoreStructures"
        Me.tvDatastoreStructures.Size = New System.Drawing.Size(383, 147)
        Me.tvDatastoreStructures.TabIndex = 10
        '
        'txtDatastoreName
        '
        Me.txtDatastoreName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDatastoreName.Location = New System.Drawing.Point(60, 41)
        Me.txtDatastoreName.MaxLength = 20
        Me.txtDatastoreName.Name = "txtDatastoreName"
        Me.txtDatastoreName.Size = New System.Drawing.Size(185, 20)
        Me.txtDatastoreName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(5, 45)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 14)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Alias"
        '
        'lblUOW
        '
        Me.lblUOW.AutoSize = True
        Me.lblUOW.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUOW.Location = New System.Drawing.Point(250, 16)
        Me.lblUOW.Name = "lblUOW"
        Me.lblUOW.Size = New System.Drawing.Size(35, 14)
        Me.lblUOW.TabIndex = 0
        Me.lblUOW.Text = "UOW "
        '
        'cmbListViewCombo
        '
        Me.cmbListViewCombo.BackColor = System.Drawing.SystemColors.Info
        Me.cmbListViewCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbListViewCombo.Font = New System.Drawing.Font("Tahoma", 7.0!)
        Me.cmbListViewCombo.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.cmbListViewCombo.ItemHeight = 11
        Me.cmbListViewCombo.Location = New System.Drawing.Point(46, 94)
        Me.cmbListViewCombo.Name = "cmbListViewCombo"
        Me.cmbListViewCombo.Size = New System.Drawing.Size(145, 19)
        Me.cmbListViewCombo.TabIndex = 18
        Me.cmbListViewCombo.Visible = False
        '
        'txtListViewText
        '
        Me.txtListViewText.BackColor = System.Drawing.SystemColors.Info
        Me.txtListViewText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtListViewText.Font = New System.Drawing.Font("Tahoma", 7.0!)
        Me.txtListViewText.Location = New System.Drawing.Point(59, 59)
        Me.txtListViewText.MaxLength = 128
        Me.txtListViewText.Name = "txtListViewText"
        Me.txtListViewText.Size = New System.Drawing.Size(96, 19)
        Me.txtListViewText.TabIndex = 17
        Me.txtListViewText.Visible = False
        Me.txtListViewText.WordWrap = False
        '
        'Label23
        '
        Me.Label23.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label23.AutoSize = True
        Me.Label23.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label23.Location = New System.Drawing.Point(312, 648)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(107, 13)
        Me.Label23.TabIndex = 0
        Me.Label23.Text = "Mapped in a Proc"
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
        Me.lvFieldAttrs.Location = New System.Drawing.Point(0, 0)
        Me.lvFieldAttrs.Name = "lvFieldAttrs"
        Me.lvFieldAttrs.Size = New System.Drawing.Size(569, 161)
        Me.lvFieldAttrs.TabIndex = 23
        Me.lvFieldAttrs.UseCompatibleStateImageBehavior = False
        Me.lvFieldAttrs.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Attribute"
        Me.ColumnHeader1.Width = 210
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Value"
        Me.ColumnHeader2.Width = 204
        '
        'Label22
        '
        Me.Label22.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label22.BackColor = System.Drawing.Color.White
        Me.Label22.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label22.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label22.ForeColor = System.Drawing.Color.Green
        Me.Label22.Location = New System.Drawing.Point(290, 647)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(16, 16)
        Me.Label22.TabIndex = 0
        Me.Label22.Text = "A"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label21.AutoSize = True
        Me.Label21.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label21.Location = New System.Drawing.Point(184, 648)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(100, 13)
        Me.Label21.TabIndex = 0
        Me.Label21.Text = "Has Foreign Key"
        '
        'Label20
        '
        Me.Label20.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label20.BackColor = System.Drawing.Color.White
        Me.Label20.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label20.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.ForeColor = System.Drawing.Color.Red
        Me.Label20.Location = New System.Drawing.Point(162, 647)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(16, 16)
        Me.Label20.TabIndex = 0
        Me.Label20.Text = "A"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label18
        '
        Me.Label18.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label18.AutoSize = True
        Me.Label18.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label18.Location = New System.Drawing.Point(28, 648)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(128, 13)
        Me.Label18.TabIndex = 0
        Me.Label18.Text = "Has Field Description"
        '
        'A4
        '
        Me.A4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.A4.BackColor = System.Drawing.Color.White
        Me.A4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.A4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.A4.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.A4.Location = New System.Drawing.Point(6, 647)
        Me.A4.Name = "A4"
        Me.A4.Size = New System.Drawing.Size(16, 16)
        Me.A4.TabIndex = 0
        Me.A4.Text = "A"
        Me.A4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtFieldDesc
        '
        Me.txtFieldDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtFieldDesc.CausesValidation = False
        Me.txtFieldDesc.Cursor = System.Windows.Forms.Cursors.Hand
        Me.txtFieldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtFieldDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFieldDesc.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.txtFieldDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtFieldDesc.Multiline = True
        Me.txtFieldDesc.Name = "txtFieldDesc"
        Me.txtFieldDesc.Size = New System.Drawing.Size(563, 37)
        Me.txtFieldDesc.TabIndex = 58
        Me.txtFieldDesc.TabStop = False
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(766, 642)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 26
        Me.cmdHelp.Text = "&Help"
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(610, 642)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 24
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(688, 642)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 25
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(0, 629)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(838, 7)
        Me.GroupBox1.TabIndex = 43
        Me.GroupBox1.TabStop = False
        '
        'tvFields
        '
        Me.tvFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tvFields.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tvFields.HideSelection = False
        Me.tvFields.HotTracking = True
        Me.tvFields.ImageIndex = 1
        Me.tvFields.ImageList = Me.ImageList1
        Me.tvFields.Location = New System.Drawing.Point(3, 19)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.SelectedImageIndex = 0
        Me.tvFields.Size = New System.Drawing.Size(242, 218)
        Me.tvFields.StateImageList = Me.ImageList1
        Me.tvFields.TabIndex = 22
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "105.ico")
        Me.ImageList1.Images.SetKeyName(1, "Field List.ico")
        '
        'gbProp
        '
        Me.gbProp.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gbProp.Controls.Add(Me.lblHostName)
        Me.gbProp.Controls.Add(Me.txtHostName)
        Me.gbProp.Controls.Add(Me.Label3)
        Me.gbProp.Controls.Add(Me.txtDatastoreName)
        Me.gbProp.Controls.Add(Me.Label15)
        Me.gbProp.Controls.Add(Me.cmbDatastoreType)
        Me.gbProp.Controls.Add(Me.txtPortOrMQMgr)
        Me.gbProp.Controls.Add(Me.Label2)
        Me.gbProp.Controls.Add(Me.lblPortorMQMgr)
        Me.gbProp.Controls.Add(Me.txtPhysicalSource)
        Me.gbProp.Controls.Add(Me.Label1)
        Me.gbProp.Controls.Add(Me.cmbAccessMethod)
        Me.gbProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProp.ForeColor = System.Drawing.Color.White
        Me.gbProp.Location = New System.Drawing.Point(3, 3)
        Me.gbProp.Name = "gbProp"
        Me.gbProp.Size = New System.Drawing.Size(424, 94)
        Me.gbProp.TabIndex = 0
        Me.gbProp.TabStop = False
        Me.gbProp.Text = "Datastore Properties"
        '
        'lblHostName
        '
        Me.lblHostName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHostName.Location = New System.Drawing.Point(4, 63)
        Me.lblHostName.Name = "lblHostName"
        Me.lblHostName.Size = New System.Drawing.Size(48, 30)
        Me.lblHostName.TabIndex = 10
        Me.lblHostName.Text = "Host Name"
        '
        'txtHostName
        '
        Me.txtHostName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHostName.Location = New System.Drawing.Point(60, 67)
        Me.txtHostName.Name = "txtHostName"
        Me.txtHostName.Size = New System.Drawing.Size(185, 20)
        Me.txtHostName.TabIndex = 9
        '
        'rbLU
        '
        Me.rbLU.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rbLU.ForeColor = System.Drawing.Color.White
        Me.rbLU.Location = New System.Drawing.Point(332, 172)
        Me.rbLU.Name = "rbLU"
        Me.rbLU.Size = New System.Drawing.Size(89, 32)
        Me.rbLU.TabIndex = 10
        Me.rbLU.TabStop = True
        Me.rbLU.Text = "Lookup Datastore"
        Me.rbLU.UseVisualStyleBackColor = True
        '
        'gbExtProps
        '
        Me.gbExtProps.Controls.Add(Me.txtUOW)
        Me.gbExtProps.Controls.Add(Me.chkSkipChangeCheck)
        Me.gbExtProps.Controls.Add(Me.lblUOW)
        Me.gbExtProps.Controls.Add(Me.chkIMSPathData)
        Me.gbExtProps.Controls.Add(Me.txtException)
        Me.gbExtProps.Controls.Add(Me.Label7)
        Me.gbExtProps.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbExtProps.ForeColor = System.Drawing.Color.White
        Me.gbExtProps.Location = New System.Drawing.Point(3, 103)
        Me.gbExtProps.Name = "gbExtProps"
        Me.gbExtProps.Size = New System.Drawing.Size(424, 63)
        Me.gbExtProps.TabIndex = 0
        Me.gbExtProps.TabStop = False
        Me.gbExtProps.Text = "Extended Properties"
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.cbIfNullNum)
        Me.gbSource.Controls.Add(Me.Label19)
        Me.gbSource.Controls.Add(Me.Label17)
        Me.gbSource.Controls.Add(Me.cbInvalidNum)
        Me.gbSource.Controls.Add(Me.Label13)
        Me.gbSource.Controls.Add(Me.Label4)
        Me.gbSource.Controls.Add(Me.Label6)
        Me.gbSource.Controls.Add(Me.cmbCharacterCode)
        Me.gbSource.Controls.Add(Me.cbIfSpaceNum)
        Me.gbSource.Controls.Add(Me.cbIfNullChar)
        Me.gbSource.Controls.Add(Me.lblInValid)
        Me.gbSource.Controls.Add(Me.lblExtType)
        Me.gbSource.Controls.Add(Me.cbInvalidChar)
        Me.gbSource.Controls.Add(Me.cbIfSpaceChar)
        Me.gbSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSource.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.gbSource.Location = New System.Drawing.Point(3, 172)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(423, 122)
        Me.gbSource.TabIndex = 20
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Source Datastore Field Properties"
        Me.ToolTip1.SetToolTip(Me.gbSource, "Both Drop-Down boxes must be set for each Field Property Set")
        '
        'cbIfNullNum
        '
        Me.cbIfNullNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbIfNullNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfNullNum.FormattingEnabled = True
        Me.cbIfNullNum.Location = New System.Drawing.Point(316, 40)
        Me.cbIfNullNum.Name = "cbIfNullNum"
        Me.cbIfNullNum.Size = New System.Drawing.Size(102, 21)
        Me.cbIfNullNum.TabIndex = 33
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Location = New System.Drawing.Point(241, 43)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(69, 13)
        Me.Label19.TabIndex = 32
        Me.Label19.Text = "for All Num"
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 43)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(108, 13)
        Me.Label17.TabIndex = 31
        Me.Label17.Text = "If Null for All Char"
        '
        'cbInvalidNum
        '
        Me.cbInvalidNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbInvalidNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbInvalidNum.FormattingEnabled = True
        Me.cbInvalidNum.Location = New System.Drawing.Point(316, 67)
        Me.cbInvalidNum.Name = "cbInvalidNum"
        Me.cbInvalidNum.Size = New System.Drawing.Size(102, 21)
        Me.cbInvalidNum.TabIndex = 30
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(241, 96)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(69, 13)
        Me.Label13.TabIndex = 29
        Me.Label13.Text = "for All Num"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(241, 70)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 13)
        Me.Label4.TabIndex = 28
        Me.Label4.Text = "for All Num"
        '
        'cbIfSpaceNum
        '
        Me.cbIfSpaceNum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbIfSpaceNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfSpaceNum.FormattingEnabled = True
        Me.cbIfSpaceNum.Location = New System.Drawing.Point(316, 93)
        Me.cbIfSpaceNum.Name = "cbIfSpaceNum"
        Me.cbIfSpaceNum.Size = New System.Drawing.Size(102, 21)
        Me.cbIfSpaceNum.TabIndex = 27
        '
        'cbIfNullChar
        '
        Me.cbIfNullChar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbIfNullChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfNullChar.FormattingEnabled = True
        Me.cbIfNullChar.Location = New System.Drawing.Point(134, 40)
        Me.cbIfNullChar.Name = "cbIfNullChar"
        Me.cbIfNullChar.Size = New System.Drawing.Size(101, 21)
        Me.cbIfNullChar.TabIndex = 26
        '
        'lblInValid
        '
        Me.lblInValid.AutoSize = True
        Me.lblInValid.Location = New System.Drawing.Point(6, 70)
        Me.lblInValid.Name = "lblInValid"
        Me.lblInValid.Size = New System.Drawing.Size(112, 13)
        Me.lblInValid.TabIndex = 25
        Me.lblInValid.Text = "Invalid for All Char"
        '
        'lblExtType
        '
        Me.lblExtType.AutoSize = True
        Me.lblExtType.Location = New System.Drawing.Point(6, 97)
        Me.lblExtType.Name = "lblExtType"
        Me.lblExtType.Size = New System.Drawing.Size(122, 13)
        Me.lblExtType.TabIndex = 22
        Me.lblExtType.Text = "If Space for All Char"
        '
        'cbInvalidChar
        '
        Me.cbInvalidChar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbInvalidChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbInvalidChar.FormattingEnabled = True
        Me.cbInvalidChar.Location = New System.Drawing.Point(134, 67)
        Me.cbInvalidChar.MaxDropDownItems = 20
        Me.cbInvalidChar.Name = "cbInvalidChar"
        Me.cbInvalidChar.Size = New System.Drawing.Size(101, 21)
        Me.cbInvalidChar.TabIndex = 21
        '
        'cbIfSpaceChar
        '
        Me.cbIfSpaceChar.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbIfSpaceChar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbIfSpaceChar.FormattingEnabled = True
        Me.cbIfSpaceChar.Location = New System.Drawing.Point(134, 92)
        Me.cbIfSpaceChar.MaxDropDownItems = 20
        Me.cbIfSpaceChar.Name = "cbIfSpaceChar"
        Me.cbIfSpaceChar.Size = New System.Drawing.Size(101, 21)
        Me.cbIfSpaceChar.TabIndex = 18
        '
        'gbStruct
        '
        Me.gbStruct.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.gbStruct.Controls.Add(Me.tvDatastoreStructures)
        Me.gbStruct.Controls.Add(Me.txtDatastoreDesc)
        Me.gbStruct.Controls.Add(Me.Label8)
        Me.gbStruct.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbStruct.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStruct.ForeColor = System.Drawing.Color.White
        Me.gbStruct.Location = New System.Drawing.Point(0, 0)
        Me.gbStruct.Name = "gbStruct"
        Me.gbStruct.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbStruct.Size = New System.Drawing.Size(395, 211)
        Me.gbStruct.TabIndex = 0
        Me.gbStruct.TabStop = False
        Me.gbStruct.Text = "Descriptions"
        '
        'scDS
        '
        Me.scDS.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scDS.Location = New System.Drawing.Point(3, 3)
        Me.scDS.Name = "scDS"
        Me.scDS.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scDS.Panel1
        '
        Me.scDS.Panel1.Controls.Add(Me.scDescriptions)
        Me.scDS.Panel1.Controls.Add(Me.gbTarget)
        Me.scDS.Panel1.Controls.Add(Me.gbSource)
        Me.scDS.Panel1.Controls.Add(Me.gbDel)
        Me.scDS.Panel1.Controls.Add(Me.gbExtProps)
        Me.scDS.Panel1.Controls.Add(Me.gbProp)
        Me.scDS.Panel1.Controls.Add(Me.rbLU)
        '
        'scDS.Panel2
        '
        Me.scDS.Panel2.Controls.Add(Me.gbField)
        Me.scDS.Size = New System.Drawing.Size(833, 626)
        Me.scDS.SplitterDistance = 363
        Me.scDS.TabIndex = 0
        Me.scDS.TabStop = False
        '
        'scDescriptions
        '
        Me.scDescriptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.scDescriptions.FixedPanel = System.Windows.Forms.FixedPanel.Panel2
        Me.scDescriptions.IsSplitterFixed = True
        Me.scDescriptions.Location = New System.Drawing.Point(432, 3)
        Me.scDescriptions.Name = "scDescriptions"
        Me.scDescriptions.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'scDescriptions.Panel1
        '
        Me.scDescriptions.Panel1.Controls.Add(Me.gbStruct)
        '
        'scDescriptions.Panel2
        '
        Me.scDescriptions.Panel2.Controls.Add(Me.gbMultiDesc)
        Me.scDescriptions.Size = New System.Drawing.Size(395, 262)
        Me.scDescriptions.SplitterDistance = 211
        Me.scDescriptions.TabIndex = 23
        '
        'gbMultiDesc
        '
        Me.gbMultiDesc.Controls.Add(Me.btnSelDesc)
        Me.gbMultiDesc.Controls.Add(Me.cbUseFile)
        Me.gbMultiDesc.Controls.Add(Me.btnMTD)
        Me.gbMultiDesc.Controls.Add(Me.txtMTD)
        Me.gbMultiDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbMultiDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbMultiDesc.ForeColor = System.Drawing.Color.White
        Me.gbMultiDesc.Location = New System.Drawing.Point(0, 0)
        Me.gbMultiDesc.Name = "gbMultiDesc"
        Me.gbMultiDesc.Size = New System.Drawing.Size(395, 47)
        Me.gbMultiDesc.TabIndex = 22
        Me.gbMultiDesc.TabStop = False
        Me.gbMultiDesc.Text = "Multi-Table Description File"
        '
        'btnSelDesc
        '
        Me.btnSelDesc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSelDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSelDesc.ForeColor = System.Drawing.Color.Black
        Me.btnSelDesc.Location = New System.Drawing.Point(331, 19)
        Me.btnSelDesc.Name = "btnSelDesc"
        Me.btnSelDesc.Size = New System.Drawing.Size(58, 23)
        Me.btnSelDesc.TabIndex = 3
        Me.btnSelDesc.Text = "Select"
        Me.ToolTip1.SetToolTip(Me.btnSelDesc, "Select Descriptions" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Contained in file")
        Me.btnSelDesc.UseVisualStyleBackColor = True
        '
        'cbUseFile
        '
        Me.cbUseFile.AutoSize = True
        Me.cbUseFile.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbUseFile.Location = New System.Drawing.Point(4, 23)
        Me.cbUseFile.Name = "cbUseFile"
        Me.cbUseFile.Size = New System.Drawing.Size(72, 17)
        Me.cbUseFile.TabIndex = 2
        Me.cbUseFile.Text = "Use File"
        Me.cbUseFile.UseVisualStyleBackColor = True
        '
        'btnMTD
        '
        Me.btnMTD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnMTD.ForeColor = System.Drawing.Color.Black
        Me.btnMTD.Location = New System.Drawing.Point(297, 19)
        Me.btnMTD.Name = "btnMTD"
        Me.btnMTD.Size = New System.Drawing.Size(28, 23)
        Me.btnMTD.TabIndex = 1
        Me.btnMTD.Text = "..."
        Me.btnMTD.UseVisualStyleBackColor = True
        '
        'txtMTD
        '
        Me.txtMTD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMTD.BackColor = System.Drawing.SystemColors.Window
        Me.txtMTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMTD.Location = New System.Drawing.Point(82, 21)
        Me.txtMTD.Name = "txtMTD"
        Me.txtMTD.ReadOnly = True
        Me.txtMTD.Size = New System.Drawing.Size(210, 20)
        Me.txtMTD.TabIndex = 0
        '
        'gbTarget
        '
        Me.gbTarget.Controls.Add(Me.cbKeyChng)
        Me.gbTarget.Controls.Add(Me.cmbOperationType)
        Me.gbTarget.Controls.Add(Me.Label16)
        Me.gbTarget.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbTarget.ForeColor = System.Drawing.Color.White
        Me.gbTarget.Location = New System.Drawing.Point(433, 250)
        Me.gbTarget.Name = "gbTarget"
        Me.gbTarget.Size = New System.Drawing.Size(394, 44)
        Me.gbTarget.TabIndex = 21
        Me.gbTarget.TabStop = False
        Me.gbTarget.Text = "Target Properties"
        '
        'cbKeyChng
        '
        Me.cbKeyChng.AutoSize = True
        Me.cbKeyChng.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cbKeyChng.Location = New System.Drawing.Point(237, 18)
        Me.cbKeyChng.Name = "cbKeyChng"
        Me.cbKeyChng.Size = New System.Drawing.Size(151, 17)
        Me.cbKeyChng.TabIndex = 7
        Me.cbKeyChng.Text = "Allow for key changes"
        Me.cbKeyChng.UseVisualStyleBackColor = True
        '
        'gbDel
        '
        Me.gbDel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDel.Controls.Add(Me.cmbTextQualifier)
        Me.gbDel.Controls.Add(Me.Label11)
        Me.gbDel.Controls.Add(Me.Label12)
        Me.gbDel.Controls.Add(Me.cmbRowDelimiter)
        Me.gbDel.Controls.Add(Me.Label10)
        Me.gbDel.Controls.Add(Me.Label9)
        Me.gbDel.Controls.Add(Me.cmbColDelimiter)
        Me.gbDel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDel.ForeColor = System.Drawing.Color.White
        Me.gbDel.Location = New System.Drawing.Point(3, 300)
        Me.gbDel.Name = "gbDel"
        Me.gbDel.Size = New System.Drawing.Size(824, 57)
        Me.gbDel.TabIndex = 0
        Me.gbDel.TabStop = False
        Me.gbDel.Text = "File Properties"
        '
        'gbField
        '
        Me.gbField.Controls.Add(Me.SplitContainer5)
        Me.gbField.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbField.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbField.ForeColor = System.Drawing.Color.White
        Me.gbField.Location = New System.Drawing.Point(0, 0)
        Me.gbField.Name = "gbField"
        Me.gbField.Size = New System.Drawing.Size(833, 259)
        Me.gbField.TabIndex = 0
        Me.gbField.TabStop = False
        Me.gbField.Text = "Field Properties"
        '
        'SplitContainer5
        '
        Me.SplitContainer5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer5.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer5.Name = "SplitContainer5"
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.gbFldname)
        '
        'SplitContainer5.Panel2
        '
        Me.SplitContainer5.Panel2.Controls.Add(Me.gbAtt)
        Me.SplitContainer5.Size = New System.Drawing.Size(827, 240)
        Me.SplitContainer5.SplitterDistance = 248
        Me.SplitContainer5.TabIndex = 0
        Me.SplitContainer5.TabStop = False
        '
        'gbFldname
        '
        Me.gbFldname.Controls.Add(Me.tvFields)
        Me.gbFldname.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFldname.ForeColor = System.Drawing.Color.White
        Me.gbFldname.Location = New System.Drawing.Point(0, 0)
        Me.gbFldname.Name = "gbFldname"
        Me.gbFldname.Size = New System.Drawing.Size(248, 240)
        Me.gbFldname.TabIndex = 0
        Me.gbFldname.TabStop = False
        Me.gbFldname.Text = "Field Names"
        '
        'gbAtt
        '
        Me.gbAtt.Controls.Add(Me.SplitContainer6)
        Me.gbAtt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAtt.ForeColor = System.Drawing.Color.White
        Me.gbAtt.Location = New System.Drawing.Point(0, 0)
        Me.gbAtt.Name = "gbAtt"
        Me.gbAtt.Size = New System.Drawing.Size(575, 240)
        Me.gbAtt.TabIndex = 0
        Me.gbAtt.TabStop = False
        Me.gbAtt.Text = "Field Attributes"
        '
        'SplitContainer6
        '
        Me.SplitContainer6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer6.Location = New System.Drawing.Point(3, 16)
        Me.SplitContainer6.Name = "SplitContainer6"
        Me.SplitContainer6.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer6.Panel1
        '
        Me.SplitContainer6.Panel1.Controls.Add(Me.cmbListViewCombo)
        Me.SplitContainer6.Panel1.Controls.Add(Me.txtListViewText)
        Me.SplitContainer6.Panel1.Controls.Add(Me.lvFieldAttrs)
        '
        'SplitContainer6.Panel2
        '
        Me.SplitContainer6.Panel2.Controls.Add(Me.gbFldDesc)
        Me.SplitContainer6.Size = New System.Drawing.Size(569, 221)
        Me.SplitContainer6.SplitterDistance = 161
        Me.SplitContainer6.TabIndex = 0
        '
        'gbFldDesc
        '
        Me.gbFldDesc.Controls.Add(Me.txtFieldDesc)
        Me.gbFldDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFldDesc.ForeColor = System.Drawing.Color.White
        Me.gbFldDesc.Location = New System.Drawing.Point(0, 0)
        Me.gbFldDesc.Name = "gbFldDesc"
        Me.gbFldDesc.Size = New System.Drawing.Size(569, 56)
        Me.gbFldDesc.TabIndex = 0
        Me.gbFldDesc.TabStop = False
        Me.gbFldDesc.Text = "Field Description"
        '
        'ctlDatastore
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.scDS)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.A4)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlDatastore"
        Me.Size = New System.Drawing.Size(841, 673)
        Me.gbProp.ResumeLayout(False)
        Me.gbProp.PerformLayout()
        Me.gbExtProps.ResumeLayout(False)
        Me.gbExtProps.PerformLayout()
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.gbStruct.ResumeLayout(False)
        Me.gbStruct.PerformLayout()
        Me.scDS.Panel1.ResumeLayout(False)
        Me.scDS.Panel2.ResumeLayout(False)
        Me.scDS.ResumeLayout(False)
        Me.scDescriptions.Panel1.ResumeLayout(False)
        Me.scDescriptions.Panel2.ResumeLayout(False)
        Me.scDescriptions.ResumeLayout(False)
        Me.gbMultiDesc.ResumeLayout(False)
        Me.gbMultiDesc.PerformLayout()
        Me.gbTarget.ResumeLayout(False)
        Me.gbTarget.PerformLayout()
        Me.gbDel.ResumeLayout(False)
        Me.gbDel.PerformLayout()
        Me.gbField.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.Panel2.ResumeLayout(False)
        Me.SplitContainer5.ResumeLayout(False)
        Me.gbFldname.ResumeLayout(False)
        Me.gbAtt.ResumeLayout(False)
        Me.SplitContainer6.Panel1.ResumeLayout(False)
        Me.SplitContainer6.Panel1.PerformLayout()
        Me.SplitContainer6.Panel2.ResumeLayout(False)
        Me.SplitContainer6.ResumeLayout(False)
        Me.gbFldDesc.ResumeLayout(False)
        Me.gbFldDesc.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

#Region "Form events"

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtDatastoreName.Enabled = True '//added on 7/24 to disable object name editing

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsDatastore

        ClearControls(Me.Controls)


    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkIMSPathData.CheckedChanged, chkSkipChangeCheck.CheckedChanged, cmbAccessMethod.SelectedIndexChanged, cmbCharacterCode.SelectedIndexChanged, cmbDatastoreType.SelectedIndexChanged, txtDatastoreDesc.TextChanged, txtException.TextChanged, txtPhysicalSource.TextChanged, txtPortOrMQMgr.TextChanged, txtUOW.TextChanged, txtDatastoreName.TextChanged, cmbListViewCombo.SelectedIndexChanged, txtListViewText.TextChanged, cbKeyChng.CheckedChanged, cmbOperationType.SelectedIndexChanged 'cmbColDelimiter.SelectedIndexChanged, cmbRowDelimiter.SelectedIndexChanged, cmbTextQualifier.SelectedIndexChanged '// modified by TK and KS 11/6/2006

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub frmDatastore_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize

        'pnlFields.SuspendLayout()
        'pnlFields.Location = pnlSelect.Location
        'pnlFields.Size = pnlSelect.Size
        'pnlFields.ResumeLayout()
        If objThis.DsDirection = DS_DIRECTION_SOURCE Then  'Or objThis.Engine Is Nothing 
            If objThis.DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                scDS.SplitterDistance = SplitSrc
            Else
                gbDel.Visible = False
                scDS.SplitterDistance = SplitSrcNoDel
            End If
            gbSource.Visible = True
            gbTarget.Visible = False
            'gbStruct.Height = SplitSrcNoDel - 5
        Else
            If objThis.DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                scDS.SplitterDistance = SplitTgt
            Else
                gbDel.Visible = False
                scDS.SplitterDistance = SplitTgtNoDel
            End If
            gbSource.Visible = False
            gbTarget.Visible = True
            gbTarget.Location = gbSource.Location
            'gbStruct.Height = SplitTgtNoDel - 5
        End If
        txtListViewText.Visible = False
        cmbListViewCombo.Visible = False

    End Sub

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)

        sl = New MyListener(lvFieldAttrs)

    End Sub

    Private Sub sl_MyScroll(ByVal sender As Object, ByVal e As System.EventArgs) Handles sl.MyScroll

        cmbListViewCombo.Visible = False
        txtListViewText.Visible = False

    End Sub

    Private Sub ctlDatastore_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LostFocus

        txtListViewText.Visible = False
        cmbListViewCombo.Visible = False

    End Sub

    '/// Added by TK 12/07
    Private Sub OnlvAttr_resize(ByVal sender As Object, ByVal e As EventArgs) Handles lvFieldAttrs.Resize

        lvFieldAttrs.Columns(0).Width = (lvFieldAttrs.Width / 2) - 12
        lvFieldAttrs.Columns(1).Width = (lvFieldAttrs.Width / 2) - 12

    End Sub

    Function InitControls() As Boolean

        Me.tvDatastoreStructures.ImageList = modDeclares.imgListSmall
        tvDatastoreStructures.Font = New Font("Arial", 8, FontStyle.Bold)

        txtListViewText.Visible = False
        cmbListViewCombo.Visible = False

        lvFieldAttrs.FullRowSelect = True

        '//Add listview columns
        tvFields.HideSelection = False

        lvFieldAttrs.SmallImageList = imgListSmall

        'pnlSelect.Visible = True
        'pnlFields.Visible = True

        '//Modified by TK  12/2009
        cmbDatastoreType.Items.Clear()
        cmbDatastoreType.Items.Add(New Mylist("Binary", enumDatastore.DS_BINARY))
        'cmbDatastoreType.Items.Add(New Mylist("Text", enumDatastore.DS_TEXT))
        cmbDatastoreType.Items.Add(New Mylist("Delimited", enumDatastore.DS_DELIMITED))
        cmbDatastoreType.Items.Add(New Mylist("XML", enumDatastore.DS_XML))
        cmbDatastoreType.Items.Add(New Mylist("Relational", enumDatastore.DS_RELATIONAL))
        cmbDatastoreType.Items.Add(New Mylist("DB2LOAD", enumDatastore.DS_DB2LOAD))
        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            cmbDatastoreType.Items.Add(New Mylist("HSSUNLOAD", enumDatastore.DS_HSSUNLOAD))
        End If
        cmbDatastoreType.Items.Add(New Mylist("IMSDB", enumDatastore.DS_IMSDB))
        cmbDatastoreType.Items.Add(New Mylist("VSAM", enumDatastore.DS_VSAM))
        cmbDatastoreType.Items.Add(New Mylist("UTSCDC", enumDatastore.DS_UTSCDC))
        cmbDatastoreType.Items.Add(New Mylist("IMSCDC", enumDatastore.DS_IMSCDC))
        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            cmbDatastoreType.Items.Add(New Mylist("IMSCDC LE", enumDatastore.DS_IMSCDCLE))
        End If
        cmbDatastoreType.Items.Add(New Mylist("DB2CDC", enumDatastore.DS_DB2CDC))
        cmbDatastoreType.Items.Add(New Mylist("VSAMCDC", enumDatastore.DS_VSAMCDC))
        'cmbDatastoreType.Items.Add(New Mylist("IBM Event", enumDatastore.DS_IBMEVENT))
        cmbDatastoreType.Items.Add(New Mylist("OracleCDC", enumDatastore.DS_ORACLECDC))

        '
        'cmbDatastoreType.Items.Add(New Mylist("XML CDC", enumDatastore.DS_XMLCDC))
        'cmbDatastoreType.Items.Add(New Mylist("Trigger based CDC", enumDatastore.DS_TRBCDC))
        'cmbDatastoreType.Items.Add(New Mylist("IMS LE Batch", enumDatastore.DS_IMSLEBATCH))
        cmbDatastoreType.Items.Add(New Mylist("SubVariable", enumDatastore.DS_SUBVAR))

        '//


        cmbCharacterCode.Items.Clear()
        cmbCharacterCode.Items.Add(New Mylist("System", DS_CHARACTERCODE_SYSTEM))
        cmbCharacterCode.Items.Add(New Mylist("ASCII", DS_CHARACTERCODE_ASCII))
        cmbCharacterCode.Items.Add(New Mylist("EBCDIC", DS_CHARACTERCODE_EBCDIC))
        If IsNewObj = True Then cmbCharacterCode.SelectedIndex = 0

        '//
        setColCombo()
        setRowCombo()
        setTextCombo()

        SetCombos()

        UpdateControls()

    End Function

    Function UpdateControls() As Boolean

        'pnlFields.Update()
        Me.Refresh()

    End Function

    Friend Function ShowHideControls(ByVal DatastoreType As enumDatastore) As Boolean

        '//Show/Hide Src/Tgt and Delimited DS option  (TK 12/18/2009)
        Pnt.X = 3
        rbLU.Visible = False

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then  'Or objThis.Engine Is Nothing 
            If DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                scDS.SplitterDistance = SplitSrc
                Pnt.Y = 300
                gbDel.Location = Pnt
            Else
                gbDel.Visible = False
                scDS.SplitterDistance = SplitSrcNoDel
            End If
            gbSource.Visible = True
            gbTarget.Visible = False
            'gbStruct.Height = SplitSrcNoDel - 5
            scDescriptions.Height = SplitSrcNoDel - 5
            'gbStruct.Update()
            'gbStruct.Refresh()
            cmbCharacterCode.Enabled = True
            txtException.Enabled = True
        Else
            If DatastoreType = enumDatastore.DS_DELIMITED Then
                gbDel.Visible = True
                scDS.SplitterDistance = SplitTgt
                Pnt.Y = 223
                gbDel.Location = Pnt
            Else
                gbDel.Visible = False
                scDS.SplitterDistance = SplitTgtNoDel
            End If
            gbSource.Visible = False
            gbTarget.Visible = True
            gbTarget.Location = gbSource.Location
            'gbStruct.Height = SplitTgtNoDel - 5
            scDescriptions.Height = SplitTgtNoDel - 5
            'gbStruct.Update()
            'gbStruct.Refresh()
            cmbCharacterCode.Enabled = False
            txtException.Enabled = True
        End If

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            gbExtProps.Enabled = True
        Else
            gbExtProps.Enabled = False
            'txtPoll.Text = ""
        End If
        '//AccessMethod is Enable for all except Relational,IMS,VSAM
        cmbAccessMethod.Enabled = Not (DatastoreType = enumDatastore.DS_VSAM Or _
        DatastoreType = enumDatastore.DS_IMSDB Or _
        DatastoreType = enumDatastore.DS_INCLUDE)
        'DatastoreType = enumDatastore.DS_RELATIONAL Or _

        SetAccessCombo(DatastoreType)


        'label UOW is visible to db2, ims and trigger based cdc
        lblUOW.Enabled = DatastoreType = modDeclares.enumDatastore.DS_DB2CDC 'Or _
        'DatastoreType = modDeclares.enumDatastore.DS_TRBCDC Or _
        'DatastoreType = enumDatastore.DS_SUBVAR Or _
        'DatastoreType = modDeclares.enumDatastore.DS_IMSCDC)

        txtUOW.Enabled = DatastoreType = modDeclares.enumDatastore.DS_DB2CDC 'Or _
        'DatastoreType = modDeclares.enumDatastore.DS_TRBCDC Or _
        'DatastoreType = enumDatastore.DS_SUBVAR Or _
        'DatastoreType = modDeclares.enumDatastore.DS_IMSCDC)

        SetPortOrMQ()

        '//Following DS has only EBCDIC format
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
        'objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
        'objThis.DatastoreType = enumDatastore.DS_IMSLEBATCH)

        '//only for relational  added by KS and TK 11/6/2006
        'chkCmmtKey.Enabled = (objThis.DatastoreType = modDeclares.enumDatastore.DS_RELATIONAL)


        '//only for CDC
        chkSkipChangeCheck.Enabled = (DatastoreType = modDeclares.enumDatastore.DS_UTSCDC Or _
                                  DatastoreType = modDeclares.enumDatastore.DS_VSAMCDC Or _
                                  DatastoreType = enumDatastore.DS_SUBVAR Or _
                                  DatastoreType = enumDatastore.DS_ORACLECDC Or _
                                  DatastoreType = enumDatastore.DS_IMSCDCLE Or _
                                  DatastoreType = enumDatastore.DS_IMSCDC Or _
                                  DatastoreType = modDeclares.enumDatastore.DS_DB2CDC)

        'SetPoll()

        '//added by npatel on 9/6/05
        SetListItemByValue(cmbDatastoreType, objThis.DatastoreType, False)

        If objThis.DsDirection = DS_DIRECTION_SOURCE Or objThis.Engine Is Nothing Then
            gbSource.Enabled = True
        End If

    End Function

#End Region

#Region "Click and Help Events"

    Public Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        'If objThis.DsDirection = DS_DIRECTION_TARGET Then
        '    If objThis.DatastoreType = enumDatastore.DS_RELATIONAL Then
        '        KeyForRel = HasKeyForRel()
        '    End If
        'End If

        If Save() = True Then
            If objThis.Engine IsNot Nothing Then
                objThis.SetIsMapped()
            End If
            HiLiteMappedNodes(tvDatastoreStructures, objThis)
            '/// Added for Addflow
            objThis.Engine.ObjAddFlowCtl.RefreshAddFlow()
        End If

    End Sub

    Public Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Dim ret As MsgBoxResult

        Try
            If objThis.IsModified = True Then
                ret = MsgBox("Do you want to save change(s) made to the opened form?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel, MsgTitle)
                If ret = MsgBoxResult.Yes Then
                    If Save() = True Then
                        RaiseEvent Saved(Me, objThis)
                    Else
                        MsgBox("Save was NOT successful" & Chr(3) & _
                               "Please review your changes and save again.", MsgBoxStyle.Exclamation, "Save Operation Failed")
                        Exit Sub
                    End If
                ElseIf ret = MsgBoxResult.No Then
                    objThis.IsModified = False
                ElseIf ret = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Me.Visible = False
            RaiseEvent Closed(Me, objThis)

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmdClose_click")
        End Try

    End Sub

    'Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If SelectFirstMatchingNode(tvFields, txtSearchField.Text) Is Nothing Then
    '        MsgBox("No matching node found for entered text", MsgBoxStyle.Critical, MsgTitle)
    '    End If

    'End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Dim HelpObj As String

        HelpObj = Me.objThis.DsDirection

        If HelpObj = "T" Then
            ShowHelp(modHelp.HHId.H_Targets)
        Else
            ShowHelp(modHelp.HHId.H_Sources)
        End If


    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles chkIMSPathData.KeyDown, chkSkipChangeCheck.KeyDown, cmbAccessMethod.KeyDown, cmbColDelimiter.KeyDown, cmbOperationType.KeyDown, cmbRowDelimiter.KeyDown, cmbTextQualifier.KeyDown, tvDatastoreStructures.KeyDown, txtDatastoreDesc.KeyDown, txtUOW.KeyDown, txtDatastoreName.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select

    End Sub

    Public Sub cmbDatastoreType_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbDatastoreType.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Datastore_Type)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

    Public Sub txtPhysicalSource_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPhysicalSource.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Physical_Name)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If
    End Sub

    Public Sub cmbCharacterCode_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbCharacterCode.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Character_Code)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

    Public Sub txtPortOrMQMgr_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtPortOrMQMgr.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Port)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

    Public Sub txtException_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtException.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Exception)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

    Public Sub fields_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvFields.KeyDown, cmbListViewCombo.KeyDown

        If e.KeyCode = Keys.F1 Then
            Dim HelpObj As String = Me.objThis.DsDirection
            If HelpObj = "S" Then
                ShowHelp(modHelp.HHId.H_Source_Field_Attribute_Options)
            Else
                ShowHelp(modHelp.HHId.H_Target_Field_Attribute_Options)
            End If
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

#End Region

#Region "Text and Combo Box Events"

    Private Sub cmbAccessMethod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAccessMethod.SelectedIndexChanged

        SetPortOrMQ()
        'SetPoll()

    End Sub

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

    Sub SetListViewCombo(ByVal Attribute As enumFieldAttributes) '// modified 11/6/2006 by TK and KS

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim x As Integer
        Dim selName As String
        Dim isKey As String
        Dim keyName As String

        IsEventFromCode = True
        cmbListViewCombo.Items.Clear()
        IsEventFromCode = False

        IsEventFromCode = True
        Select Case Attribute
            '// added by TK and KS 11/6/2006
            '/// Modified by TK 1/10/2007 for child selections
            Case modDeclares.enumFieldAttributes.ATTR_FKEY

                Dim objSel As clsDSSelection
                Dim tempstring As String = ""
                Dim tempobj As clsDSSelection
                Dim childSel As clsDSSelection = Nothing
                Dim templist As New ArrayList
                Dim StrSelName As String
                Dim fld As clsField
                Dim flag As Boolean = False

                '/// initialize variables
                cmbListViewCombo.Items.Add(New Mylist("", ""))

                If tvDatastoreStructures.SelectedNode.Checked = True Then
                    tempstring = tvDatastoreStructures.SelectedNode.Text
                    '/// a structure cannot have itself as a Foreign Key,  
                    '/// so add it to the exclusion list
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
                    '// DSselection and all it's children's children etc... and add these  
                    '// to the exclusion list
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
                    '/// now go through the all the datastore selections
                    '/// and add key fields of all structures not on the exclusion list
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


            Case modDeclares.enumFieldAttributes.ATTR_DATEFORMAT
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("MMDDYY", "MMDDYY"))
                cmbListViewCombo.Items.Add(New Mylist("DDMMYY", "DDMMYY"))
                cmbListViewCombo.Items.Add(New Mylist("YYDDMM", "YYDDMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYMMDD", "YYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYDDMM", "YYYYDDMM"))
                cmbListViewCombo.Items.Add(New Mylist("YYYYMMDD", "YYYYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("MMDDYYYY", "MMDDYYYY"))
                cmbListViewCombo.Items.Add(New Mylist("CIYYMMDD", "CIYYMMDD"))
                cmbListViewCombo.Items.Add(New Mylist("HHMMSS", "HHMMSS"))
                cmbListViewCombo.Items.Add(New Mylist("ALPHA", "ALPHA"))
                cmbListViewCombo.Items.Add(New Mylist("ALNUM", "ALNUM"))

            Case modDeclares.enumFieldAttributes.ATTR_RETYPE
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("FCHAR", "FCHAR"))
                cmbListViewCombo.Items.Add(New Mylist("BINARY", "BINARY"))
                cmbListViewCombo.Items.Add(New Mylist("DATE", "DATE"))
                cmbListViewCombo.Items.Add(New Mylist("VCHAR", "VCHAR"))
                cmbListViewCombo.Items.Add(New Mylist("FULLW", "FULLW"))
                cmbListViewCombo.Items.Add(New Mylist("HALFW", "HALFW"))
                cmbListViewCombo.Items.Add(New Mylist("NUMBER", "NUMBER"))
                'cmbListViewCombo.Items.Add(New Mylist("TIME", "TIME"))
                'cmbListViewCombo.Items.Add(New Mylist("TIMESTAMP", "TIMESTAMP"))
                'cmbListViewCombo.Items.Add(New Mylist("VARCHAR", "VARCHAR"))
                'cmbListViewCombo.Items.Add(New Mylist("ZONE", "ZONE"))
                'cmbListViewCombo.Items.Add(New Mylist("XMLATTC", "XMLATTC"))
                'cmbListViewCombo.Items.Add(New Mylist("XMLATTPC", "XMLATTPC"))
                'cmbListViewCombo.Items.Add(New Mylist("XMLCDATA", "XMLCDATA"))
                'cmbListViewCombo.Items.Add(New Mylist("XMLPCDATA", "XMLPCDATA"))

            Case modDeclares.enumFieldAttributes.ATTR_INVALID
                cmbListViewCombo.Items.Add(New Mylist("", ""))
                cmbListViewCombo.Items.Add(New Mylist("SETNULL", "SETNULL"))
                cmbListViewCombo.Items.Add(New Mylist("SETSPACE", "SETSPACE"))
                cmbListViewCombo.Items.Add(New Mylist("SETZERO", "SETZERO"))
                cmbListViewCombo.Items.Add(New Mylist("SETEXP", "SETEXP"))
                cmbListViewCombo.Items.Add(New Mylist("0001-01-01", "0001-01-01"))
                cmbListViewCombo.Items.Add(New Mylist("00.00.00", "00.00.00"))
                cmbListViewCombo.Items.Add(New Mylist("00.00", "00.00"))
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
            LogError(ex, "ctlDatastore SetCombos")
        End Try

    End Sub

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
            LogError(ex, "ctlDatastore cmbListViewCombo_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub txtListViewText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtListViewText.TextChanged

        Try
            If IsEventFromCode = True Then Exit Sub

            If Not (tvFields.SelectedNode Is Nothing) Then
                CType(tvFields.SelectedNode.Tag, clsField).SetSingleFieldAttr(lvItem.Index, txtListViewText.Text)
                lvItem.SubItems(1).Text = txtListViewText.Text
                CType(tvFields.SelectedNode.Tag, clsField).IsModified = True
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore txtListViewText_TextChanged")
        End Try

    End Sub

    '//added by npatel on 9/6/05
    Private Sub cmbDatastoreType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDatastoreType.SelectedIndexChanged

        If IsEventFromCode = True Then
            Exit Sub
        End If
        objThis.DatastoreType = CType(cmbDatastoreType.Items(cmbDatastoreType.SelectedIndex), Mylist).ItemData
        ShowHideControls(objThis.DatastoreType)

    End Sub

    Sub SetPortOrMQ()

        txtPortOrMQMgr.Enabled = ((cmbAccessMethod.Enabled = True) And ((cmbAccessMethod.Text = "TCP/IP") Or _
                                                                        (cmbAccessMethod.Text = "CDC Store")))
        lblPortorMQMgr.Enabled = ((cmbAccessMethod.Enabled = True) And ((cmbAccessMethod.Text = "TCP/IP") Or _
                                                                        (cmbAccessMethod.Text = "CDC Store")))

    End Sub

    'Sub SetPoll()

    '    txtPoll.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
    '                      objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
    '                      objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
    '                      cmbAccessMethod.Text = "VSAM CDCStore") Or _
    '                      objThis.DatastoreType = enumDatastore.DS_UTSCDC
    '    If txtPoll.Enabled = False Then
    '        txtPoll.Text = ""
    '    End If

    '    lblPoll.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
    '                        objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
    '                        objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
    '                        cmbAccessMethod.Text = "VSAM CDCStore") Or _
    '                        objThis.DatastoreType = enumDatastore.DS_UTSCDC

    '    txtRestart.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
    '                      objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
    '                      objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
    '                      objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
    '                      cmbAccessMethod.Text = "VSAM CDCStore") Or _
    '                      objThis.DatastoreType = enumDatastore.DS_UTSCDC
    '    If txtRestart.Enabled = False Then
    '        txtRestart.Text = ""
    '    End If

    '    lblRestart.Enabled = ((objThis.DatastoreType = enumDatastore.DS_VSAMCDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_SUBVAR Or _
    '                        objThis.DatastoreType = enumDatastore.DS_ORACLECDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_IMSCDCLE Or _
    '                        objThis.DatastoreType = enumDatastore.DS_IMSCDC Or _
    '                        objThis.DatastoreType = enumDatastore.DS_DB2CDC) And _
    '                        cmbAccessMethod.Text = "VSAM CDCStore") Or _
    '                        objThis.DatastoreType = enumDatastore.DS_UTSCDC

    'End Sub

    Sub SetAccessCombo(ByVal DStype As enumDatastore)

        If objThis.DsDirection = DS_DIRECTION_SOURCE Then
            Select Case DStype
                Case enumDatastore.DS_BINARY, enumDatastore.DS_TEXT, enumDatastore.DS_DELIMITED, enumDatastore.DS_HSSUNLOAD, _
                enumDatastore.DS_XML, enumDatastore.DS_DB2LOAD, enumDatastore.DS_IBMEVENT
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("TCP/IP", DS_ACCESSMETHOD_IP))
                    cmbAccessMethod.SelectedIndex = 0

                Case enumDatastore.DS_UTSCDC, enumDatastore.DS_IMSCDC, enumDatastore.DS_IMSCDCLE, _
                enumDatastore.DS_VSAMCDC, enumDatastore.DS_SUBVAR
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("CDC Store", DS_ACCESSMETHOD_CDCSTORE))
                    cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("TCP/IP", DS_ACCESSMETHOD_IP))
                    'cmbAccessMethod.Items.Add(New Mylist("VSAM CDCStore", DS_ACCESSMETHOD_VSAM))

                    cmbAccessMethod.SelectedIndex = 0

                    'added July 2012
                Case enumDatastore.DS_DB2CDC
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.SelectedIndex = 0

                Case enumDatastore.DS_RELATIONAL
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("CDC Store", DS_ACCESSMETHOD_CDCSTORE))
                    'cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    'cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("TCP/IP", DS_ACCESSMETHOD_IP))

                    cmbAccessMethod.SelectedIndex = 0

                Case enumDatastore.DS_ORACLECDC
                    cmbAccessMethod.Items.Clear()
                    cmbAccessMethod.Items.Add(New Mylist("CDC Store", DS_ACCESSMETHOD_CDCSTORE))
                    'cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
                    cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
                    cmbAccessMethod.Items.Add(New Mylist("TCP/IP", DS_ACCESSMETHOD_IP))

                    cmbAccessMethod.SelectedIndex = 0

                Case Else
                    cmbAccessMethod.Items.Clear()
            End Select
        Else
            cmbAccessMethod.Items.Clear()
            cmbAccessMethod.Items.Add(New Mylist("File", DS_ACCESSMETHOD_FILE))
            cmbAccessMethod.Items.Add(New Mylist("MQSeries", DS_ACCESSMETHOD_MQSERIES))
            cmbAccessMethod.Items.Add(New Mylist("TCP/IP", DS_ACCESSMETHOD_IP))
            cmbAccessMethod.SelectedIndex = 0
        End If

        If cmbAccessMethod.Enabled = True Then
            SetListItemByValue(cmbAccessMethod, objThis.DsAccessMethod, False)
        End If


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
            LogError(ex, "ctlDatastore SetCharCodeCombo")
        End Try

    End Sub

    Private Sub cmbOperationType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOperationType.SelectedIndexChanged

        Try
            Select Case CType(cmbOperationType.SelectedItem, Mylist).Name
                Case "", "Insert", "Replace"
                    check4key = False               'added 7/2012 by TK
                    cbKeyChng.Enabled = False
                Case "Update", "Delete", "Change", "Modify"
                    check4key = True                'added 7/2012 by TK
                    cbKeyChng.Enabled = True
            End Select

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbOperationType_SelectedIndexChanged")
        End Try

    End Sub

    Private Sub cbKeyChng_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbKeyChng.CheckedChanged

        Try
            If cbKeyChng.Checked = True Then
                objThis.IsKeyChng = "1"
            Else
                objThis.IsKeyChng = "0"
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore cbKeyChng_CheckedChanged")
        End Try

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

    Private Sub cmbColDelimiter_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbColDelimiter.KeyDown

        Try
            Dim cmbTxt As String = ""

            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) And cmbColDelimiter.Text <> "" And _
            cmbColDelimiter.SelectedIndex < 1 Then

                cmbTxt = cmbColDelimiter.Text
                objThis.ColumnDelimiter = cmbTxt
                cmbColDelimiter.Items.Add(New Mylist(cmbTxt, cmbTxt))
                SetListItemByValue(cmbColDelimiter, objThis.ColumnDelimiter, True)
                OnChange(Me, New EventArgs)
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbColDelimiter_KeyDown")
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

    Private Sub cmbRowDelimiter_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbRowDelimiter.KeyDown

        Try
            Dim cmbTxt As String = ""

            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) And cmbRowDelimiter.Text <> "" And _
            cmbRowDelimiter.SelectedIndex < 1 Then

                cmbTxt = cmbRowDelimiter.Text
                objThis.RowDelimiter = cmbTxt
                cmbRowDelimiter.Items.Add(New Mylist(cmbTxt, cmbTxt))
                SetListItemByValue(cmbRowDelimiter, objThis.RowDelimiter, True)
                OnChange(Me, New EventArgs)
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbRowDelimiter_KeyDown")
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

    Private Sub cmbTextQualifier_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles cmbTextQualifier.KeyDown

        Try
            Dim cmbTxt As String = ""

            If (e.KeyCode = Keys.Enter Or e.KeyCode = Keys.Tab) And cmbTextQualifier.Text <> "" And _
            cmbTextQualifier.SelectedIndex < 1 Then

                cmbTxt = cmbTextQualifier.Text
                objThis.TextQualifier = cmbTxt
                cmbTextQualifier.Items.Add(New Mylist(cmbTxt, cmbTxt))
                SetListItemByValue(cmbTextQualifier, objThis.TextQualifier, True)
                OnChange(Me, New EventArgs)
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore cmbTextQualifier_KeyDown")
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
            If coltxt <> "Tilde{~}" And texttxt <> "Tilde{~}" Then
                cmbRowDelimiter.Items.Add(New Mylist("Tilde{~}", DS_DELIMITER_TILDE))
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

#End Region

#Region "Treeview and Listview Functions"

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        'txtFieldDesc.Text = ""
        'e.Node.SelectedImageIndex = e.Node.ImageIndex
        'ShowFieldAttributes(e.Node)

        If prevfld IsNot Nothing Then
            If prevfld.FieldDescModified = True Then
                prevfld.FieldDesc = txtFieldDesc.Text
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

    Private Sub tvDatastoreStructures_AfterCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvDatastoreStructures.AfterCheck
        '//Couple of things to do when user check/uncheck any node

        Try
            If IsEventFromCode = True Then Exit Try

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
                        Next
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
                        IsEventFromCode = False
                    End If
                End If
            Else
                If CType(e.Node.Tag, INode).IsFolderNode = False Then
                    If CType(e.Node.Tag, INode).Type = NODE_STRUCT_SEL Then
                        Dim objsel As clsStructureSelection = CType(e.Node.Tag, clsStructureSelection)
                        For Each DSsel As clsDSSelection In objsel.ObjDSselections
                            If DSsel.IsMapped = True And DSsel.ObjDatastore Is objThis Then
                                If MsgBox("The selection, " & e.Node.Text & ", that you have unchecked has fields that are mapped in a Procedure," & Chr(13) & "are you sure you would like to remove the selection from this Datastore?", MsgBoxStyle.OkCancel, "Remove Mapped Selection?") = MsgBoxResult.Cancel Then
                                    e.Node.Checked = True
                                    Exit Sub
                                End If
                            End If
                        Next

                    End If
                End If
            End If
            tvDatastoreStructures.SelectedNode = e.Node
            OnChange(Me, New EventArgs)


        Catch ex As Exception
            LogError(ex, "ctlDatastore tvdatastorestructures.aftercheck")
        End Try

    End Sub

    Private Sub tvDatastoreStructures_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvDatastoreStructures.AfterSelect

        If CType(tvDatastoreStructures.SelectedNode.Tag, INode).Type = NODE_FO_STRUCT Then Exit Sub
        ShowHideFieldAttr()

    End Sub

    Private Sub lvFieldAttrs_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvFieldAttrs.MouseClick

        'Get the item on the row that is clicked.
        lvItem = Me.lvFieldAttrs.GetItemAt(e.X, e.Y)
        ' Make sure that an item is clicked.

        If Not (lvItem Is Nothing) Then
            Select Case lvItem.Index
                Case modDeclares.enumFieldAttributes.ATTR_NCHILDREN, _
                     modDeclares.enumFieldAttributes.ATTR_LEVEL, _
                     modDeclares.enumFieldAttributes.ATTR_TIMES, _
                     modDeclares.enumFieldAttributes.ATTR_OCCURS, _
                     modDeclares.enumFieldAttributes.ATTR_DATATYPE, _
                     modDeclares.enumFieldAttributes.ATTR_OFFSET, _
                     modDeclares.enumFieldAttributes.ATTR_LENGTH, _
                     modDeclares.enumFieldAttributes.ATTR_SCALE, _
                     modDeclares.enumFieldAttributes.ATTR_CANNULL, _
                     modDeclares.enumFieldAttributes.ATTR_DATEFORMAT

                    cmbListViewCombo.Visible = False
                    txtListViewText.Visible = False
                    Exit Sub
            End Select

            ' Get the bounds of the item that is clicked.
            Dim ClickedItem As Rectangle = lvItem.SubItems(1).Bounds
            Dim Col1XPos As Integer = ClickedItem.X

            ' Adjust the location to account for the location of the ListView.
            ClickedItem.Y += Me.lvFieldAttrs.Top
            ClickedItem.X += Me.lvFieldAttrs.Left

            '/// Now make sure text and combo boxes don't scroll off the listview
            If Col1XPos > lvFieldAttrs.DisplayRectangle.Right Then
                'scrolled off screen
                Return
            End If

            '/// make sure text and combo boxes are not so wide that they scroll off the screen
            If Col1XPos + ClickedItem.Width > lvFieldAttrs.DisplayRectangle.Right Then
                'adjusts right edge of Box and width accoring to where edge of Control is
                ClickedItem.Width = lvFieldAttrs.DisplayRectangle.Right - Col1XPos
            End If

            If Col1XPos < 0 Then
                'right edge of box is scrolled off screen to the left
                'so adjust left edge and width of left
                ClickedItem.X = ClickedItem.X - Col1XPos
                ClickedItem.Width = ClickedItem.Width + Col1XPos
            End If

            '///for debug
            'MsgBox("clickedItem.x = " & ClickedItem.X & Chr(13) & "clickedItem.y = " & ClickedItem.Y & Chr(13) & "lvfldattr.left = " & lvFieldAttrs.Left & Chr(13) & "lvfldattr.right = " & lvFieldAttrs.Right & Chr(13) & "col1xpos = " & Col1XPos)

            If lvItem.Index = modDeclares.enumFieldAttributes.ATTR_ISKEY Or _
                lvItem.Index = modDeclares.enumFieldAttributes.ATTR_FKEY Or _
                lvItem.Index = modDeclares.enumFieldAttributes.ATTR_RETYPE Or _
                lvItem.Index = modDeclares.enumFieldAttributes.ATTR_EXTTYPE Or _
                lvItem.Index = modDeclares.enumFieldAttributes.ATTR_INVALID Then

                ' Assign calculated bounds to the ComboBox.
                Me.cmbListViewCombo.Bounds = ClickedItem

                ShowLVCombo(lvItem)
            ElseIf lvItem.Index = modDeclares.enumFieldAttributes.ATTR_DATEFORMAT Then
                'Do Not Edit DateFormat
            Else
                ' Assign calculated bounds to the TextBox.
                Me.txtListViewText.Bounds = ClickedItem

                ShowLVText(lvItem)
            End If
        End If

    End Sub

    Private Sub tvFields_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvFields.Click
        txtListViewText.Visible = False
        cmbListViewCombo.Visible = False

    End Sub

    Private Sub tvFields_Afterclick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvFields.NodeMouseClick, tvFields.NodeMouseDoubleClick

        If prevfld IsNot Nothing Then
            If prevfld.FieldDescModified = True Then
                prevfld.FieldDesc = txtFieldDesc.Text
            End If
        End If
        HiLiteFieldDescNodes(tvFields.TopNode, True, tvFields)
        HiLiteFieldKeyNodes(tvFields.TopNode, True, tvFields)
        IsEventFromCode = True
        txtFieldDesc.Text = ""
        IsEventFromCode = False
        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub tvDatastoreStructures_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvDatastoreStructures.KeyDown

        If tvDatastoreStructures.SelectedNode Is Nothing Then Exit Sub

        If e.KeyCode = Keys.Enter Then
            ShowHideFieldAttr()
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

#End Region

#Region "Object Functions"

    Public Function Save() As Boolean

        Try
            Dim OldName As String = ""
            Me.Cursor = Cursors.WaitCursor

            'If ValidateNewName(txtDatastoreName.Text) = False Then
            '    Save = False
            '    Me.Cursor = Cursors.Default
            '    Exit Function
            'End If

            If txtPhysicalSource.Text.Trim = "" Then
                MsgBox("'Physical Name' Must not be left blank", MsgBoxStyle.OkOnly, "Missing Physical Name")
                Save = False
                Exit Try
            End If

            '//save form data to the object before saving object to database
            If objThis.DatastoreName <> txtDatastoreName.Text Then
                objThis.IsRenamed = RenameDatastore(objThis, txtDatastoreName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtDatastoreName.Text = objThis.DatastoreName
            Else
                OldName = objThis.DatastoreName
                objThis.DatastoreName = txtDatastoreName.Text
            End If
            '//save structure selections too 
            If SaveCurrentSelection() = False Then
                Save = False
                Exit Try
            End If

            If objThis.DsDirection = DS_DIRECTION_TARGET Or objThis.IsLookUp = True Then
                If (objThis.DatastoreType = enumDatastore.DS_RELATIONAL Or objThis.DatastoreType = enumDatastore.DS_IMSDB Or objThis.DatastoreType = enumDatastore.DS_VSAM) And check4key = True Then
                    KeyForRel = HasKeyForRel()
                End If
            End If
            '?????? Make them set a key???????
            'If KeyForRel = False Then
            '    Exit Try
            'End If

            'If objThis.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    If SaveGlobalValues() = False Then
            '        Exit Function
            '    End If
            'End If

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    Save = False
                    Exit Try
                End If
            Else
                If objThis.Save = False Then
                    Save = False
                    Exit Try
                    'Else
                    '    If objThis.Engine Is Nothing Then
                    '        For Each sys As clsSystem In objThis.Environment.Systems
                    '            For Each eng As clsEngine In sys.Engines
                    '                For Each src As clsDatastore In eng.Sources
                    '                    If OldName = "" Then
                    '                        OldName = objThis.DatastoreName
                    '                    End If
                    '                    If src.DatastoreName = OldName Then
                    '                        src.DeleteATTR()
                    '                        src.DatastoreName = objThis.DatastoreName
                    '                        src = objThis.Clone(eng)  'UpdateChildDS(src)
                    '                        If src IsNot Nothing Then
                    '                            src.DsDirection = "S"
                    '                            If src.Save() = False Then
                    '                                MsgBox("Datastore: " & src.Text & " failed to update correctly in Engine: " & _
                    '                                eng.Text, MsgBoxStyle.Information, "Problem saving Datastore")
                    '                            Else
                    '                                src.DeleteATTR()
                    '                                src.InsertATTR()
                    '                            End If
                    '                        Else
                    '                            MsgBox("Datastore: " & src.Text & " failed to update correctly in Engine: " & _
                    '                            eng.Text, MsgBoxStyle.Information, "Problem saving Datastore")
                    '                        End If
                    '                    End If
                    '                Next
                    '                For Each tgt As clsDatastore In eng.Targets
                    '                    If OldName = "" Then
                    '                        OldName = objThis.DatastoreName
                    '                    End If
                    '                    If tgt.DatastoreName = OldName Then
                    '                        tgt.DeleteATTR()
                    '                        tgt.DatastoreName = objThis.DatastoreName
                    '                        tgt = objThis.Clone(eng)  'UpdateChildDS(src)
                    '                        If tgt IsNot Nothing Then
                    '                            tgt.DsDirection = "T"
                    '                            If tgt.Save() = False Then
                    '                                MsgBox("Datastore: " & tgt.Text & " failed to update correctly in Engine: " & _
                    '                                eng.Text, MsgBoxStyle.Information, "Problem saving Datastore")
                    '                            Else
                    '                                tgt.DeleteATTR()
                    '                                tgt.InsertATTR()
                    '                            End If
                    '                        Else
                    '                            MsgBox("Datastore: " & tgt.Text & " failed to update correctly in Engine: " & _
                    '                            eng.Text, MsgBoxStyle.Information, "Problem saving Datastore")
                    '                        End If
                    '                    End If
                    '                Next
                    '            Next
                    '        Next
                    '    End If
                End If
            End If

            Save = True

            cmdSave.Enabled = False

            If objThis.IsRenamed = True Then
                RaiseEvent Renamed(Me, objThis)
            Else
                RaiseEvent Saved(Me, objThis)
            End If

            objThis.IsRenamed = False

        Catch ex As Exception
            LogError(ex, "ctlDatastore Save()")
            Save = False
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Function

    '    Function SaveGlobalValues() As Boolean

    '        Dim ErrorMsg As String = ""
    '        Try
    '            If objThis.ExtTypeChar <> "" And objThis.ExtTypeNum = "" Then
    '                ErrorMsg = "You have set a Global ExtType without specifying a field type to apply this ExtType to." & Chr(13) & _
    '                "Please set a field type in the drop-down box next to the Global ExtType"
    '                GoTo msg
    '            End If
    '            If objThis.IfNullChar <> "" And objThis.IfNullNum = "" Then
    '                ErrorMsg = "You have set a Global IfNull without specifying a field type to apply this IfNull to." & Chr(13) & _
    '                "Please set a field type in the drop-down box next to the Global IfNull"
    '                GoTo Msg
    '            End If
    '            If objThis.InValidChar <> "" And objThis.InValidNum = "" Then
    '                ErrorMsg = "You have set a Global InValid without specifying a field type to apply this InValid to." & Chr(13) & _
    '                "Please set a field type in the drop-down box next to the Global InValid"
    '                GoTo Msg
    '            End If
    '            'If objThis.SaveGlobals() = False Then
    '            '    ErrorMsg = "An Error occured while saving Datastore-wide Field Properties to a file."
    '            '    GoTo Msg
    '            'End If

    'msg:        If ErrorMsg <> "" Then
    '                MsgBox(ErrorMsg, MsgBoxStyle.OkOnly, MsgTitle)
    '                Return False
    '                Exit Function
    '            End If

    '            Return True

    '        Catch ex As Exception
    '            LogError(ex, "ctlDatastore SaveGlobalValues")
    '            Return False
    '        End Try

    '    End Function

    Public Function EditObj(ByRef obj As clsDatastore, ByVal cNode As TreeNode, Optional ByVal nd As TreeNode = Nothing) As clsDatastore

        Try
            Me.Cursor = Cursors.WaitCursor

            Me.Enabled = True
            IsNewObj = False
            StartLoad()
            objThis = obj '//Load the form Datastore object
            objThis.LoadMe()
            objThis.LoadItems() 'True
            InitControls()
            ShowHideControls(objThis.DatastoreType)

            'If objThis.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    objThis.LoadGlobals()
            'End If

            'rbLU.Checked = objThis.IsLookUp
            If objThis.IsLookUp = True Then
                gbProp.Text = "Lookup Datastore Properties"
            Else
                gbProp.Text = "Datastore Properties"
            End If
            'rbLU.Enabled = rbLU.Checked

            UpdateFields()

            FillStructure(cNode, nd)

            'If objThis.Engine IsNot Nothing Then
            '    objThis.SetIsMapped()
            'End If

            HiLiteMappedNodes(tvDatastoreStructures, objThis)

            'If Not ActiveQueryDSlist Is Nothing Then
            '    If ActiveQueryDSlist.Contains(obj) = True Then
            '        Me.Enabled = False
            '    End If
            'End If

            'If objThis.DsDirection = DS_DIRECTION_TARGET Then
            '    If objThis.DatastoreType = enumDatastore.DS_RELATIONAL Then
            '        KeyForRel = HasKeyForRel()
            '    End If
            'End If

            EndLoad()
            Me.Cursor = Cursors.Default

            EditObj = objThis

        Catch ex As Exception
            LogError(ex, "ctlDatastore EditObj")
            EditObj = Nothing
            Me.Cursor = Cursors.Default
        End Try

    End Function

    '//cNode is current environment structure root node
    '// Not used ... Tom Karasch
    Friend Function NewObj(ByVal cNode As TreeNode, Optional ByVal DatastoreType As enumDatastore = modDeclares.enumDatastore.DS_UNKNOWN) As clsDatastore

        IsNewObj = True

        StartLoad()

        '// added by KS and TK 11/6/2006
        'If (txtPoll.Text.Trim = "") Then
        '    txtPoll.Text = "0"
        'End If

        objThis.DatastoreType = DatastoreType

        InitControls()

        ShowHideControls(DatastoreType)

        FillStructure(cNode)

        txtDatastoreName.SelectAll()

        EndLoad()

        NewObj = objThis

        'doAgain:
        '        Select Case Me.ShowDialog
        '            Case Windows.Forms.DialogResult.OK
        '                NewObj = objThis
        '            Case Windows.Forms.DialogResult.Retry
        '                tvDatastoreStructures.ExpandAll()
        '                GoTo doAgain
        '            Case Else
        '                Return Nothing
        '        End Select

    End Function

    '//On Edit call this function
    '//Set values from obj to form controls
    Private Function UpdateFields() As Boolean

        txtDatastoreName.Text = objThis.DatastoreName
        txtDatastoreDesc.Text = objThis.DatastoreDescription
        txtPhysicalSource.Text = objThis.DsPhysicalSource
        txtException.Text = objThis.ExceptionDatastore
        txtUOW.Text = objThis.DsUOW
        'txtPoll.Text = objThis.Poll  '// add by TK 12/22/09
        'txtRestart.Text = objThis.Restart

        '//// Added 7/2012 by TK
        txtMTD.Text = objThis.MTDfile
        txtHostName.Text = objThis.DsHostName
        cbUseFile.Checked = objThis.UseMTD


        If cmbAccessMethod.Enabled = True Then
            SetListItemByValue(cmbAccessMethod, objThis.DsAccessMethod, False)
        End If
        SetListItemByValue(cmbCharacterCode, objThis.DsCharacterCode, False)

        If objThis.DsAccessMethod = DS_ACCESSMETHOD_IP Or objThis.DsAccessMethod = DS_ACCESSMETHOD_CDCSTORE Then
            txtPortOrMQMgr.Text = objThis.DsPort
        ElseIf objThis.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Then
            txtPortOrMQMgr.Text = objThis.DsQueMgr
        End If


        If objThis.DsDirection = DS_DIRECTION_TARGET Then
            SetListItemByValue(cmbOperationType, objThis.OperationType, False)
        End If

        'Me.chkCmmtKey.Checked = IIf(objThis.IsCmmtKey = "1", True, False)  '// added by TK and KS 11/6/2006
        Me.chkIMSPathData.Checked = IIf(objThis.IsIMSPathData = "1", True, False)
        'Me.chkOrdered.Checked = IIf(objThis.IsOrdered = "1", True, False)
        Me.chkSkipChangeCheck.Checked = IIf(objThis.IsSkipChangeCheck = "1", True, False)
        'Me.chkBeforeImage.Checked = IIf(objThis.IsBeforeImage = "1", True, False) '//new by npatel on 8/10/05
        Me.cbKeyChng.Checked = IIf(objThis.IsKeyChng = "1", True, False)

        'If SetListItemByValue(cmbColDelimiter, objThis.ColumnDelimiter) = False Then
        '    cmbColDelimiter.Text = objThis.ColumnDelimiter
        'End If
        SetListItemByValue(cmbColDelimiter, objThis.ColumnDelimiter, True)
        'If SetListItemByValue(cmbRowDelimiter, objThis.RowDelimiter) = False Then
        '    cmbRowDelimiter.Text = objThis.RowDelimiter
        'End If
        SetListItemByValue(cmbRowDelimiter, objThis.RowDelimiter, True)
        'If SetListItemByValue(cmbTextQualifier, objThis.TextQualifier) = False Then
        '    cmbTextQualifier.Text = objThis.TextQualifier
        'End If
        SetListItemByValue(cmbTextQualifier, objThis.TextQualifier, True)

        SetListItemByValue(cbIfSpaceChar, objThis.ExtTypeChar, False)
        SetListItemByValue(cbIfSpaceNum, objThis.ExtTypeNum, False)
        SetListItemByValue(cbIfNullChar, objThis.IfNullChar, False)
        SetListItemByValue(cbIfNullNum, objThis.IfNullNum, False)
        SetListItemByValue(cbInvalidChar, objThis.InValidChar, False)
        SetListItemByValue(cbInvalidNum, objThis.InValidNum, False)

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
                GetCharDesc = "Semicolun"
            Case Else
                GetCharDesc = strChar
        End Select

    End Function

    'Added for new tree structure
    'Function UpdateChildDS(ByVal inDS As clsDatastore) As clsDatastore

    '    Try
    '        Dim DS As New clsDatastore

    '        'If ValidateNewName(txtDatastoreName.Text) = False Then
    '        '    Exit Function
    '        'End If

    '        '//save form data to the object before saving object to database
    '        'If DS.DatastoreName <> txtDatastoreName.Text Then
    '        '    DS.IsRenamed = RenameDatastore(DS, txtDatastoreName.Text)
    '        'End If

    '        'If DS.IsRenamed = False Then
    '        '    txtDatastoreName.Text = DS.DatastoreName
    '        'Else
    '        DS.DatastoreName = txtDatastoreName.Text
    '        'End If

    '        DS.DatastoreDescription = txtDatastoreDesc.Text

    '        If DS.DsDirection <> "" Then
    '            If DS.DsDirection = DS_DIRECTION_TARGET Then
    '                DS.OperationType = GetStr(GetVal(CType(cmbOperationType.SelectedItem, Mylist).ItemData))
    '            End If
    '        End If

    '        If cmbCharacterCode.Enabled = True Then
    '            If CType(cmbCharacterCode.SelectedItem, Mylist) IsNot Nothing Then
    '                DS.DsCharacterCode = CType(cmbCharacterCode.SelectedItem, Mylist).ItemData
    '            Else
    '                DS.DsCharacterCode = ""
    '            End If
    '        Else
    '            DS.DsCharacterCode = ""
    '        End If

    '        '//npatel 10/7/05
    '        If cmbAccessMethod.Enabled = True Then
    '            Dim Itm As Mylist
    '            Itm = cmbAccessMethod.SelectedItem
    '            If (Itm Is Nothing) = False Then
    '                If Itm.ItemData = DS_ACCESSMETHOD_IP Then
    '                    DS.DsPort = txtPortOrMQMgr.Text
    '                ElseIf Itm.ItemData = DS_ACCESSMETHOD_MQSERIES Then
    '                    DS.DsQueMgr = txtPortOrMQMgr.Text
    '                Else
    '                    DS.DsPort = txtPortOrMQMgr.Text
    '                End If
    '            End If
    '            DS.DsAccessMethod = CType(cmbAccessMethod.SelectedItem, Mylist).ItemData
    '        Else
    '            DS.DsAccessMethod = ""
    '            DS.DsQueMgr = ""
    '            DS.DsPort = ""
    '        End If

    '        DS.Poll = txtPoll.Text  '// added by TK and KS 11/6/2006
    '        DS.DsPhysicalSource = txtPhysicalSource.Text
    '        DS.ExceptionDatastore = txtException.Text
    '        DS.DsUOW = txtUOW.Text

    '        'DS.IsCmmtKey = IIf(chkCmmtKey.Checked, "1", "0") '// added by TK and KS 11/6/2006
    '        DS.IsIMSPathData = IIf(chkIMSPathData.Checked, "1", "0")
    '        'DS.IsOrdered = IIf(chkOrdered.Checked, "1", "0")
    '        DS.IsSkipChangeCheck = IIf(chkSkipChangeCheck.Checked, "1", "0")
    '        'DS.IsBeforeImage = IIf(chkBeforeImage.Checked, "1", "0") '// new by npatel on 8/10/05

    '        If DS.DatastoreType = modDeclares.enumDatastore.DS_DELIMITED Then
    '            If (cmbColDelimiter.SelectedItem Is Nothing) = True Then
    '                DS.ColumnDelimiter = Me.cmbColDelimiter.Text
    '            Else
    '                DS.ColumnDelimiter = CType(cmbColDelimiter.SelectedItem, Mylist).ItemData
    '            End If

    '            If (cmbRowDelimiter.SelectedItem Is Nothing) = True Then
    '                DS.RowDelimiter = Me.cmbRowDelimiter.Text
    '            Else
    '                DS.RowDelimiter = CType(cmbRowDelimiter.SelectedItem, Mylist).ItemData
    '            End If

    '            If (cmbTextQualifier.SelectedItem Is Nothing) = True Then
    '                DS.TextQualifier = Me.cmbTextQualifier.Text
    '            Else
    '                DS.TextQualifier = CType(cmbTextQualifier.SelectedItem, Mylist).ItemData
    '            End If
    '        End If

    '        '//before we clear old selection save it to another array so later on while saving this object we can compare for any new addition in selection list
    '        DS.OldObjSelections.Clear()
    '        DS.OldObjSelections = DS.ObjSelections.Clone

    '        '//clear previous selection and store new selection
    '        DS.ObjSelections.Clear()

    '        For Each sel As clsDSSelection In objThis.ObjSelections
    '            Dim newSel As clsDSSelection = sel.Clone(DS)
    '            'For Each fld As clsField In sel.DSSelectionFields
    '            '    Dim newFld As clsField = fld.Clone(newSel)
    '            '    newSel.DSSelectionFields.Add(newFld)
    '            'Next
    '            DS.ObjSelections.Add(newSel)
    '        Next




    '        ''//now save structure selection
    '        'Dim fNode, pNode, cNode As TreeNode
    '        'Dim tempObj As New clsDSSelection
    '        'Dim tempsel As New clsDSSelection
    '        'Dim i As Integer = 0
    '        'Dim flag As Boolean = False

    '        ''//for each folder node
    '        'For Each fNode In tvDatastoreStructures.Nodes
    '        '    '//each structure under folder
    '        '    For Each pNode In fNode.Nodes
    '        '        '//check for structures if selected then we dont consider 
    '        '        '//child selection otherwise scan for children selection
    '        '        '// ALSO, any selection that is checked needs to be "cloned"
    '        '        '// into a Datastore Selection instead of a Structure Selection
    '        '        '// so it is "independent" of the Structure and becomes part
    '        '        '// of the Datastore
    '        '        If pNode.Checked = True Then
    '        '            tempsel = DS.CloneSSeltoDSSel(pNode.Tag)
    '        '            For i = 0 To DS.OldObjSelections.Count - 1
    '        '                tempObj = DS.OldObjSelections(i)
    '        '                tempObj.LoadItems()
    '        '                '// if the DSSelection exists already in old DSselections, 
    '        '                '//then add it to current DS selections
    '        '                If tempObj.Text Is tempsel.Text Then
    '        '                    DS.ObjSelections.Add(tempObj)
    '        '                    flag = True
    '        '                    Exit For
    '        '                End If
    '        '            Next
    '        '            If flag = False Then
    '        '                DS.ObjSelections.Add(tempsel)
    '        '            End If
    '        '            flag = False
    '        '        Else
    '        '            '//each selection Of Selection
    '        '            For Each cNode In pNode.Nodes
    '        '                If cNode.Checked = True Then
    '        '                    tempsel = DS.CloneSSeltoDSSel(cNode.Tag)
    '        '                    For i = 0 To DS.OldObjSelections.Count - 1
    '        '                        tempObj = DS.OldObjSelections(i)
    '        '                        tempObj.LoadItems()
    '        '                        If tempObj.Text Is tempsel.Text Then
    '        '                            DS.ObjSelections.Add(tempObj)
    '        '                            flag = True
    '        '                            Exit For
    '        '                        End If
    '        '                    Next
    '        '                    If flag = False Then
    '        '                        DS.ObjSelections.Add(tempsel)
    '        '                    End If
    '        '                    flag = False
    '        '                End If
    '        '            Next
    '        '        End If
    '        '    Next
    '        'Next

    '        '/// once all datastore selections are saved from form, make sure all parents are set
    '        DS.SetDSselParents()

    '        Return DS

    '    Catch ex As Exception
    '        LogError(ex, "ctlDatastore UpdateChildDS")
    '        Return Nothing
    '    End Try

    'End Function

    '//write form controls values to object
    Function SaveCurrentSelection() As Boolean

        Try
            'IsEventFromCode = True

            objThis.DatastoreDescription = txtDatastoreDesc.Text

            If objThis.DsDirection <> "" Then
                If objThis.DsDirection = DS_DIRECTION_TARGET Then
                    If cmbOperationType.SelectedItem IsNot Nothing Then
                        objThis.OperationType = GetStr(GetVal(CType(cmbOperationType.SelectedItem, Mylist).ItemData))
                    End If
                End If
            End If

            If cmbCharacterCode.Enabled = True Then
                If CType(cmbCharacterCode.SelectedItem, Mylist) IsNot Nothing Then
                    objThis.DsCharacterCode = CType(cmbCharacterCode.SelectedItem, Mylist).ItemData
                Else
                    objThis.DsCharacterCode = ""
                End If
            Else
                objThis.DsCharacterCode = ""
            End If

            '//npatel 10/7/05
            If cmbAccessMethod.Enabled = True Then
                Dim Itm As Mylist
                Itm = cmbAccessMethod.SelectedItem
                If (Itm Is Nothing) = False Then
                    If Itm.ItemData = DS_ACCESSMETHOD_IP Or Itm.ItemData = DS_ACCESSMETHOD_CDCSTORE Then
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

            objThis.DsUOW = txtUOW.Text   '/// Save UOW
            objThis.DsPhysicalSource = txtPhysicalSource.Text  '/// Save Physical Name
            objThis.ExceptionDatastore = txtException.Text   '/// Save Exception Datastore
            'objThis.Restart = txtRestart.Text   '/// Save Restart Que
            'objThis.Poll = txtPoll.Text  '// Save Poll Number

            '/// Added 7/2012  by TK
            objThis.UseMTD = cbUseFile.Checked
            objThis.DsHostName = txtHostName.Text.Trim
            objThis.DsHostName = txtHostName.Text.Trim


            'objThis.IsCmmtKey = IIf(chkCmmtKey.Checked, "1", "0") '// added by TK and KS 11/6/2006
            objThis.IsIMSPathData = IIf(chkIMSPathData.Checked, "1", "0")
            'objThis.IsOrdered = IIf(chkOrdered.Checked, "1", "0")
            objThis.IsSkipChangeCheck = IIf(chkSkipChangeCheck.Checked, "1", "0")
            'objThis.IsBeforeImage = IIf(chkBeforeImage.Checked, "1", "0") '// new by npatel on 8/10/05
            objThis.IsKeyChng = IIf(cbKeyChng.Checked, "1", "0")

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
            'IsEventFromCode = False

            If ValidateSelection() = False Then
                SaveCurrentSelection = True
                Exit Function
            End If
            '//before we clear old selection save it to another array so later on while saving this object we can compare for any new addition in selection list
            objThis.OldObjSelections.Clear()
            objThis.OldObjSelections = objThis.ObjSelections.Clone

            '//clear previous selection and store new selection
            objThis.ObjSelections.Clear()


            '//now save structure selection
            Dim fNode, pNode, cNode As TreeNode
            Dim tempObj As New clsDSSelection
            Dim tempsel As New clsDSSelection
            Dim i As Integer = 0
            Dim flag As Boolean = False

            '//for each folder node
            For Each fNode In tvDatastoreStructures.Nodes
                '//each structure under folder
                For Each pNode In fNode.Nodes
                    '//check for structures if selected then we dont consider 
                    '//child selection otherwise scan for children selection
                    '// ALSO, any selection that is checked needs to be "cloned"
                    '// into a Datastore Selection instead of a Structure Selection
                    '// so it is "independent" of the Structure and becomes part
                    '// of the Datastore
                    If pNode.Checked = True Then
                        tempsel = objThis.CloneSSeltoDSSel(pNode.Tag)
                        For i = 0 To objThis.OldObjSelections.Count - 1
                            tempObj = objThis.OldObjSelections(i)
                            'tempObj.LoadItems()
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
                                    'tempObj.LoadItems()
                                    If tempObj.SelectionName Is tempsel.SelectionName Then
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
            LogError(ex, "ctlDatastore SaveCurrentSelection")
            Return False
        End Try

    End Function

    Private Sub ShowHideFieldAttr()

        If IsNewObj = True Then
            If ValidateSelection() = False Then Exit Sub
        End If

        '//Dont reload everytime once its loaded from database

        tvFields.BackColor = Color.LightBlue
        lvFieldAttrs.BackColor = Color.LightBlue

        'lblFieldName.ForeColor = Color.LightYellow
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

    Function ValidateSelection() As Boolean

        If tvDatastoreStructures.GetNodeCount(True) <= 0 Then
            ValidateSelection = False
            Exit Function
        End If
        ValidateSelection = True

    End Function

    '//This will load treeview of fields
    '//If New Datastore then fields will be from XML file
    '//If Edit Datastore then fields wil be from database
    Function FillFieldAttr() As Boolean

        Dim objSel As clsStructureSelection
        Dim dsSel As clsDSSelection
        Dim isfound As Boolean = False
        Dim i As Integer

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
                'dsSel.LoadItems()
                If objSel Is dsSel.ObjSelection Then
                    isfound = True
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

        If cNode Is Nothing Then
            Exit Function
        End If

        lvFieldAttrs.BeginUpdate()

        lvFieldAttrs.GridLines = True
        lvFieldAttrs.Items.Clear()

        lvFieldAttrs.Items.Add(modDeclares.TXT_ISKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
        lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY)) '// added by TK and KS 11/6/2006
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

            If i < 9 Then
                lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Black
                lvFieldAttrs.Items(i).ImageIndex = 25
            Else
                lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).SubItems(1).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).ImageIndex = 30
            End If

        Next

        Dim startcolor As Color

        'lblFieldName.Text = cNode.Text
        'lblFieldNameDesc.Text = lblFieldName.Text
        prevfld = CType(cNode.Tag, clsField)
        IsEventFromCode = True
        txtFieldDesc.Text = prevfld.FieldDesc
        startcolor = cNode.ForeColor
        If prevfld.FieldDesc <> "" And startcolor <> Color.Red Then
            cNode.ForeColor = Color.Blue
        Else
            cNode.ForeColor = startcolor
        End If
        IsEventFromCode = False

        lvFieldAttrs.EndUpdate()

    End Function

    Function FillStructure(ByVal cNode As TreeNode, Optional ByVal nd As TreeNode = Nothing) As Boolean

        Try
            Dim cStructFolder As TreeNode
            Dim activenode As INode = Nothing
            Dim NodeType As String = CType(cNode.Tag, INode).Type

            If nd IsNot Nothing Then '//cNode is representing DataStore Selection
                activenode = CType(nd.Tag, clsDSSelection)
            End If

            Select Case NodeType
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    cNode = CType(cNode.Tag, clsDatastore).Environment.ObjTreeNode
                Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                    cNode = CType(cNode.Tag, clsDSSelection).ObjDatastore.Environment.ObjTreeNode
                Case NODE_FO_SOURCEDATASTORE, NODE_FO_TARGETDATASTORE
                    cNode = CType(cNode.Tag, clsFolderNode).Parent.ObjTreeNode
                    cNode = CType(cNode.Tag, clsEngine).ObjSystem.Environment.ObjTreeNode
                Case Else
                    Return False
                    Exit Function
            End Select

            cStructFolder = cNode.Nodes(1)

            AddChildrenNoFolder(cStructFolder.Clone, tvDatastoreStructures.Nodes)
            tvDatastoreStructures.CollapseAll()


            '//Now load selected fields and check if object is already created and user is editing it
            If IsNewObj = False Then
                '// This is to make sure that changes to Procs, etc. are up to date.
                objThis.LoadItems() 'True
                CheckSelected()
                If objThis.Engine IsNot Nothing Then
                    objThis.SetIsMapped()
                End If

                If Not activenode Is Nothing Then
                    tvDatastoreStructures.SelectedNode = SelectFirstMatchingNode(tvDatastoreStructures, activenode.Text, , , True)
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlDatastore FillStructure")
            Return False
        End Try

    End Function

    Sub AddChildrenNoFolder(ByVal cNodeSource As TreeNode, ByVal cNodeTarget As TreeNodeCollection)

        Try
            Dim nd As TreeNode
            Dim newNode As TreeNode
            Dim obj As INode

            For Each nd In cNodeSource.Nodes
                obj = CType(nd.Tag, INode)
                '//now only add applicable entries in treeview
                '//e.g. for Binary DS all structure but for XML DS only DTD is valid 
                '//structure so only add DTD structures and subselection
                '//If not valid structure or selection then no need to find their 
                '//children coz they will be the same type
                If IsValidStructureForDS(obj) = True Then
                    newNode = cNodeTarget.Add(nd.Text)
                    If obj.Type = NODE_STRUCT = True Then
                        newNode.Tag = CType(obj, clsStructure).SysAllSelection
                    Else
                        newNode.Tag = nd.Tag
                    End If
                    newNode.ImageIndex = nd.ImageIndex
                    newNode.SelectedImageIndex = nd.SelectedImageIndex

                    AddChildrenNoFolder(nd, newNode.Nodes)
                End If
            Next

        Catch ex As Exception
            LogError(ex, "ctlDatastore AddChildrenNoFolder")
        End Try

    End Sub

    Function CheckSelected() As Boolean

        Dim i As Integer
        Dim cnt As Integer = tvDatastoreStructures.GetNodeCount(True)
        Dim dsSel As clsDSSelection

        If cnt <= 0 Then Exit Function
        Try
            tvDatastoreStructures.BeginUpdate()
            '//Looop all selected DSSelections and check that Selection in treeview
            For i = 0 To objThis.ObjSelections.Count - 1
                dsSel = CType(objThis.ObjSelections(i), clsDSSelection)
                SelectFirstMatchingNode(tvDatastoreStructures, dsSel.Text, True, dsSel.ObjSelection.Key, True, False)
            Next

        Catch ex As Exception
            LogError(ex, "ctlDatastore CheckSelected")
        Finally
            tvDatastoreStructures.EndUpdate()
        End Try

    End Function

    '//This function check structure type and selected datastore type and if 
    '//applicable structure then returns true so we can add into the treeview
    Private Function IsValidStructureForDS(ByVal obj As INode) As Boolean

        IsValidStructureForDS = True

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

            MsgBox("At least one Key Field must be set in any Relational, IMSDB or VSAM Target Datastore." & Chr(13) & _
            "No key fields have been set in this Datastore. Please set at least one key field.", _
            MsgBoxStyle.Exclamation, "Please Set a key field")

            HasKeyForRel = False

        Catch ex As Exception
            LogError(ex, "ctlDatastore HasKeyForRel")
            Return False
        End Try

    End Function

#End Region

    Private Sub txtFieldDesc_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFieldDesc.TextChanged

        Try
            If IsEventFromCode = True Then Exit Sub
            objThis.IsModified = True
            objThis.IsLoaded = False
            cmdSave.Enabled = True
            If prevfld IsNot Nothing Then
                prevfld.FieldDescModified = True
            End If

            RaiseEvent Modified(Me, objThis)

        Catch ex As Exception
            LogError(ex, "ctlDatastore txtFieldDesc_TextChanged")
        End Try


    End Sub

    Private Sub txtFieldDesc_lostfocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFieldDesc.LostFocus, txtFieldDesc.MouseLeave, txtFieldDesc.Leave

        If prevfld IsNot Nothing Then
            prevfld.FieldDesc = txtFieldDesc.Text
        End If

    End Sub

    Private Sub txtMTD_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMTD.TextChanged

        Try
            objThis.MTDfile = txtMTD.Text.Trim

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore txtMTD_TextChanged")
        End Try

    End Sub

    Private Sub btnMTD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMTD.Click

        Try
            Dim strFilter As String = "SQL Description File(*.sql)|*.sql|DDL Description File(*.ddl)|*.ddl|All files (*.*)|*.*"

            dlgOpen.InitialDirectory = objThis.Environment.LocalDDLDir

            dlgOpen.Filter = strFilter
            If modGeneral.GetExtensionFromFilePath(txtMTD.Text.ToUpper) = "SQL" Then
                dlgOpen.FilterIndex = 0
            ElseIf modGeneral.GetExtensionFromFilePath(txtMTD.Text.ToUpper) = "DDL" Then
                dlgOpen.FilterIndex = 1
            Else
                dlgOpen.FilterIndex = 2
            End If

            dlgOpen.FileName = txtMTD.Text


            If dlgOpen.ShowDialog() = DialogResult.OK Then
                txtMTD.Text = dlgOpen.FileName
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore btnMTD_Click")
        End Try

    End Sub

    Private Sub txtHostName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtHostName.TextChanged

        Try
            If System.IO.Directory.Exists(txtHostName.Text.Trim) = True Then
                If SaveCurrentSelection() = True Then

                End If
                objThis.DsHostName = txtHostName.Text.Trim
            End If

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore txtHostName_TextChanged")
        End Try

    End Sub

    Private Sub cbUseFile_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbUseFile.CheckedChanged

        Try
            If cbUseFile.Checked = True Then
                objThis.UseMTD = True
            Else
                objThis.UseMTD = False
            End If

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlDatastore cbUseFile_CheckedChanged")
        End Try

    End Sub

    Private Sub btnSelDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelDesc.Click

        Try
            Dim frm As frmDS_MTDdescriptions    '/// MTDform
            Dim col As New Collection

            frm = New frmDS_MTDdescriptions

            Me.SaveCurrentSelection()

            col = frm.EditObj(objThis)

            If col IsNot Nothing Then
                objThis.DescList = col
                OnChange(Me, New EventArgs)
            End If

        Catch ex As Exception
            LogError(ex, "ctlDatastore btnSelDesc_Click")
        End Try

    End Sub

End Class
