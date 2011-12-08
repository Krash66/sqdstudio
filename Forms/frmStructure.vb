Public Class frmStructure
    Inherits SQDStudio.frmBlank

    Public objThis As New clsStructure
    Public objEnv As New clsEnvironment
    Public arrObjThis As New ArrayList '//new 10/22/05
    Private ObjToChg As clsStructure  '//added to be able to replace structure source file 

    Dim IsNewObj As Boolean
    Dim IsChangeObj As Boolean = False
    Dim lastTreeviewSearchText As String
    Dim lastTreeviewSearchNode As TreeNode
    Dim colSkipNodes As New ArrayList
    Dim GottenFiles As Collection
    Dim FileNames As Collection
    Friend WithEvents btnGetXML As System.Windows.Forms.Button

    Dim DMLFile As String = ""

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

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
    Friend WithEvents cmdShowHideFieldAttr As System.Windows.Forms.Button
    Friend WithEvents pnlFields As System.Windows.Forms.Panel
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlSelect As System.Windows.Forms.Panel
    Friend WithEvents chkAddDBD As System.Windows.Forms.CheckBox
    Friend WithEvents cmdBrowseFile As System.Windows.Forms.Button
    Friend WithEvents txtFilePath As System.Windows.Forms.TextBox
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents cmdBrowseDBDFile As System.Windows.Forms.Button
    Friend WithEvents txtDBDFilePath As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseCobolFile As System.Windows.Forms.Button
    Friend WithEvents txtCobolFilePath As System.Windows.Forms.TextBox
    Friend WithEvents lblCobolFile As System.Windows.Forms.Label
    Friend WithEvents lblDBD As System.Windows.Forms.Label
    Friend WithEvents txtStructName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents tvSegments As System.Windows.Forms.TreeView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtSearchField As System.Windows.Forms.TextBox
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents lblFieldName As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents txtStructDesc As System.Windows.Forms.TextBox
    Friend WithEvents OpenMultiFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents chkFTP As System.Windows.Forms.CheckBox
    Friend WithEvents gbDMLConn As System.Windows.Forms.GroupBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents cbConn As System.Windows.Forms.ComboBox
    Friend WithEvents gbCOBOL As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbProp As System.Windows.Forms.GroupBox
    Friend WithEvents gbSeg As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.cmdShowHideFieldAttr = New System.Windows.Forms.Button
        Me.pnlFields = New System.Windows.Forms.Panel
        Me.lblFieldName = New System.Windows.Forms.Label
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.txtSearchField = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.pnlSelect = New System.Windows.Forms.Panel
        Me.gbProp = New System.Windows.Forms.GroupBox
        Me.btnGetXML = New System.Windows.Forms.Button
        Me.gbDMLConn = New System.Windows.Forms.GroupBox
        Me.cbConn = New System.Windows.Forms.ComboBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtStructName = New System.Windows.Forms.TextBox
        Me.gbCOBOL = New System.Windows.Forms.GroupBox
        Me.txtCobolFilePath = New System.Windows.Forms.TextBox
        Me.cmdBrowseCobolFile = New System.Windows.Forms.Button
        Me.lblCobolFile = New System.Windows.Forms.Label
        Me.lblDBD = New System.Windows.Forms.Label
        Me.txtDBDFilePath = New System.Windows.Forms.TextBox
        Me.chkAddDBD = New System.Windows.Forms.CheckBox
        Me.cmdBrowseDBDFile = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.lblFile = New System.Windows.Forms.Label
        Me.txtFilePath = New System.Windows.Forms.TextBox
        Me.cmdBrowseFile = New System.Windows.Forms.Button
        Me.chkFTP = New System.Windows.Forms.CheckBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.txtStructDesc = New System.Windows.Forms.TextBox
        Me.gbSeg = New System.Windows.Forms.GroupBox
        Me.tvSegments = New System.Windows.Forms.TreeView
        Me.OpenMultiFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Panel1.SuspendLayout()
        Me.pnlFields.SuspendLayout()
        Me.pnlSelect.SuspendLayout()
        Me.gbProp.SuspendLayout()
        Me.gbDMLConn.SuspendLayout()
        Me.gbCOBOL.SuspendLayout()
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbSeg.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(607, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(0, 394)
        Me.GroupBox1.Size = New System.Drawing.Size(605, 10)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(340, 412)
        Me.cmdOk.TabIndex = 13
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(426, 412)
        Me.cmdCancel.TabIndex = 14
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(512, 412)
        Me.cmdHelp.TabIndex = 15
        '
        'Label1
        '
        Me.Label1.Text = "Description definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(517, 39)
        Me.Label2.Text = "Enter a unique Description name, and specify the file containing the Description " & _
            "definition."
        '
        'cmdShowHideFieldAttr
        '
        Me.cmdShowHideFieldAttr.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdShowHideFieldAttr.Location = New System.Drawing.Point(51, 412)
        Me.cmdShowHideFieldAttr.Name = "cmdShowHideFieldAttr"
        Me.cmdShowHideFieldAttr.Size = New System.Drawing.Size(152, 24)
        Me.cmdShowHideFieldAttr.TabIndex = 12
        Me.cmdShowHideFieldAttr.Text = "Button1"
        '
        'pnlFields
        '
        Me.pnlFields.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlFields.Controls.Add(Me.lblFieldName)
        Me.pnlFields.Controls.Add(Me.cmdSearch)
        Me.pnlFields.Controls.Add(Me.txtSearchField)
        Me.pnlFields.Controls.Add(Me.Label5)
        Me.pnlFields.Controls.Add(Me.Label4)
        Me.pnlFields.Controls.Add(Me.tvFields)
        Me.pnlFields.Controls.Add(Me.lvFieldAttrs)
        Me.pnlFields.Location = New System.Drawing.Point(49, 71)
        Me.pnlFields.Name = "pnlFields"
        Me.pnlFields.Padding = New System.Windows.Forms.Padding(5)
        Me.pnlFields.Size = New System.Drawing.Size(596, 318)
        Me.pnlFields.TabIndex = 32
        '
        'lblFieldName
        '
        Me.lblFieldName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldName.ForeColor = System.Drawing.Color.Blue
        Me.lblFieldName.Location = New System.Drawing.Point(313, 9)
        Me.lblFieldName.Name = "lblFieldName"
        Me.lblFieldName.Size = New System.Drawing.Size(170, 16)
        Me.lblFieldName.TabIndex = 38
        '
        'cmdSearch
        '
        Me.cmdSearch.Location = New System.Drawing.Point(197, 5)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(30, 20)
        Me.cmdSearch.TabIndex = 37
        Me.cmdSearch.Text = "Go"
        '
        'txtSearchField
        '
        Me.txtSearchField.Location = New System.Drawing.Point(55, 6)
        Me.txtSearchField.MaxLength = 128
        Me.txtSearchField.Name = "txtSearchField"
        Me.txtSearchField.Size = New System.Drawing.Size(136, 20)
        Me.txtSearchField.TabIndex = 36
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 9)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 12)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Search"
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(233, 9)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 12)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Attributes"
        '
        'tvFields
        '
        Me.tvFields.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tvFields.Location = New System.Drawing.Point(8, 32)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.Size = New System.Drawing.Size(215, 278)
        Me.tvFields.TabIndex = 34
        '
        'lvFieldAttrs
        '
        Me.lvFieldAttrs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvFieldAttrs.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.lvFieldAttrs.FullRowSelect = True
        Me.lvFieldAttrs.GridLines = True
        Me.lvFieldAttrs.Location = New System.Drawing.Point(227, 32)
        Me.lvFieldAttrs.Name = "lvFieldAttrs"
        Me.lvFieldAttrs.Scrollable = CType(configurationAppSettings.GetValue("lvFieldAttrs.Scrollable", GetType(Boolean)), Boolean)
        Me.lvFieldAttrs.Size = New System.Drawing.Size(363, 278)
        Me.lvFieldAttrs.TabIndex = 33
        Me.lvFieldAttrs.UseCompatibleStateImageBehavior = False
        Me.lvFieldAttrs.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Attribute"
        Me.ColumnHeader1.Width = 100
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Value"
        Me.ColumnHeader2.Width = 140
        '
        'pnlSelect
        '
        Me.pnlSelect.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlSelect.Controls.Add(Me.gbProp)
        Me.pnlSelect.Controls.Add(Me.gbSeg)
        Me.pnlSelect.Location = New System.Drawing.Point(3, 74)
        Me.pnlSelect.Name = "pnlSelect"
        Me.pnlSelect.Size = New System.Drawing.Size(596, 314)
        Me.pnlSelect.TabIndex = 33
        '
        'gbProp
        '
        Me.gbProp.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbProp.Controls.Add(Me.btnGetXML)
        Me.gbProp.Controls.Add(Me.gbDMLConn)
        Me.gbProp.Controls.Add(Me.Label3)
        Me.gbProp.Controls.Add(Me.Label7)
        Me.gbProp.Controls.Add(Me.txtStructName)
        Me.gbProp.Controls.Add(Me.gbCOBOL)
        Me.gbProp.Controls.Add(Me.Label6)
        Me.gbProp.Controls.Add(Me.gbName)
        Me.gbProp.Controls.Add(Me.chkFTP)
        Me.gbProp.Controls.Add(Me.gbDesc)
        Me.gbProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProp.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbProp.Location = New System.Drawing.Point(3, 3)
        Me.gbProp.Name = "gbProp"
        Me.gbProp.Size = New System.Drawing.Size(414, 307)
        Me.gbProp.TabIndex = 77
        Me.gbProp.TabStop = False
        Me.gbProp.Text = "Description Properties"
        '
        'btnGetXML
        '
        Me.btnGetXML.Location = New System.Drawing.Point(266, 45)
        Me.btnGetXML.Name = "btnGetXML"
        Me.btnGetXML.Size = New System.Drawing.Size(142, 23)
        Me.btnGetXML.TabIndex = 76
        Me.btnGetXML.Text = "Get XML Message"
        Me.btnGetXML.UseVisualStyleBackColor = True
        '
        'gbDMLConn
        '
        Me.gbDMLConn.Controls.Add(Me.cbConn)
        Me.gbDMLConn.Controls.Add(Me.Label11)
        Me.gbDMLConn.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDMLConn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDMLConn.Location = New System.Drawing.Point(5, 125)
        Me.gbDMLConn.Name = "gbDMLConn"
        Me.gbDMLConn.Size = New System.Drawing.Size(330, 52)
        Me.gbDMLConn.TabIndex = 73
        Me.gbDMLConn.TabStop = False
        Me.gbDMLConn.Text = "Select a Connection for DML"
        '
        'cbConn
        '
        Me.cbConn.FormattingEnabled = True
        Me.cbConn.Location = New System.Drawing.Point(118, 20)
        Me.cbConn.Name = "cbConn"
        Me.cbConn.Size = New System.Drawing.Size(206, 21)
        Me.cbConn.TabIndex = 3
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(5, 23)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(107, 13)
        Me.Label11.TabIndex = 2
        Me.Label11.Text = "Connection Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 55
        Me.Label3.Text = "Name"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Maroon
        Me.Label7.Location = New System.Drawing.Point(48, 64)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(38, 14)
        Me.Label7.TabIndex = 70
        Me.Label7.Text = "Note :"
        '
        'txtStructName
        '
        Me.txtStructName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStructName.Location = New System.Drawing.Point(50, 19)
        Me.txtStructName.MaxLength = 128
        Me.txtStructName.Name = "txtStructName"
        Me.txtStructName.Size = New System.Drawing.Size(358, 20)
        Me.txtStructName.TabIndex = 1
        '
        'gbCOBOL
        '
        Me.gbCOBOL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbCOBOL.Controls.Add(Me.txtCobolFilePath)
        Me.gbCOBOL.Controls.Add(Me.cmdBrowseCobolFile)
        Me.gbCOBOL.Controls.Add(Me.lblCobolFile)
        Me.gbCOBOL.Controls.Add(Me.lblDBD)
        Me.gbCOBOL.Controls.Add(Me.txtDBDFilePath)
        Me.gbCOBOL.Controls.Add(Me.chkAddDBD)
        Me.gbCOBOL.Controls.Add(Me.cmdBrowseDBDFile)
        Me.gbCOBOL.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbCOBOL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbCOBOL.Location = New System.Drawing.Point(6, 125)
        Me.gbCOBOL.Name = "gbCOBOL"
        Me.gbCOBOL.Size = New System.Drawing.Size(402, 95)
        Me.gbCOBOL.TabIndex = 74
        Me.gbCOBOL.TabStop = False
        Me.gbCOBOL.Text = "COBOL Description Files"
        '
        'txtCobolFilePath
        '
        Me.txtCobolFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCobolFilePath.Location = New System.Drawing.Point(85, 67)
        Me.txtCobolFilePath.MaxLength = 255
        Me.txtCobolFilePath.Name = "txtCobolFilePath"
        Me.txtCobolFilePath.Size = New System.Drawing.Size(280, 20)
        Me.txtCobolFilePath.TabIndex = 6
        '
        'cmdBrowseCobolFile
        '
        Me.cmdBrowseCobolFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseCobolFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseCobolFile.Location = New System.Drawing.Point(371, 66)
        Me.cmdBrowseCobolFile.Name = "cmdBrowseCobolFile"
        Me.cmdBrowseCobolFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseCobolFile.TabIndex = 7
        Me.cmdBrowseCobolFile.Text = "..."
        '
        'lblCobolFile
        '
        Me.lblCobolFile.AutoSize = True
        Me.lblCobolFile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCobolFile.Location = New System.Drawing.Point(6, 70)
        Me.lblCobolFile.Name = "lblCobolFile"
        Me.lblCobolFile.Size = New System.Drawing.Size(26, 14)
        Me.lblCobolFile.TabIndex = 58
        Me.lblCobolFile.Text = "File"
        '
        'lblDBD
        '
        Me.lblDBD.AutoSize = True
        Me.lblDBD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDBD.Location = New System.Drawing.Point(6, 44)
        Me.lblDBD.Name = "lblDBD"
        Me.lblDBD.Size = New System.Drawing.Size(73, 14)
        Me.lblDBD.TabIndex = 57
        Me.lblDBD.Text = "IMS DBD File"
        '
        'txtDBDFilePath
        '
        Me.txtDBDFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDBDFilePath.Location = New System.Drawing.Point(85, 41)
        Me.txtDBDFilePath.MaxLength = 255
        Me.txtDBDFilePath.Name = "txtDBDFilePath"
        Me.txtDBDFilePath.Size = New System.Drawing.Size(280, 20)
        Me.txtDBDFilePath.TabIndex = 8
        '
        'chkAddDBD
        '
        Me.chkAddDBD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkAddDBD.Location = New System.Drawing.Point(85, 19)
        Me.chkAddDBD.Name = "chkAddDBD"
        Me.chkAddDBD.Size = New System.Drawing.Size(114, 16)
        Me.chkAddDBD.TabIndex = 2
        Me.chkAddDBD.Text = "Add DBD Fields"
        '
        'cmdBrowseDBDFile
        '
        Me.cmdBrowseDBDFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseDBDFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseDBDFile.Location = New System.Drawing.Point(371, 40)
        Me.cmdBrowseDBDFile.Name = "cmdBrowseDBDFile"
        Me.cmdBrowseDBDFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseDBDFile.TabIndex = 9
        Me.cmdBrowseDBDFile.Text = "..."
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Blue
        Me.Label6.Location = New System.Drawing.Point(92, 64)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(288, 45)
        Me.Label6.TabIndex = 69
        Me.Label6.Text = "If structure represents an IMS segment, check the ""Add DBD Source"", specify the I" & _
            "MS DBD file name, and highlight a segment name in the ""Segments"" tree."
        '
        'gbName
        '
        Me.gbName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbName.Controls.Add(Me.lblFile)
        Me.gbName.Controls.Add(Me.txtFilePath)
        Me.gbName.Controls.Add(Me.cmdBrowseFile)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbName.Location = New System.Drawing.Point(6, 67)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(402, 51)
        Me.gbName.TabIndex = 75
        Me.gbName.TabStop = False
        Me.gbName.Text = "Description File"
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
        'txtFilePath
        '
        Me.txtFilePath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFilePath.Location = New System.Drawing.Point(38, 19)
        Me.txtFilePath.MaxLength = 255
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.Size = New System.Drawing.Size(327, 20)
        Me.txtFilePath.TabIndex = 4
        '
        'cmdBrowseFile
        '
        Me.cmdBrowseFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFile.Location = New System.Drawing.Point(371, 18)
        Me.cmdBrowseFile.Name = "cmdBrowseFile"
        Me.cmdBrowseFile.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFile.TabIndex = 5
        Me.cmdBrowseFile.Text = "..."
        '
        'chkFTP
        '
        Me.chkFTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFTP.Location = New System.Drawing.Point(50, 45)
        Me.chkFTP.Name = "chkFTP"
        Me.chkFTP.Size = New System.Drawing.Size(134, 16)
        Me.chkFTP.TabIndex = 3
        Me.chkFTP.Text = "Browse using FTP"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtStructDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbDesc.Location = New System.Drawing.Point(6, 226)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(402, 75)
        Me.gbDesc.TabIndex = 73
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'txtStructDesc
        '
        Me.txtStructDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtStructDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStructDesc.Location = New System.Drawing.Point(6, 19)
        Me.txtStructDesc.MaxLength = 255
        Me.txtStructDesc.Multiline = True
        Me.txtStructDesc.Name = "txtStructDesc"
        Me.txtStructDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtStructDesc.Size = New System.Drawing.Size(390, 50)
        Me.txtStructDesc.TabIndex = 10
        '
        'gbSeg
        '
        Me.gbSeg.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbSeg.Controls.Add(Me.tvSegments)
        Me.gbSeg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbSeg.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbSeg.Location = New System.Drawing.Point(423, 3)
        Me.gbSeg.Name = "gbSeg"
        Me.gbSeg.Size = New System.Drawing.Size(168, 307)
        Me.gbSeg.TabIndex = 76
        Me.gbSeg.TabStop = False
        Me.gbSeg.Text = "Selected Segments"
        '
        'tvSegments
        '
        Me.tvSegments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvSegments.Location = New System.Drawing.Point(3, 16)
        Me.tvSegments.Name = "tvSegments"
        Me.tvSegments.Size = New System.Drawing.Size(162, 288)
        Me.tvSegments.TabIndex = 11
        '
        'OpenMultiFileDialog1
        '
        Me.OpenMultiFileDialog1.ShowReadOnly = True
        Me.OpenMultiFileDialog1.SupportMultiDottedExtensions = True
        '
        'frmStructure
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.CancelButton = Nothing
        Me.ClientSize = New System.Drawing.Size(604, 448)
        Me.Controls.Add(Me.pnlSelect)
        Me.Controls.Add(Me.cmdShowHideFieldAttr)
        Me.Controls.Add(Me.pnlFields)
        Me.MinimumSize = New System.Drawing.Size(550, 0)
        Me.Name = "frmStructure"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Description Properties"
        Me.Controls.SetChildIndex(Me.pnlFields, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdShowHideFieldAttr, 0)
        Me.Controls.SetChildIndex(Me.pnlSelect, 0)
        Me.Panel1.ResumeLayout(False)
        Me.pnlFields.ResumeLayout(False)
        Me.pnlFields.PerformLayout()
        Me.pnlSelect.ResumeLayout(False)
        Me.gbProp.ResumeLayout(False)
        Me.gbProp.PerformLayout()
        Me.gbDMLConn.ResumeLayout(False)
        Me.gbDMLConn.PerformLayout()
        Me.gbCOBOL.ResumeLayout(False)
        Me.gbCOBOL.PerformLayout()
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbSeg.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddDBD.CheckedChanged, txtCobolFilePath.TextChanged, txtStructName.TextChanged, txtDBDFilePath.TextChanged, txtStructDesc.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        txtStructName.Select()

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False
        '//added so buttons show correctly on load
        cmdShowHideFieldAttr.Text = "Show Fields/Attributes"

    End Sub

    Friend Function NewObj(Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN, Optional ByVal DMLPath As String = "", Optional ByVal Env As clsEnvironment = Nothing) As ArrayList

        IsNewObj = True

        StartLoad()

        DMLFile = DMLPath

        InitControls()

        If DMLFile <> "" Then
            txtFilePath.Enabled = False
            cmdBrowseFile.Enabled = False
            cmdOk.Enabled = True
        End If

        ShowHideControls(StructType)
        objThis.StructureType = StructType

        If objThis.StructureType = enumStructure.STRUCT_REL_DML Or objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
            If Env IsNot Nothing Then
                objThis.Environment = Env
                SetComboConn(Env)
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
                GoTo doAgain
            End If
        End If
        txtStructName.SelectAll()
        txtStructName.Focus()

        If DMLFile <> "" Then
            txtStructName.Text = GetFileNameWithoutExtenstionFromPath(DMLFile)
            Dim SubStrArr As String()
            Dim SepChar As Char() = "."
            SubStrArr = DMLFile.Split(SepChar)
            txtFilePath.Text = SubStrArr(1) & "." & SubStrArr(2)
            Dim ConnName As String = SubStrArr(0)
            For Each conn As clsConnection In Env.Connections
                If conn.Text = ConnName Then
                    objThis.Connection = conn
                    objThis.Connection.LoadMe()
                    Exit For
                End If
            Next
            DMLFile = objThis.Connection.UserId & "." & _
            objThis.Connection.Password & "." & _
            objThis.Connection.Database & "." & txtFilePath.Text
        End If

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = arrObjThis
            Case Windows.Forms.DialogResult.Retry
                tvSegments.ExpandAll()
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Function ChangeObj(ByRef OldObj As clsStructure) As clsStructure

        Dim StructType As enumStructure

        IsNewObj = False
        IsChangeObj = True
        StartLoad()
        InitControls()

        ObjToChg = OldObj

        StructType = ObjToChg.StructureType

        If ObjToChg.fPath2 = "" Then
            cmdBrowseDBDFile.Enabled = False
        End If

        ShowHideControls(StructType)

        If StructType = enumStructure.STRUCT_REL_DML Or StructType = enumStructure.STRUCT_REL_DML_FILE Then
            If ObjToChg.Environment IsNot Nothing Then
                SetComboConn(ObjToChg.Environment)
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
                GoTo doAgain
            End If
        End If

        If StructType = enumStructure.STRUCT_COBOL Or StructType = enumStructure.STRUCT_COBOL_IMS Then
            txtCobolFilePath.Text = ObjToChg.fPath1
            txtDBDFilePath.Text = ObjToChg.fPath2
            txtCobolFilePath.SelectAll()
        Else
            txtFilePath.Text = ObjToChg.fPath1
            txtFilePath.SelectAll()
        End If

        txtStructName.ReadOnly = True
        txtStructName.BackColor = Color.LightBlue
        txtStructName.ForeColor = Color.DarkBlue
        txtStructName.Text = ObjToChg.StructureName

        InitChangeObj() '// puts oldObj properties into new objthis (almost like clone)

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                '// This means the old structure object was successfully removed and the new 
                '// object was successfully inserted, so we return the new object THEN
                '// Reload the entire project so all object relationships are reset
                '// based on reading metadata again with the new object values.
                ChangeObj = objThis
            Case Windows.Forms.DialogResult.Retry
                tvSegments.ExpandAll()
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Function InitChangeObj() As Boolean

        objThis.Environment = ObjToChg.Environment
        objThis.fPath1 = ""
        objThis.fPath2 = ""
        objThis.IMSDBName = ""
        objThis.IsModified = ObjToChg.IsModified
        objThis.ObjFields.Clear()
        objThis.ObjTreeNode = ObjToChg.ObjTreeNode
        objThis.Parent = ObjToChg.Parent
        objThis.SegmentName = ""
        objThis.SeqNo = ObjToChg.SeqNo
        objThis.StructureDescription = ObjToChg.StructureDescription
        objThis.StructureName = ObjToChg.StructureName
        objThis.StructureSelections.Clear()
        objThis.StructureType = ObjToChg.StructureType
        objThis.Connection = ObjToChg.Connection


    End Function

    Private Sub frmStructure_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize

        pnlFields.SuspendLayout()
        pnlFields.Location = pnlSelect.Location
        pnlFields.Size = pnlSelect.Size
        pnlFields.ResumeLayout()

    End Sub

#End Region

#Region "Form Fill Functions"

    Function UpdateControls() As Boolean

        If IsNewObj = True Or IsChangeObj = True Then
            If (DMLFile <> "") Or (txtFilePath.Text <> "") Or (txtCobolFilePath.Text <> "") Or (chkAddDBD.Checked = True And txtCobolFilePath.Text <> "") Then
                cmdShowHideFieldAttr.Enabled = True
            Else
                cmdShowHideFieldAttr.Enabled = False
            End If
        Else
            cmdShowHideFieldAttr.Enabled = True
        End If

        If pnlSelect.Visible = True Then
            txtDBDFilePath.Visible = chkAddDBD.Checked
            cmdBrowseDBDFile.Visible = chkAddDBD.Checked
            tvSegments.Visible = chkAddDBD.Checked
            cmdShowHideFieldAttr.Text = "Show Fields/Attributes"
            cmdOk.Enabled = True
        Else
            cmdShowHideFieldAttr.Text = "Hide Fields/Attributes"
            cmdOk.Enabled = False
        End If

        If IsNewObj = False And IsChangeObj = False Then
            txtDBDFilePath.Enabled = False
            cmdBrowseDBDFile.Enabled = False
        End If

    End Function

    Function InitControls() As Boolean
        '//Add listview columns
        tvFields.HideSelection = False

        lvFieldAttrs.SmallImageList = imgListSmall

        pnlSelect.Visible = True
        pnlFields.Visible = False

        pnlFields.Size = pnlSelect.Size
        pnlFields.Location = pnlSelect.Location

        UpdateControls()

    End Function

    Sub SetComboConn(ByVal env As clsEnvironment)

        cbConn.Items.Clear()

        For Each conn As clsConnection In env.Connections
            cbConn.Items.Add(New Mylist(conn.ConnectionName, conn))
        Next
        cbConn.SelectedIndex = -1

    End Sub

    Friend Function ShowHideControls(ByVal StructType As enumStructure) As Boolean

        Dim bIsCOBOLorIMS As Boolean

        bIsCOBOLorIMS = ((StructType = enumStructure.STRUCT_COBOL) Or _
                         (StructType = enumStructure.STRUCT_IMS) Or _
                         (StructType = enumStructure.STRUCT_COBOL_IMS))

        If DMLFile = "" And (StructType = enumStructure.STRUCT_REL_DML_FILE Or StructType = enumStructure.STRUCT_REL_DML) Then
            '// for DML
            gbDMLConn.Visible = True
        Else
            gbDMLConn.Visible = False
        End If

        '//For COB, DBD
        lblCobolFile.Visible = bIsCOBOLorIMS
        lblDBD.Visible = bIsCOBOLorIMS
        'lblSeg.Visible = bIsCOBOLorIMS

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
        gbCOBOL.Visible = bIsCOBOLorIMS
        gbName.Visible = Not bIsCOBOLorIMS
        gbSeg.Visible = False

        Label7.Visible = False
        Label6.Visible = False

        '///set Get XML Message Button
        If StructType = enumStructure.STRUCT_XMLDTD Then
            btnGetXML.Visible = False
        Else
            btnGetXML.Visible = False
        End If


        '//Dont allow changing field once object is created coz it might conflict other place
        If IsNewObj = False And IsChangeObj = False Then
            cmdBrowseCobolFile.Enabled = False
            cmdBrowseDBDFile.Enabled = False
            cmdBrowseFile.Enabled = False
            txtFilePath.Enabled = False
            txtCobolFilePath.Enabled = False
            txtDBDFilePath.Enabled = False
            chkAddDBD.Enabled = False
        End If

        If StructType = modDeclares.enumStructure.STRUCT_COBOL And (IsNewObj = True Or IsChangeObj = True) Then
            'gbCOBOL.Visible = True
            txtCobolFilePath.Enabled = True
            txtDBDFilePath.Enabled = False
            chkAddDBD.Checked = False
            chkAddDBD.Visible = False
            cmdBrowseDBDFile.Visible = False
            lblDBD.Visible = False
            txtDBDFilePath.Visible = False
            tvSegments.Visible = False
            'lblSeg.Visible = False
        ElseIf StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS And (IsNewObj = True Or IsChangeObj = True) Then
            Label7.Visible = True
            Label7.BringToFront()
            Label6.Visible = True
            Label6.BringToFront()
            gbSeg.Visible = True
            txtCobolFilePath.Enabled = True
            txtDBDFilePath.Enabled = True
            chkAddDBD.Checked = True
        ElseIf StructType = modDeclares.enumStructure.STRUCT_IMS And (IsNewObj = True Or IsChangeObj = True) Then
            txtCobolFilePath.Enabled = False
            txtDBDFilePath.Enabled = True
            chkAddDBD.Checked = True
        End If

    End Function

    '//This will load treeview of fields
    '//If New Structure then fields will be from XML file
    '//If Edit structure then fields wil be from database
    Function FillFieldAttr(ByVal objThis As clsStructure) As Boolean

        Try
            tvFields.BeginUpdate()
            lvFieldAttrs.BeginUpdate()
            tvFields.Nodes.Clear()

            If IsNewObj = True Or IsChangeObj = True Then

                Dim outXMLPath As String
                outXMLPath = GetOutDumpXML(objThis)

                If outXMLPath = "" Then
                    Return False
                    Exit Try
                End If
                LoadTreeViewFromXmlFile(objThis, outXMLPath, tvFields, True)
                'If IO.File.Exists(outXMLPath) Then
                '    IO.File.Delete(outXMLPath)
                'End If
            Else
                objThis.LoadItems()
                '//This will fill out parent/child fields tree from Array of fields
                AddFieldsToTreeView(objThis, , tvFields)
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmStructure FillFieldAttr")
            Return False
        Finally
            tvFields.EndUpdate()
            lvFieldAttrs.EndUpdate()
        End Try

    End Function

    Function FillSegments() As Boolean

        Dim outXMLPath As String
        Try

            If IsNewObj = True Or IsChangeObj = True Then
                outXMLPath = GetSQDumpXML(txtDBDFilePath.Text, modDeclares.enumStructure.STRUCT_IMS)
                If outXMLPath = "" Then
                    Return False
                    Exit Try
                End If
                LoadTreeViewFromXmlFile(objThis, outXMLPath, tvSegments)
                'If IO.File.Exists(outXMLPath) Then
                '    IO.File.Delete(outXMLPath)
                'End If
            Else
                If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                    tvSegments.Nodes.Clear()
                    tvSegments.Nodes.Add(objThis.SegmentName)
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmStruct FillSegments")
            Return False
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


        Dim f As New System.Drawing.Font(lvFieldAttrs.Font, FontStyle.Bold)
        Dim i As Integer = 0

        For i = 0 To lvFieldAttrs.Items.Count - 1
            lvFieldAttrs.Items(i).UseItemStyleForSubItems = False
            lvFieldAttrs.Items(i).SubItems(0).Font = f
            lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Gray
            lvFieldAttrs.Items(i).SubItems(1).ForeColor = Color.Gray
            lvFieldAttrs.Items(i).ImageIndex = 30
        Next

        lblFieldName.Text = cNode.Text

        lvFieldAttrs.EndUpdate()

    End Function

#End Region

#Region "Click Events"

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

        Dim filename, segname As String
        Dim nd As TreeNode
        Dim filepath1 As String = ""
        Dim filepath2 As String = ""

        Me.Cursor = Cursors.WaitCursor

        Try
            If IsChangeObj = True Then
                If SaveChange() = True Then
                    Me.Close()
                    DialogResult = Windows.Forms.DialogResult.OK
                Else
                    DialogResult = Windows.Forms.DialogResult.Retry
                End If
                Exit Try
            End If

            'The next seven lines are added to resolve the exception caused if you 
            '// type the cobol file address instead of using the selection browser
            If (objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS) Then
                If chkAddDBD.Checked = False Then
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL
                Else
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
                End If
            End If

            If IsNewObj = True Then
                If Me.chkFTP.Checked Then
                    FileNames = Me.GottenFiles
                Else
                    If FileNames Is Nothing Then
                        '/// All but COBOL and IMS
                        If Not (objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or _
                                objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS) Then
                            '/// DML files
                            If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                                If gbDMLConn.Visible = False Then
                                    FileNames = New Collection
                                    FileNames.Add(DMLFile, DMLFile)
                                Else
                                    If txtFilePath.Text <> "" Then
                                        If cbConn.Text = "" Then
                                            MsgBox("You Must choose a Connection", MsgBoxStyle.Information, MsgTitle)
                                            Me.Cursor = Cursors.Default
                                            DialogResult = Windows.Forms.DialogResult.Retry
                                            Exit Try
                                        Else
                                            FileNames = New Collection
                                            objThis.Connection.LoadMe()
                                            FileNames.Add(objThis.Connection.UserId & "." & _
                                                          objThis.Connection.Password & "." & _
                                                          objThis.Connection.Database & "." & _
                                                          txtFilePath.Text, _
                                                          objThis.Connection.UserId & "." & _
                                                          objThis.Connection.Password & "." & _
                                                          objThis.Connection.Database & "." & _
                                                          txtFilePath.Text)
                                        End If
                                    Else
                                        MsgBox("File path cannot be blank.", MsgBoxStyle.Critical, MsgTitle)
                                        Me.Cursor = Cursors.Default
                                        DialogResult = Windows.Forms.DialogResult.Retry
                                        Exit Try
                                    End If
                                End If
                                '/// All others excluding COBOL IMS
                            Else
                                If txtFilePath.Text <> "" Then
                                    FileNames = New Collection
                                    FileNames.Add(txtFilePath.Text.Clone, txtFilePath.Text.Clone)
                                Else
                                    MsgBox("File path cannot be blank.", MsgBoxStyle.Critical, MsgTitle)
                                    Me.Cursor = Cursors.Default
                                    DialogResult = Windows.Forms.DialogResult.Retry
                                    Exit Try
                                End If
                            End If
                        Else
                            If txtCobolFilePath.Text <> "" Then
                                FileNames = New Collection
                                FileNames.Add(txtCobolFilePath.Text.Clone, txtCobolFilePath.Text.Clone)
                            Else
                                MsgBox("Cobol file path cannot be blank.", MsgBoxStyle.Critical, MsgTitle)
                                Me.Cursor = Cursors.Default
                                DialogResult = Windows.Forms.DialogResult.Retry
                                Exit Try
                            End If
                        End If
                    End If
                End If

                '//check if multiple files are selected
                'If FileNames.Count > 1 Then  '(txtStructName.Text = "")
                'If FileNames.Count >= 1 Then
                '//if yes

                Dim saveName As String

                For Each filename In FileNames
                    If txtFilePath.Visible Then
                        If gbDMLConn.Visible = False Then
                            filepath1 = filename
                            filepath2 = ""
                        Else
                            If cbConn.Text = "" Then
                                MsgBox("You Must choose a Connection", MsgBoxStyle.Information, MsgTitle)
                                Me.Cursor = Cursors.Default
                                DialogResult = Windows.Forms.DialogResult.Retry
                                Exit Try
                            Else
                                objThis.Connection.LoadMe()
                                filepath1 = objThis.Connection.UserId & "." & _
                                objThis.Connection.Password & "." & _
                                objThis.Connection.Database & "." & _
                                filename
                                filepath2 = ""
                            End If
                        End If
                    Else
                        filepath1 = filename
                        filepath2 = txtDBDFilePath.Text
                    End If

                    '/// COBOL IMS files
                    If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                        '/// Let user choose Segment in tree in case of COBOL filename different than segment name
                        If tvSegments.SelectedNode Is Nothing Then
                            segname = IO.Path.GetFileNameWithoutExtension(filename)
                            segname = Strings.UCase(segname)
                            nd = SelectFirstMatchingNode(tvSegments, segname, , , True)
                        Else
                            nd = tvSegments.SelectedNode
                            segname = nd.Text
                        End If



                        If Not nd Is Nothing Then
                            '//no matching segment found for selected Cobol file 
                            '//so this file is skipped
                            Dim obj As clsStructure
                            obj = objThis.Clone(objThis.Parent)
                            '//We found matching segment for one of the selected 
                            '//cobol file name
                            'saveName = segname
                            saveName = IO.Path.GetFileNameWithoutExtension(filename)

TryAgain1:                  If DoSave(obj, saveName, txtStructDesc.Text, filepath1, filepath2, True) = True Then
                                nd.BackColor = Color.GreenYellow
                                nd.EnsureVisible()
                                Me.Text = "Adding ... [" & IO.Path.GetFileName(filename) & "]"
                            Else
                                saveName = InputBox("Please type a different name for your Description", "Change Duplicate Description Name", saveName)
                                If saveName = "" Then
                                    GoTo nextFilename
                                Else
                                    GoTo TryAgain1
                                End If
                            End If
                        Else
                            MsgBox("Could not find segment: " & segname & " in file " & _
                            IO.Path.GetFileName(filepath2) & Chr(13) & "Please add this Description individually, and choose a segment for it.", MsgBoxStyle.OkOnly, "Missing Segment")
                            ErrorLog("Segment: " & segname & " could not be found in " & IO.Path.GetFileName(filepath2))
                        End If
                    Else
                        Dim obj As clsStructure
                        obj = objThis.Clone(objThis.Parent)
                        saveName = IO.Path.GetFileNameWithoutExtension(filename)
                        '//We found matching segment for one of the selected cobol file name
TryAgain2:              If DoSave(obj, saveName, txtStructDesc.Text, filepath1, filepath2, True) = True Then
                            Me.Text = "Adding ... [" & IO.Path.GetFileName(filename) & "]"
                        Else
                            saveName = InputBox("Please type a different name for your Description", "Change Duplicate Description Name", saveName)
                            If saveName = "" Then
                                GoTo nextFilename
                            Else
                                GoTo TryAgain2
                            End If
                        End If
                    End If
                    tvSegments.SelectedNode = Nothing
nextFilename:   Next
                Me.Text = "Description Properties"
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.OK


                'Else
                '    If txtFilePath.Visible = True Then
                '        If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                '            If gbDMLConn.Visible = False Then
                '                filepath1 = DMLFile
                '                filepath2 = ""
                '            Else
                '                If cbConn.Text = "" Then
                '                    MsgBox("You Must choose a Connection", MsgBoxStyle.Information, MsgTitle)
                '                    Me.Cursor = Cursors.Default
                '                    DialogResult = Windows.Forms.DialogResult.Retry
                '                    Exit Try
                '                Else
                '                    filepath1 = objThis.Connection.UserId & "." & objThis.Connection.Password & "." & objThis.Connection.Database & "." & txtFilePath.Text
                '                    filepath2 = ""
                '                End If
                '            End If
                '        Else
                '            filepath1 = txtFilePath.Text
                '            filepath2 = ""
                '        End If
                '    Else
                '        filepath1 = txtCobolFilePath.Text
                '        filepath2 = txtDBDFilePath.Text
                '    End If
                '    If DoSave(objThis, txtStructName.Text, txtStructDesc.Text, filepath1, filepath2) = True Then
                '        DialogResult = Windows.Forms.DialogResult.OK
                '    Else
                '        DialogResult = Windows.Forms.DialogResult.Retry
                '    End If
                'End If
            Else
                If txtFilePath.Visible = True Then
                    If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                        If gbDMLConn.Visible = False Then
                            filepath1 = DMLFile
                            filepath2 = ""
                        Else
                            If cbConn.Text = "" Then
                                MsgBox("You Must choose a Connection", MsgBoxStyle.Information, MsgTitle)
                                Me.Cursor = Cursors.Default
                                DialogResult = Windows.Forms.DialogResult.Retry
                                Exit Try
                            Else
                                objThis.Connection.LoadMe()
                                filepath1 = objThis.Connection.UserId & "." & _
                                objThis.Connection.Password & "." & _
                                objThis.Connection.Database & "." & _
                                txtFilePath.Text
                                filepath2 = ""
                            End If
                        End If
                    Else
                        filepath1 = txtFilePath.Text
                        filepath2 = ""
                    End If
                Else
                    filepath1 = txtCobolFilePath.Text
                    filepath2 = txtDBDFilePath.Text
                End If
                If DoSave(objThis, txtStructName.Text, txtStructDesc.Text, filepath1, filepath2) = True Then
                    DialogResult = Windows.Forms.DialogResult.OK
                Else
                    DialogResult = Windows.Forms.DialogResult.Retry
                End If
            End If

        Catch ex As Exception
            LogError(ex, "frmStructure cmdOKclick")
            DialogResult = Windows.Forms.DialogResult.Retry
            If IsNewObj = True Then
                '//if operation fails while adding multiple structure at once 
                '//then we need to rollback all addition occured during that operation
                RollbackAddOperation()
            End If
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub chkAddDBD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkAddDBD.CheckedChanged

        UpdateControls()

    End Sub

    Private Sub cmdBrowseFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFile.Click

        Dim strFilter As String
        Dim filename As String  '// added by TK and KS 11/6/2006
        Dim extString As String = ".*"

        objEnv = objThis.Environment
        objEnv.LoadMe()
        If IsChangeObj = False Then
            OpenMultiFileDialog1.Multiselect = True
        Else
            OpenMultiFileDialog1.Multiselect = False
        End If

        Select Case objThis.StructureType
            Case modDeclares.enumStructure.STRUCT_C
                strFilter = "C Header files (*.h)|*.h|All files (*.*)|*.*"
                If objEnv.LocalCDir <> "" Then
                    OpenMultiFileDialog1.InitialDirectory = objEnv.LocalCDir
                Else
                    OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalCFolderPath
                End If
                OpenMultiFileDialog1.FileName = ""
                extString = ".h"
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                strFilter = "Relational DDL files (*.ddl)|*.ddl|SQLserver SQL files (*.sql)|*.sql|All files (*.*)|*.*"
                If objEnv.LocalDDLDir <> "" Then
                    OpenMultiFileDialog1.InitialDirectory = objEnv.LocalDDLDir
                Else
                    OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalDDLFolderPath
                End If
                OpenMultiFileDialog1.FileName = ""
                extString = ".ddl"
            Case modDeclares.enumStructure.STRUCT_REL_DML_FILE
                strFilter = "Relational DML files (*.dml)|*.dml|All files (*.*)|*.*"
                If objEnv.LocalDDLDir <> "" Then
                    OpenMultiFileDialog1.InitialDirectory = objEnv.LocalDMLDir
                Else
                    OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalDDLFolderPath
                End If
                OpenMultiFileDialog1.FileName = ""
                extString = ".dml"
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                strFilter = "XML DTD Files (*.dtd)|*.dtd|XML Message Files (*.xml)|*.xml|All files (*.*)|*.*"
                If objEnv.LocalDTDDir <> "" Then
                    OpenMultiFileDialog1.InitialDirectory = objEnv.LocalDTDDir
                Else
                    OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalDTDFolderPath
                End If
                OpenMultiFileDialog1.FileName = ""
                extString = ".dtd"
            Case Else
                strFilter = "All files (*.*)|*.*"
        End Select

        OpenMultiFileDialog1.Filter = strFilter

        If Me.chkFTP.Checked Then
            Dim FTPClient As frmFTPClient = New frmFTPClient
            Me.GottenFiles = FTPClient.BrowseFile(OpenMultiFileDialog1.InitialDirectory, extString)
            If (Me.GottenFiles.Count > 0) Then
                Me.txtFilePath.Text = OpenMultiFileDialog1.InitialDirectory & "\" & Me.GottenFiles(1)
            End If
        Else
            If OpenMultiFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtFilePath.Text = OpenMultiFileDialog1.FileName
                If txtStructName.Text.Trim = "" Then
                    txtStructName.Text = IO.Path.GetFileNameWithoutExtension(txtFilePath.Text)
                End If

                '// added by KS and TK 11/6/2006
                txtFilePath.Text = ""
                FileNames = New Collection
                For Each filename In OpenMultiFileDialog1.FileNames
                    txtFilePath.Text = txtFilePath.Text & filename & " "
                    FileNames.Add(filename.Clone, filename.Clone)
                Next

                txtFilePath.Text = txtFilePath.Text.Trim

                UpdateControls()
            End If
        End If

        OpenMultiFileDialog1.Multiselect = False

    End Sub

    Private Sub cmdBrowseCobolFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseCobolFile.Click

        Dim filename As String

        OpenMultiFileDialog1.Multiselect = True
        objEnv = objThis.Environment
        objEnv.LoadMe()
        OpenMultiFileDialog1.Filter = "Cobol Copybook (*.cob)|*.cob|All files (*.*)|*.*"

        If objEnv.LocalCobolDir <> "" Then
            OpenMultiFileDialog1.InitialDirectory = objEnv.LocalCobolDir
        Else
            OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalCOBOLFolderPath
        End If
        OpenMultiFileDialog1.FileName = ""

        If Me.chkFTP.Checked Then
            Dim FTPClient As frmFTPClient = New frmFTPClient
            Me.GottenFiles = FTPClient.BrowseFile(OpenMultiFileDialog1.InitialDirectory, ".cob")
            If (Me.GottenFiles.Count > 0) Then
                Me.txtCobolFilePath.Text = OpenMultiFileDialog1.InitialDirectory & "\" & Me.GottenFiles(1)
            End If
        Else
            If OpenMultiFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                txtCobolFilePath.Text = ""         'KS100906'
                FileNames = New Collection
                For Each filename In OpenMultiFileDialog1.FileNames
                    txtCobolFilePath.Text = txtCobolFilePath.Text & OpenMultiFileDialog1.FileName & " " 'KS100906'
                    FileNames.Add(filename.Clone, filename.Clone)
                Next
                If txtStructName.Text.Trim = "" And FileNames.Count <= 1 Then
                    txtStructName.Text = IO.Path.GetFileNameWithoutExtension(txtCobolFilePath.Text)
                End If
                txtCobolFilePath.Text = txtCobolFilePath.Text.Trim 'KS100906'

                If (chkAddDBD.Checked = True And txtDBDFilePath.Text <> "") Then
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
                Else
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL
                End If
                UpdateControls()
                'LocalCOBOLFolderPath = dlgOpen.FileName '//new: 10/8/05
            End If
        End If

        OpenMultiFileDialog1.Multiselect = False

    End Sub

    Private Sub cmdBrowseDBDFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseDBDFile.Click

        Dim tmpGottenFiles As Collection

        Try
            objEnv = objThis.Environment
            objEnv.LoadMe()
            OpenMultiFileDialog1.Filter = "IMS DBD files (*.dbd)|*.dbd|All files (*.*)|*.*"

            If objEnv.LocalDBDDir <> "" Then
                OpenMultiFileDialog1.InitialDirectory = objEnv.LocalDBDDir
            Else
                OpenMultiFileDialog1.InitialDirectory = objEnv.DefaultStrDir 'LocalDBDFolderPath
            End If
            OpenMultiFileDialog1.FileName = ""

            If Me.chkFTP.Checked Then
                Dim FTPClient As frmFTPClient = New frmFTPClient

                tmpGottenFiles = FTPClient.BrowseFile(OpenMultiFileDialog1.InitialDirectory, ".dbd")
                If (Me.GottenFiles.Count > 0) Then
                    Me.txtDBDFilePath.Text = OpenMultiFileDialog1.InitialDirectory & "\" & tmpGottenFiles(0)
                End If

            Else
                If OpenMultiFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                    txtDBDFilePath.Text = OpenMultiFileDialog1.FileName

                    If (txtCobolFilePath.Text <> "") Then
                        objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
                    Else
                        objThis.StructureType = modDeclares.enumStructure.STRUCT_IMS
                    End If

                    '//As soon user select *.dbd file display available segments in the file
                    FillSegments()

                    UpdateControls()
                End If
            End If

        Catch ex As Exception
            LogError(ex, "frmStruct browseDBD")
        End Try

    End Sub

    Private Sub cbConn_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbConn.SelectedIndexChanged

        If cbConn.Text <> "" Then
            objThis.Connection = CType(cbConn.SelectedItem, Mylist).ItemData
        End If

    End Sub

#End Region

#Region "Help"

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Add_Structures)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

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

#Region "TreeView and ListView Functions"

    Private Sub tvSegments_AfterCheck(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvSegments.AfterCheck
        '//Only process this logic if action using mouse or keyboard other wise this event can fire on and on ....
        If IsEventFromCode = False Then
            CheckUncheckNodes(e.Node)
        End If
    End Sub

    Private Sub cmdShowHideFieldAttr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdShowHideFieldAttr.Click

        Dim path As String

        'UpdateControls()

        If (txtCobolFilePath.Text.Trim = "") Then
            path = txtFilePath.Text
        Else
            path = txtCobolFilePath.Text
        End If

        If IsNewObj = True And DMLFile = "" Then
            If ValidateSelection(path, True) = False Then Exit Sub
        End If

        pnlFields.Visible = Not pnlFields.Visible
        pnlSelect.Visible = Not pnlSelect.Visible

        UpdateControls()  '// moved here so that view/hide fields  button is correct

        '//Dont reload everytime once its loaded from database
        If tvFields.GetNodeCount(True) > 0 And IsNewObj = False Then Exit Sub

        tvFields.BackColor = Color.LightBlue
        lvFieldAttrs.BackColor = Color.LightBlue
        lblFieldName.ForeColor = Color.Red
        lblFieldName.Text = "********* Loading *********"
        Me.Refresh()

        If pnlFields.Visible = True Then
            If objThis.StructureType = enumStructure.STRUCT_REL_DML Or objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                If DMLFile <> "" Then
                    objThis.fPath1 = DMLFile
                Else
                    objThis.fPath1 = objThis.Connection.UserId & "." & objThis.Connection.Password & "." & objThis.Connection.Database & "." & path
                End If
            Else
                objThis.fPath1 = path
            End If

            If objThis.StructureType = enumStructure.STRUCT_COBOL Or objThis.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                objThis.fPath2 = txtDBDFilePath.Text
            Else
                objThis.fPath2 = ""
            End If

            If FillFieldAttr(objThis) = False Then
                'DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If
        End If

        tvFields.BackColor = Color.White
        lvFieldAttrs.BackColor = Color.White
        lblFieldName.Text = ""
        lblFieldName.ForeColor = Color.Blue

        If tvFields.GetNodeCount(True) > 0 Then
            tvFields.SelectedNode = tvFields.Nodes.Item(0)
        End If

    End Sub

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub txtSearchField_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchField.TextChanged

        modGeneral.SelectFirstMatchingNode(tvFields, txtSearchField.Text)

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

    End Sub

#End Region

    Sub RollbackAddOperation()

        Dim obj As clsStructure = Nothing
        'Dim cnn As System.Data.Odbc.OdbcConnection
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        'cnn = New Odbc.OdbcConnection(obj.Project.MetaConnectionString)
        'cnn.Open()
        Try
            For Each obj In arrObjThis
                cmd.Connection = cnn
                '//We need to put in transaction because we will add structure 
                '// and fields in two steps so if one fails rollback all
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                If obj.Delete(cmd, cnn, True) = True Then
                    tran.Commit()
                Else
                    tran.Rollback()
                End If
            Next
            arrObjThis.Clear()

        Catch ex As Exception
            tran.Rollback()
            LogError(ex)
        Finally
            'cnn.Close()
        End Try

    End Sub

    Function DoSave(ByVal objThis As clsStructure, ByVal strStrctureName As String, ByVal strStrctureDesc As String, Optional ByVal fPath1 As String = "", Optional ByVal fPath2 As String = "", Optional ByVal Skip_Name_Validation As Boolean = False) As Boolean

        Try
            If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                If objThis.ValidateNewObject(strStrctureName) = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Return False
                    Exit Function
                Else
                    DialogResult = Windows.Forms.DialogResult.OK
                End If
            Else
                If ValidateSelection(fPath1, Skip_Name_Validation) = False Or objThis.ValidateNewObject(strStrctureName) = False Or ValidateNewName128(strStrctureName) = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Return False
                    Exit Function
                Else
                    DialogResult = Windows.Forms.DialogResult.OK
                End If
            End If
            

            objThis.StructureName = strStrctureName.Trim     '/// changed 12/07 by TK
            objThis.StructureDescription = strStrctureDesc '//new : 8/15/05

            '//Make all file path case-sensitive : Unix uses case sensitive names so if user enters c:\wrseg.cob and on disk name appears WRSeg.cob then correct it to WRSeg.cob
            If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                objThis.fPath1 = GetCaseSensetivePath(fPath1)
                objThis.fPath2 = GetCaseSensetivePath(fPath2)
            ElseIf objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                objThis.fPath1 = fPath1
            Else
                objThis.fPath1 = GetCaseSensetivePath(fPath1)
            End If

            If IsNewObj = True Then
                Dim IMSDBName As String
                Dim XMLFilePath As String
                XMLFilePath = GetOutDumpXML(objThis)

                If XMLFilePath = "" Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Return False
                    Exit Function
                End If
                '//Before we add new structure set structure fields
                IMSDBName = LoadFldArrFromXmlFile(XMLFilePath, objThis.ObjFields)

                '//load array doesnt set the paren so we do manually
                Dim i As Integer
                For i = 0 To objThis.ObjFields.Count - 1
                    objThis.ObjFields(i).Parent = objThis
                    objThis.ObjFields(i).ParentName = objThis.Text
                    objThis.ObjFields(i).SeqNo = i
                Next

                objThis.IMSDBName = IMSDBName

                If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                    Dim StrArr As String()
                    StrArr = Strings.Split(objThis.fPath1, ".")
                    objThis.fPath1 = StrArr(3) & "." & StrArr(4)
                End If

                If objThis.AddNew = True Then
                    '//if adding new structure then store in arraylist
                    arrObjThis.Add(objThis)
                Else
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Return False
                    Exit Function
                End If
            Else
                If objThis.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                    objThis.fPath1 = txtFilePath.Text
                End If
                objThis.Save()
            End If

            Return True

        Catch ex As Exception
            LogError(ex)
            Return False
        End Try

    End Function

    Function SaveChange() As Boolean

        Dim filepath1 As String = ""
        Dim filepath2 As String = ""

        Me.Cursor = Cursors.WaitCursor

        Try
            'The next seven lines are added to resolve the exception caused if you 
            '// type the cobol file address instead of using the selection browser
            If (objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS) Then
                If chkAddDBD.Checked = False Then
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL
                Else
                    objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS
                End If
            End If

            If Me.chkFTP.Checked Then
                FileNames = Me.GottenFiles
                If FileNames.Count > 1 Then
                    MsgBox("You may only replace one structure file at a time", MsgBoxStyle.Information, MsgTitle)
                    SaveChange = False
                    Exit Function
                End If
            Else
                If FileNames Is Nothing Then
                    If Not (objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS) Then
                        If txtFilePath.Text = "" Then
                            MsgBox("File path cannot be blank.", MsgBoxStyle.Critical, MsgTitle)
                            Me.Cursor = Cursors.Default
                            SaveChange = False
                            Exit Function
                        End If
                    Else
                        If txtCobolFilePath.Text = "" Then
                            MsgBox("Cobol file path cannot be blank.", MsgBoxStyle.Critical, MsgTitle)
                            Me.Cursor = Cursors.Default
                            SaveChange = False
                            Exit Function
                        End If
                    End If
                End If
            End If

            If txtFilePath.Visible = True Then
                filepath1 = txtFilePath.Text
                filepath2 = ""
            Else
                filepath1 = txtCobolFilePath.Text
                filepath2 = txtDBDFilePath.Text
            End If

            '/// Everything in New Object is set, Now Try to save the changed structure file
            If DoSaveChange(filepath1, filepath2) = True Then
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.OK
                SaveChange = True
            Else
                DialogResult = Windows.Forms.DialogResult.Retry
                SaveChange = False
            End If

        Catch ex As Exception
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
            SaveChange = False
        Finally
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Function DoSaveChange(ByVal fPath1 As String, ByVal fPath2 As String) As Boolean

        Dim IMSDBName As String
        Dim XMLFilePath As String
        Dim i As Integer
        
        Try
            'MsgBox("fpath1 = " & fPath1 & " and fpath 2 = " & fPath2)

            '//Make all file path case-sensitive : Unix uses case sensitive names so if 
            '//user enters c:\wrseg.cob and on disk name appears WRSeg.cob then correct it to WRSeg.cob
            If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Or _
            objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                objThis.fPath1 = GetCaseSensetivePath(fPath1)
                objThis.fPath2 = GetCaseSensetivePath(fPath2)
            Else
                objThis.fPath1 = GetCaseSensetivePath(fPath1)
            End If

            XMLFilePath = GetOutDumpXML(objThis)

            If XMLFilePath = "" Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Function
            End If
            '//Before we add new structure set structure fields
            objThis.ObjFields.Clear()
            IMSDBName = LoadFldArrFromXmlFile(XMLFilePath, objThis.ObjFields)

            '//load array doesnt set the parent so we do manually

            For i = 0 To objThis.ObjFields.Count - 1
                objThis.ObjFields(i).Parent = objThis
                objThis.ObjFields(i).ParentName = objThis.StructureName
                objThis.ObjFields(i).SeqNo = i
            Next

            objThis.IMSDBName = IMSDBName

            '/// Now go to modRename to delete the old structure from the environment
            '/// and add the new structure to the environment, If function returns true then
            '// change was successful
            DoSaveChange = ReplaceStrFile(objThis, ObjToChg)  'ObjToChg.Environment.

        Catch ex As Exception
            LogError(ex, "frmstructure DoSaveChange")
            DoSaveChange = False
        End Try

    End Function

    Function ValidateSelection(ByVal path As String, Optional ByVal Skip_ObjectName_Validation As Boolean = False) As Boolean

        ValidateSelection = True

        If IsNewObj = False Then Exit Function '//new : 8/15/05

        If Skip_ObjectName_Validation = False AndAlso txtStructName.Text.Trim = "" Then
            MsgBox("Object name can not be blank.", MsgBoxStyle.Critical, MsgTitle)
            ValidateSelection = False
            Exit Function
        End If

        If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then

            If path.Trim = "" Then
                MsgBox("Please specify a cobol file (*.cob) path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(path) = False Then 'KS100906'
                MsgBox("Can not find the specified file " & path & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If chkAddDBD.Checked = True Then
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
            If path.Trim = "" Then
                MsgBox("Please specify a cobol file (*.cob) path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(path) = False Then
                MsgBox("Can not find the specified file " & path & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
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
            If path.Trim = "" Then  '// modified by TK and KS 11/6/2006
                MsgBox("Please select a valid file.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If

            If IO.File.Exists(path) = False Then '// modified by TK and KS 11/6/2006
                MsgBox("Can not find the specified file " & txtFilePath.Text & ". Please make sure that you have entered a correct file path.", MsgBoxStyle.Critical, MsgTitle)
                ValidateSelection = False
                Exit Function
            End If
        End If

    End Function

    Function GetOutDumpXML(ByVal objThis As clsStructure) As String

        If objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL Then
            GetOutDumpXML = GetSQDumpXML(objThis.fPath1, objThis.StructureType)
        ElseIf objThis.StructureType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
            If tvSegments.SelectedNode IsNot Nothing Then
                objThis.SegmentName = tvSegments.SelectedNode.Text
                GetOutDumpXML = GetSQDumpXML(objThis.fPath2, objThis.StructureType, objThis.SegmentName, Quote(objThis.fPath1, """"))
            Else
                MsgBox("Please select an IMS segment", MsgBoxStyle.Information, "Select Segment")
                GetOutDumpXML = ""
            End If
        Else
            '///If it's an XML Message, convert it first to DTD
            If objThis.fPath1.Contains(".xml") = True Then
                Dim oldName As String = objThis.fPath1
                Dim OutPath As String = ProcessXMLmessage(objThis.fPath1, objEnv.LocalDTDDir)
                If OutPath <> "" Then
                    objThis.fPath1 = OutPath
                    txtFilePath.Text = OutPath
                    txtStructName.Text = GetFileNameWithoutExtenstionFromPath(OutPath)
                    If FileNames.Contains(oldName) = True Then
                        FileNames.Remove(oldName)
                    End If
                    FileNames.Add(OutPath, OutPath)
                End If
            End If
            GetOutDumpXML = GetSQDumpXML(objThis.fPath1, objThis.StructureType)
        End If

    End Function

    'Private Sub btnGetXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetXML.Click

    '    Try
    '        Dim frm As New frmXMLconv
    '        Dim DTDpath As String

    '        'Open XMLconversion Form and get a converted XML Message
    '        DTDpath = frm.GetDTD()

    '        If DTDpath <> "" Then
    '            txtFilePath.Text = DTDpath
    '            txtStructName.Text = GetFileNameWithoutExtenstionFromPath(DTDpath)
    '            UpdateControls()
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "frmStructure btnGetXML_Click")
    '    End Try

    'End Sub

    'Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

    'End Sub
End Class


