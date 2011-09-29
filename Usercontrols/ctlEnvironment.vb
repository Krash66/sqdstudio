Public Class ctlEnvironment
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    
    Dim objThis As New clsEnvironment
    Dim IsNewObj As Boolean
    Friend WithEvents FBD1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents gbRelDir As System.Windows.Forms.GroupBox
    Friend WithEvents txtRelDir As System.Windows.Forms.TextBox
    Friend WithEvents btnRelDir As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label


    Private destination As String

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
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtLocalCDir As System.Windows.Forms.TextBox
    Friend WithEvents txtEnvironmentDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtEnvironmentName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtLocalDTDDir As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtLocalDDLDir As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtLocalCobolDir As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLocalScriptDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalC As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalCobol As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalDDL As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalDTD As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalScript As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtLocalDBDDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseFileLocalModel As System.Windows.Forms.Button
    Friend WithEvents cmdPutScr As System.Windows.Forms.Button
    Friend WithEvents btnBrowseFileLocalDML As System.Windows.Forms.Button
    Friend WithEvents txtLocalDMLDir As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents gbName As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbScrDir As System.Windows.Forms.GroupBox
    Friend WithEvents gbFTP As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents btnFTPdtd As System.Windows.Forms.Button
    Friend WithEvents btnFTPdbd As System.Windows.Forms.Button
    Friend WithEvents btnFTPddl As System.Windows.Forms.Button
    Friend WithEvents btnFTPCobol As System.Windows.Forms.Button
    Friend WithEvents btnFTPc As System.Windows.Forms.Button
    Friend WithEvents btnFTPdml As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtLocalCDir = New System.Windows.Forms.TextBox
        Me.txtEnvironmentDesc = New System.Windows.Forms.TextBox
        Me.txtEnvironmentName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtLocalDTDDir = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtLocalDDLDir = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtLocalCobolDir = New System.Windows.Forms.TextBox
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalC = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalCobol = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalDDL = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalDTD = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalScript = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtLocalScriptDir = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtLocalDBDDir = New System.Windows.Forms.TextBox
        Me.cmdBrowseFileLocalModel = New System.Windows.Forms.Button
        Me.cmdPutScr = New System.Windows.Forms.Button
        Me.btnBrowseFileLocalDML = New System.Windows.Forms.Button
        Me.txtLocalDMLDir = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.gbName = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbScrDir = New System.Windows.Forms.GroupBox
        Me.gbFTP = New System.Windows.Forms.GroupBox
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.btnFTPdtd = New System.Windows.Forms.Button
        Me.btnFTPdbd = New System.Windows.Forms.Button
        Me.btnFTPddl = New System.Windows.Forms.Button
        Me.btnFTPCobol = New System.Windows.Forms.Button
        Me.btnFTPc = New System.Windows.Forms.Button
        Me.btnFTPdml = New System.Windows.Forms.Button
        Me.FBD1 = New System.Windows.Forms.FolderBrowserDialog
        Me.gbRelDir = New System.Windows.Forms.GroupBox
        Me.btnRelDir = New System.Windows.Forms.Button
        Me.txtRelDir = New System.Windows.Forms.TextBox
        Me.gbName.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbScrDir.SuspendLayout()
        Me.gbFTP.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.gbRelDir.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(337, 507)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 201
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(417, 507)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 202
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 491)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(561, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 113)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(15, 14)
        Me.Label8.TabIndex = 43
        Me.Label8.Text = "C"
        '
        'txtLocalCDir
        '
        Me.txtLocalCDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalCDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalCDir.Location = New System.Drawing.Point(54, 110)
        Me.txtLocalCDir.MaxLength = 300
        Me.txtLocalCDir.Name = "txtLocalCDir"
        Me.txtLocalCDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalCDir.TabIndex = 5
        '
        'txtEnvironmentDesc
        '
        Me.txtEnvironmentDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEnvironmentDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEnvironmentDesc.Location = New System.Drawing.Point(6, 19)
        Me.txtEnvironmentDesc.MaxLength = 1000
        Me.txtEnvironmentDesc.Multiline = True
        Me.txtEnvironmentDesc.Name = "txtEnvironmentDesc"
        Me.txtEnvironmentDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEnvironmentDesc.Size = New System.Drawing.Size(319, 42)
        Me.txtEnvironmentDesc.TabIndex = 8
        '
        'txtEnvironmentName
        '
        Me.txtEnvironmentName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEnvironmentName.Location = New System.Drawing.Point(54, 19)
        Me.txtEnvironmentName.MaxLength = 128
        Me.txtEnvironmentName.Name = "txtEnvironmentName"
        Me.txtEnvironmentName.Size = New System.Drawing.Size(277, 20)
        Me.txtEnvironmentName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 34
        Me.Label3.Text = "Name"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 165)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 14)
        Me.Label7.TabIndex = 33
        Me.Label7.Text = "DTD"
        '
        'txtLocalDTDDir
        '
        Me.txtLocalDTDDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDTDDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalDTDDir.Location = New System.Drawing.Point(54, 162)
        Me.txtLocalDTDDir.MaxLength = 300
        Me.txtLocalDTDDir.Name = "txtLocalDTDDir"
        Me.txtLocalDTDDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalDTDDir.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 87)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(28, 14)
        Me.Label5.TabIndex = 36
        Me.Label5.Text = "DDL"
        '
        'txtLocalDDLDir
        '
        Me.txtLocalDDLDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDDLDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalDDLDir.Location = New System.Drawing.Point(54, 84)
        Me.txtLocalDDLDir.MaxLength = 300
        Me.txtLocalDDLDir.Name = "txtLocalDDLDir"
        Me.txtLocalDDLDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalDDLDir.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 14)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Cobol"
        '
        'txtLocalCobolDir
        '
        Me.txtLocalCobolDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalCobolDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalCobolDir.Location = New System.Drawing.Point(54, 32)
        Me.txtLocalCobolDir.MaxLength = 300
        Me.txtLocalCobolDir.Name = "txtLocalCobolDir"
        Me.txtLocalCobolDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalCobolDir.TabIndex = 4
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(497, 507)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 203
        Me.cmdHelp.Text = "&Help"
        '
        'cmdBrowseFileLocalC
        '
        Me.cmdBrowseFileLocalC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalC.Location = New System.Drawing.Point(483, 111)
        Me.cmdBrowseFileLocalC.Name = "cmdBrowseFileLocalC"
        Me.cmdBrowseFileLocalC.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalC.TabIndex = 105
        Me.cmdBrowseFileLocalC.Text = "..."
        '
        'cmdBrowseFileLocalCobol
        '
        Me.cmdBrowseFileLocalCobol.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalCobol.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalCobol.Location = New System.Drawing.Point(483, 31)
        Me.cmdBrowseFileLocalCobol.Name = "cmdBrowseFileLocalCobol"
        Me.cmdBrowseFileLocalCobol.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalCobol.TabIndex = 104
        Me.cmdBrowseFileLocalCobol.Text = "..."
        '
        'cmdBrowseFileLocalDDL
        '
        Me.cmdBrowseFileLocalDDL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalDDL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalDDL.Location = New System.Drawing.Point(483, 83)
        Me.cmdBrowseFileLocalDDL.Name = "cmdBrowseFileLocalDDL"
        Me.cmdBrowseFileLocalDDL.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalDDL.TabIndex = 103
        Me.cmdBrowseFileLocalDDL.Text = "..."
        '
        'cmdBrowseFileLocalDTD
        '
        Me.cmdBrowseFileLocalDTD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalDTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalDTD.Location = New System.Drawing.Point(483, 161)
        Me.cmdBrowseFileLocalDTD.Name = "cmdBrowseFileLocalDTD"
        Me.cmdBrowseFileLocalDTD.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalDTD.TabIndex = 102
        Me.cmdBrowseFileLocalDTD.Text = "..."
        '
        'cmdBrowseFileLocalScript
        '
        Me.cmdBrowseFileLocalScript.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalScript.Location = New System.Drawing.Point(483, 29)
        Me.cmdBrowseFileLocalScript.Name = "cmdBrowseFileLocalScript"
        Me.cmdBrowseFileLocalScript.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalScript.TabIndex = 106
        Me.cmdBrowseFileLocalScript.Text = "..."
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(39, 14)
        Me.Label1.TabIndex = 111
        Me.Label1.Text = "Script"
        '
        'txtLocalScriptDir
        '
        Me.txtLocalScriptDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalScriptDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalScriptDir.Location = New System.Drawing.Point(54, 30)
        Me.txtLocalScriptDir.MaxLength = 300
        Me.txtLocalScriptDir.Name = "txtLocalScriptDir"
        Me.txtLocalScriptDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalScriptDir.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 61)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(28, 14)
        Me.Label2.TabIndex = 113
        Me.Label2.Text = "DBD"
        '
        'txtLocalDBDDir
        '
        Me.txtLocalDBDDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDBDDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalDBDDir.Location = New System.Drawing.Point(54, 58)
        Me.txtLocalDBDDir.Name = "txtLocalDBDDir"
        Me.txtLocalDBDDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalDBDDir.TabIndex = 7
        '
        'cmdBrowseFileLocalModel
        '
        Me.cmdBrowseFileLocalModel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalModel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalModel.Location = New System.Drawing.Point(483, 59)
        Me.cmdBrowseFileLocalModel.Name = "cmdBrowseFileLocalModel"
        Me.cmdBrowseFileLocalModel.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalModel.TabIndex = 107
        Me.cmdBrowseFileLocalModel.Text = "..."
        '
        'cmdPutScr
        '
        Me.cmdPutScr.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPutScr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPutScr.Location = New System.Drawing.Point(6, 17)
        Me.cmdPutScr.Name = "cmdPutScr"
        Me.cmdPutScr.Size = New System.Drawing.Size(32, 21)
        Me.cmdPutScr.TabIndex = 204
        Me.cmdPutScr.Text = "->"
        '
        'btnBrowseFileLocalDML
        '
        Me.btnBrowseFileLocalDML.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnBrowseFileLocalDML.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowseFileLocalDML.Location = New System.Drawing.Point(483, 135)
        Me.btnBrowseFileLocalDML.Name = "btnBrowseFileLocalDML"
        Me.btnBrowseFileLocalDML.Size = New System.Drawing.Size(25, 21)
        Me.btnBrowseFileLocalDML.TabIndex = 206
        Me.btnBrowseFileLocalDML.Text = "..."
        '
        'txtLocalDMLDir
        '
        Me.txtLocalDMLDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDMLDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLocalDMLDir.Location = New System.Drawing.Point(54, 136)
        Me.txtLocalDMLDir.Name = "txtLocalDMLDir"
        Me.txtLocalDMLDir.Size = New System.Drawing.Size(423, 20)
        Me.txtLocalDMLDir.TabIndex = 208
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 139)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(31, 14)
        Me.Label9.TabIndex = 209
        Me.Label9.Text = "DML"
        '
        'gbName
        '
        Me.gbName.Controls.Add(Me.gbDesc)
        Me.gbName.Controls.Add(Me.txtEnvironmentName)
        Me.gbName.Controls.Add(Me.Label3)
        Me.gbName.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbName.ForeColor = System.Drawing.Color.White
        Me.gbName.Location = New System.Drawing.Point(3, 3)
        Me.gbName.Name = "gbName"
        Me.gbName.Size = New System.Drawing.Size(343, 124)
        Me.gbName.TabIndex = 210
        Me.gbName.TabStop = False
        Me.gbName.Text = "Environment Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtEnvironmentDesc)
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 45)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(331, 70)
        Me.gbDesc.TabIndex = 42
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbScrDir
        '
        Me.gbScrDir.Controls.Add(Me.gbFTP)
        Me.gbScrDir.Controls.Add(Me.cmdBrowseFileLocalScript)
        Me.gbScrDir.Controls.Add(Me.txtLocalScriptDir)
        Me.gbScrDir.Controls.Add(Me.Label1)
        Me.gbScrDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbScrDir.ForeColor = System.Drawing.Color.White
        Me.gbScrDir.Location = New System.Drawing.Point(3, 132)
        Me.gbScrDir.Name = "gbScrDir"
        Me.gbScrDir.Size = New System.Drawing.Size(566, 68)
        Me.gbScrDir.TabIndex = 211
        Me.gbScrDir.TabStop = False
        Me.gbScrDir.Text = "Local Script/Project Directory"
        '
        'gbFTP
        '
        Me.gbFTP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFTP.Controls.Add(Me.cmdPutScr)
        Me.gbFTP.ForeColor = System.Drawing.Color.White
        Me.gbFTP.Location = New System.Drawing.Point(514, 12)
        Me.gbFTP.Name = "gbFTP"
        Me.gbFTP.Size = New System.Drawing.Size(46, 47)
        Me.gbFTP.TabIndex = 205
        Me.gbFTP.TabStop = False
        Me.gbFTP.Text = "FTP"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label4)
        Me.GroupBox5.Controls.Add(Me.GroupBox3)
        Me.GroupBox5.Controls.Add(Me.Label7)
        Me.GroupBox5.Controls.Add(Me.txtLocalCobolDir)
        Me.GroupBox5.Controls.Add(Me.Label6)
        Me.GroupBox5.Controls.Add(Me.cmdBrowseFileLocalModel)
        Me.GroupBox5.Controls.Add(Me.Label9)
        Me.GroupBox5.Controls.Add(Me.txtLocalDBDDir)
        Me.GroupBox5.Controls.Add(Me.txtLocalDDLDir)
        Me.GroupBox5.Controls.Add(Me.Label2)
        Me.GroupBox5.Controls.Add(Me.txtLocalDMLDir)
        Me.GroupBox5.Controls.Add(Me.Label5)
        Me.GroupBox5.Controls.Add(Me.txtLocalDTDDir)
        Me.GroupBox5.Controls.Add(Me.btnBrowseFileLocalDML)
        Me.GroupBox5.Controls.Add(Me.txtLocalCDir)
        Me.GroupBox5.Controls.Add(Me.Label8)
        Me.GroupBox5.Controls.Add(Me.cmdBrowseFileLocalDTD)
        Me.GroupBox5.Controls.Add(Me.cmdBrowseFileLocalDDL)
        Me.GroupBox5.Controls.Add(Me.cmdBrowseFileLocalCobol)
        Me.GroupBox5.Controls.Add(Me.cmdBrowseFileLocalC)
        Me.GroupBox5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox5.ForeColor = System.Drawing.Color.White
        Me.GroupBox5.Location = New System.Drawing.Point(3, 205)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(566, 233)
        Me.GroupBox5.TabIndex = 212
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Local Description Directories"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(19, 205)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(528, 13)
        Me.Label4.TabIndex = 214
        Me.Label4.Text = "*** Warning: Changes here will result in Changes to the Corresponding Description" & _
            " Paths ***"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.btnFTPdtd)
        Me.GroupBox3.Controls.Add(Me.btnFTPdbd)
        Me.GroupBox3.Controls.Add(Me.btnFTPddl)
        Me.GroupBox3.Controls.Add(Me.btnFTPCobol)
        Me.GroupBox3.Controls.Add(Me.btnFTPc)
        Me.GroupBox3.Controls.Add(Me.btnFTPdml)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.ForeColor = System.Drawing.Color.White
        Me.GroupBox3.Location = New System.Drawing.Point(514, 15)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(46, 179)
        Me.GroupBox3.TabIndex = 210
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "FTP"
        '
        'btnFTPdtd
        '
        Me.btnFTPdtd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPdtd.Location = New System.Drawing.Point(8, 146)
        Me.btnFTPdtd.Name = "btnFTPdtd"
        Me.btnFTPdtd.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPdtd.TabIndex = 108
        Me.btnFTPdtd.Text = "<-"
        '
        'btnFTPdbd
        '
        Me.btnFTPdbd.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFTPdbd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPdbd.Location = New System.Drawing.Point(8, 44)
        Me.btnFTPdbd.Name = "btnFTPdbd"
        Me.btnFTPdbd.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPdbd.TabIndex = 205
        Me.btnFTPdbd.Text = "<-"
        '
        'btnFTPddl
        '
        Me.btnFTPddl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPddl.Location = New System.Drawing.Point(8, 68)
        Me.btnFTPddl.Name = "btnFTPddl"
        Me.btnFTPddl.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPddl.TabIndex = 109
        Me.btnFTPddl.Text = "<-"
        '
        'btnFTPCobol
        '
        Me.btnFTPCobol.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPCobol.Location = New System.Drawing.Point(8, 16)
        Me.btnFTPCobol.Name = "btnFTPCobol"
        Me.btnFTPCobol.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPCobol.TabIndex = 110
        Me.btnFTPCobol.Text = "<-"
        '
        'btnFTPc
        '
        Me.btnFTPc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPc.Location = New System.Drawing.Point(8, 94)
        Me.btnFTPc.Name = "btnFTPc"
        Me.btnFTPc.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPc.TabIndex = 111
        Me.btnFTPc.Text = "<-"
        '
        'btnFTPdml
        '
        Me.btnFTPdml.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFTPdml.Location = New System.Drawing.Point(8, 120)
        Me.btnFTPdml.Name = "btnFTPdml"
        Me.btnFTPdml.Size = New System.Drawing.Size(32, 21)
        Me.btnFTPdml.TabIndex = 207
        Me.btnFTPdml.Text = "<-"
        '
        'FBD1
        '
        '
        'gbRelDir
        '
        Me.gbRelDir.Controls.Add(Me.btnRelDir)
        Me.gbRelDir.Controls.Add(Me.txtRelDir)
        Me.gbRelDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbRelDir.ForeColor = System.Drawing.Color.White
        Me.gbRelDir.Location = New System.Drawing.Point(3, 133)
        Me.gbRelDir.Name = "gbRelDir"
        Me.gbRelDir.Size = New System.Drawing.Size(505, 56)
        Me.gbRelDir.TabIndex = 213
        Me.gbRelDir.TabStop = False
        Me.gbRelDir.Text = "Default Directory (Relative Path)"
        Me.gbRelDir.Visible = False
        '
        'btnRelDir
        '
        Me.btnRelDir.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.btnRelDir.Location = New System.Drawing.Point(459, 18)
        Me.btnRelDir.Name = "btnRelDir"
        Me.btnRelDir.Size = New System.Drawing.Size(34, 23)
        Me.btnRelDir.TabIndex = 1
        Me.btnRelDir.Text = ".."
        Me.btnRelDir.UseVisualStyleBackColor = False
        '
        'txtRelDir
        '
        Me.txtRelDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRelDir.Location = New System.Drawing.Point(9, 20)
        Me.txtRelDir.Name = "txtRelDir"
        Me.txtRelDir.Size = New System.Drawing.Size(438, 20)
        Me.txtRelDir.TabIndex = 0
        '
        'ctlEnvironment
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.gbScrDir)
        Me.Controls.Add(Me.gbRelDir)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.gbName)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlEnvironment"
        Me.Size = New System.Drawing.Size(577, 539)
        Me.gbName.ResumeLayout(False)
        Me.gbName.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbScrDir.ResumeLayout(False)
        Me.gbScrDir.PerformLayout()
        Me.gbFTP.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.gbRelDir.ResumeLayout(False)
        Me.gbRelDir.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form events"

    Private Sub StartLoad()
        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtEnvironmentName.Enabled = True '//added on 7/24 to disable object name editing

        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsEnvironment

        ClearControls(Me.Controls)
    End Sub

    Private Sub EndLoad()
        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False
    End Sub

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

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEnvironmentName.TextChanged, txtEnvironmentDesc.TextChanged, txtLocalDTDDir.TextChanged, txtLocalDDLDir.TextChanged, txtLocalCobolDir.TextChanged, txtLocalCDir.TextChanged, txtLocalScriptDir.TextChanged, txtLocalDBDDir.TextChanged, txtLocalDMLDir.TextChanged, txtRelDir.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_Environment)
    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtEnvironmentDesc.KeyDown, txtEnvironmentName.KeyDown, txtLocalCDir.KeyDown, txtLocalCobolDir.KeyDown, txtLocalDDLDir.KeyDown, txtLocalDTDDir.KeyDown, txtLocalDBDDir.KeyDown, txtLocalScriptDir.KeyDown, Me.KeyDown, btnBrowseFileLocalDML.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    '//new by npatel on 8/13/05
    Private Sub cmdBrowseLocalDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFileLocalDTD.Click, cmdBrowseFileLocalDDL.Click, cmdBrowseFileLocalCobol.Click, cmdBrowseFileLocalC.Click, cmdBrowseFileLocalScript.Click, btnBrowseFileLocalDML.Click, cmdBrowseFileLocalModel.Click, btnRelDir.Click

        Dim DefaultDir As String
        'Dim testdir As String

        If objThis.DefaultStrDir <> "" Then
            DefaultDir = objThis.DefaultStrDir
        Else
            DefaultDir = ""
        End If
        'testdir = System.Environment.GetEnvironmentVariable("PATH_TEST")
        'MsgBox("Test directory is:   " & testdir, MsgBoxStyle.OkOnly)

        Dim DTDdir As String = ""
        Dim DTDpre As String = ""
        Dim Cdir As String = ""
        Dim Cpre As String = ""
        Dim Coboldir As String = ""
        Dim Cobolpre As String = ""
        Dim DDLdir As String = ""
        Dim DDLpre As String = ""
        Dim DMLdir As String = ""
        Dim DMLpre As String = ""
        Dim Scriptdir As String = ""
        Dim Scriptpre As String = ""
        'model Dir is actually DBD dir
        Dim Modeldir As String = ""
        Dim Modelpre As String = ""
        'relative directory/base dir/default dir
        Dim RelDir As String = ""
        Dim RelPre As String = ""

        'Dim TempArr As Char()
        'Dim TempStr As String

        Dim preIDX As Integer
        Dim pathIDX As Integer

        '//new 8/15/05
        Select Case CType(sender, Button).Name
            Case "cmdBrowseFileLocalDTD"
                If txtLocalDTDDir.Text.Contains("%") = True Then
                    preIDX = txtLocalDTDDir.Text.IndexOf("%")
                    pathIDX = txtLocalDTDDir.Text.LastIndexOf("%")
                    DTDpre = Strings.Mid(txtLocalDTDDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    DTDdir = Strings.Right(txtLocalDTDDir.Text, txtLocalDTDDir.Text.Length - pathIDX - 1)
                    DTDdir = System.Environment.GetEnvironmentVariable(DTDpre) & DTDdir
                Else
                    DTDdir = txtLocalDTDDir.Text
                End If
            Case "cmdBrowseFileLocalC"
                If txtLocalCDir.Text.Contains("%") = True Then
                    preIDX = txtLocalCDir.Text.IndexOf("%")
                    pathIDX = txtLocalCDir.Text.LastIndexOf("%")
                    Cpre = Strings.Mid(txtLocalCDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    Cdir = Strings.Right(txtLocalCDir.Text, txtLocalCDir.Text.Length - pathIDX - 1)
                    Cdir = System.Environment.GetEnvironmentVariable(Cpre) & Cdir
                Else
                    Cdir = txtLocalCDir.Text
                End If
            Case "cmdBrowseFileLocalCobol"
                If txtLocalCobolDir.Text.Contains("%") = True Then
                    preIDX = txtLocalCobolDir.Text.IndexOf("%")
                    pathIDX = txtLocalCobolDir.Text.LastIndexOf("%")
                    Cobolpre = Strings.Mid(txtLocalCobolDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    Coboldir = Strings.Right(txtLocalCobolDir.Text, txtLocalCobolDir.Text.Length - pathIDX - 1)
                    Coboldir = System.Environment.GetEnvironmentVariable(Cobolpre) & Coboldir
                Else
                    Coboldir = txtLocalCobolDir.Text
                End If
            Case "cmdBrowseFileLocalDDL"
                If txtLocalDDLDir.Text.Contains("%") = True Then
                    preIDX = txtLocalDDLDir.Text.IndexOf("%")
                    pathIDX = txtLocalDDLDir.Text.LastIndexOf("%")
                    DDLpre = Strings.Mid(txtLocalDDLDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    DDLdir = Strings.Right(txtLocalDDLDir.Text, txtLocalDDLDir.Text.Length - pathIDX - 1)
                    DDLdir = System.Environment.GetEnvironmentVariable(DDLpre) & DDLdir
                Else
                    DDLdir = txtLocalDDLDir.Text
                End If
            Case "btnBrowseFileLocalDML"
                If txtLocalDMLDir.Text.Contains("%") = True Then
                    preIDX = txtLocalDMLDir.Text.IndexOf("%")
                    pathIDX = txtLocalDMLDir.Text.LastIndexOf("%")
                    DMLpre = Strings.Mid(txtLocalDMLDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    DMLdir = Strings.Right(txtLocalDMLDir.Text, txtLocalDMLDir.Text.Length - pathIDX - 1)
                    DMLdir = System.Environment.GetEnvironmentVariable(DMLpre) & DMLdir
                Else
                    DMLdir = txtLocalDMLDir.Text
                End If
            Case "cmdBrowseFileLocalScript"
                If txtLocalScriptDir.Text.Contains("%") = True Then
                    preIDX = txtLocalScriptDir.Text.IndexOf("%")
                    pathIDX = txtLocalScriptDir.Text.LastIndexOf("%")
                    Scriptpre = Strings.Mid(txtLocalScriptDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    Scriptdir = Strings.Right(txtLocalScriptDir.Text, txtLocalScriptDir.Text.Length - pathIDX - 1)
                    Scriptdir = System.Environment.GetEnvironmentVariable(Scriptpre) & Scriptdir
                Else
                    Scriptdir = txtLocalScriptDir.Text
                End If
            Case "cmdBrowseFileLocalModel"
                If txtLocalDBDDir.Text.Contains("%") = True Then
                    preIDX = txtLocalDBDDir.Text.IndexOf("%")
                    pathIDX = txtLocalDBDDir.Text.LastIndexOf("%")
                    Modelpre = Strings.Mid(txtLocalDBDDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    Modeldir = Strings.Right(txtLocalDBDDir.Text, txtLocalDBDDir.Text.Length - pathIDX - 1)
                    Modeldir = System.Environment.GetEnvironmentVariable(Modelpre) & Modeldir
                Else
                    Modeldir = txtLocalDBDDir.Text
                End If
            Case "btnRelDir"
                If txtRelDir.Text.Contains("%") = True Then
                    preIDX = txtRelDir.Text.IndexOf("%")
                    pathIDX = txtRelDir.Text.LastIndexOf("%")
                    RelPre = Strings.Mid(txtRelDir.Text, preIDX + 2, pathIDX - preIDX - 1)
                    RelDir = Strings.Right(txtRelDir.Text, txtRelDir.Text.Length - pathIDX - 1)
                    RelDir = System.Environment.GetEnvironmentVariable(RelPre) & RelDir
                Else
                    RelDir = txtRelDir.Text
                End If
        End Select

        Select Case CType(sender, Button).Name
            Case "cmdBrowseFileLocalDTD"
                If DTDdir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = DTDdir
                End If
            Case "cmdBrowseFileLocalC"
                If Cdir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = Cdir
                End If
            Case "cmdBrowseFileLocalCobol"
                If Coboldir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = Coboldir
                End If
            Case "cmdBrowseFileLocalDDL"
                If DDLdir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = DDLdir
                End If
            Case "btnBrowseFileLocalDML"
                If DMLdir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = DMLdir
                End If
            Case "cmdBrowseFileLocalScript"
                If Scriptdir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = Scriptdir
                End If
            Case "cmdBrowseFileLocalModel"
                If Modeldir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = Modeldir
                End If
            Case "btnRelDir"
                If RelDir = "" Then
                    FBD1.SelectedPath = DefaultDir
                Else
                    FBD1.SelectedPath = RelDir
                End If
        End Select

        If FBD1.ShowDialog() = DialogResult.OK Then
            Select Case CType(sender, Button).Name
                Case "cmdBrowseFileLocalDTD"
                    txtLocalDTDDir.Text = FBD1.SelectedPath
                    objThis.LocalDTDDir = FBD1.SelectedPath
                Case "cmdBrowseFileLocalC"
                    txtLocalCDir.Text = FBD1.SelectedPath
                    objThis.LocalCDir = FBD1.SelectedPath
                Case "cmdBrowseFileLocalCobol"
                    txtLocalCobolDir.Text = FBD1.SelectedPath
                    objThis.LocalCobolDir = FBD1.SelectedPath
                Case "cmdBrowseFileLocalDDL"
                    txtLocalDDLDir.Text = FBD1.SelectedPath
                    objThis.LocalDDLDir = FBD1.SelectedPath
                Case "btnBrowseFileLocalDML"
                    txtLocalDMLDir.Text = FBD1.SelectedPath
                    objThis.LocalDMLDir = FBD1.SelectedPath
                Case "cmdBrowseFileLocalScript"
                    txtLocalScriptDir.Text = FBD1.SelectedPath
                    objThis.LocalScriptDir = FBD1.SelectedPath
                Case "cmdBrowseFileLocalModel"
                    txtLocalDBDDir.Text = FBD1.SelectedPath
                    objThis.LocalDBDDir = FBD1.SelectedPath
                Case "btnRelDir"
                    txtRelDir.Text = FBD1.SelectedPath
                    objThis.RelDir = FBD1.SelectedPath
            End Select
            objThis.DefaultStrDir = FBD1.SelectedPath
        End If

        objThis.PathChange = True

    End Sub

#End Region

#Region "Text box click events"

    Private Sub cmdGetDTD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPdtd.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient

        destination = FTPClient.BrowseFile(Me.txtLocalDTDDir.Text, ".dtd", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDTDDir = destination
            txtLocalDTDDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetDDL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPddl.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient

        destination = FTPClient.BrowseFile(Me.txtLocalDDLDir.Text, ".ddl", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDDLDir = destination
            txtLocalDDLDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetCOB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPCobol.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient

        destination = FTPClient.BrowseFile(Me.txtLocalCobolDir.Text, ".cob", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalCobolDir = destination
            txtLocalCobolDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPc.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient

        destination = FTPClient.BrowseFile(Me.txtLocalCDir.Text, ".h", modDeclares.enumCalledFrom.BY_ENVIRONMENT)
        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalCDir = destination
            txtLocalCDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetDML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPdml.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient

        destination = FTPClient.BrowseFile(Me.txtLocalDMLDir.Text, ".dml", modDeclares.enumCalledFrom.BY_ENVIRONMENT)
        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDMLDir = destination
            txtLocalDMLDir.Text = destination
        End If

    End Sub

    'Private Sub cmdBrowseFileLocalModel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFileLocalModel.Click

    '    Dim DefaultDir As String

    '    If objThis.DefaultStrDir <> "" Then
    '        DefaultDir = objThis.DefaultStrDir
    '    Else
    '        DefaultDir = ""
    '    End If

    '    If txtLocalModelDir.Text = "" Then
    '        dlgBrowseFolder.SelectedPath = DefaultDir 'TrunkFileName(LocalModelFolderPath)
    '    Else
    '        dlgBrowseFolder.SelectedPath = txtLocalModelDir.Text
    '    End If

    '    If dlgBrowseFolder.ShowDialog() = DialogResult.OK Then
    '        txtLocalModelDir.Text = dlgBrowseFolder.SelectedPath
    '        objThis.LocalDBDDir = dlgBrowseFolder.SelectedPath
    '        'LocalModelFolderPath = dlgBrowseFolder.SelectedPath
    '    End If

    'End Sub

    Private Sub cmdPutScr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPutScr.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalScriptDir.Text, ".inl", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalScriptDir = destination
            txtLocalScriptDir.Text = destination
        End If

    End Sub

    Private Sub cmdPutMdl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTPdbd.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalScriptDir.Text, ".ddl", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalScriptDir = destination
            txtLocalScriptDir.Text = destination
        End If

    End Sub

#End Region

    Public Function Save() As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            '// First Check Name Validity before Saving
            If ValidateNewName(txtEnvironmentName.Text) = False Then
                Save = False
                Me.Cursor = Cursors.Default
                Exit Function
            End If
            '// Check to see if the object has been renamed
            If objThis.EnvironmentName <> txtEnvironmentName.Text Then
                objThis.IsRenamed = RenameEnvironment(objThis, txtEnvironmentName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtEnvironmentName.Text = objThis.EnvironmentName
            Else
                objThis.EnvironmentName = txtEnvironmentName.Text
            End If

            objThis.EnvironmentDescription = txtEnvironmentDesc.Text
            objThis.LocalDTDDir = txtLocalDTDDir.Text
            objThis.LocalDDLDir = txtLocalDDLDir.Text
            objThis.LocalCobolDir = txtLocalCobolDir.Text
            objThis.LocalCDir = txtLocalCDir.Text
            objThis.LocalDMLDir = txtLocalDMLDir.Text
            objThis.LocalScriptDir = txtLocalScriptDir.Text
            objThis.LocalDBDDir = txtLocalDBDDir.Text
            objThis.RelDir = txtRelDir.Text

            objThis.SetStructureDir()

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
            LogError(ex, "ctlEnv Save")
            Save = False
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Public Function EditObj(ByVal obj As INode) As clsEnvironment

        Try
            Me.Cursor = Cursors.WaitCursor

            IsNewObj = False

            StartLoad()

            objThis = obj '//Load the form env object
            objThis.LoadMe()

            UpdateFields()

            EditObj = objThis

            EndLoad()
            Me.Cursor = Cursors.Default

        Catch ex As Exception
            LogError(ex, "ctlEnvironment EditObj")
            Me.Cursor = Cursors.Default
            EditObj = Nothing
        End Try

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtEnvironmentName.Text = objThis.EnvironmentName
        txtEnvironmentDesc.Text = objThis.EnvironmentDescription
        txtLocalDTDDir.Text = objThis.LocalDTDDir
        txtLocalDDLDir.Text = objThis.LocalDDLDir
        txtLocalCobolDir.Text = objThis.LocalCobolDir
        txtLocalCDir.Text = objThis.LocalCDir
        txtLocalDMLDir.Text = objThis.LocalDMLDir
        txtLocalScriptDir.Text = objThis.LocalScriptDir
        txtLocalDBDDir.Text = objThis.LocalDBDDir
        txtRelDir.Text = objThis.RelDir

    End Function

    'Function TrunkFileName(ByVal filePath As String) As String

    '    Dim i As Integer = filePath.LastIndexOf("\")

    '    If filePath.IndexOf(".") > 0 Then
    '        filePath = filePath.Substring(0, i)
    '    End If

    '    Return (filePath)

    'End Function

End Class
