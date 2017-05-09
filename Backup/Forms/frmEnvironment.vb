Public Class frmEnvironment
    Inherits SQDStudio.frmBlank

    Public objThis As New clsEnvironment

    Dim IsNewObj As Boolean

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
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtEnvironmentName As System.Windows.Forms.TextBox
    Friend WithEvents txtEnvironmentDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtLocalDTDDir As System.Windows.Forms.TextBox
    Friend WithEvents txtLocalDDLDir As System.Windows.Forms.TextBox
    Friend WithEvents txtLocalCobolDir As System.Windows.Forms.TextBox
    Friend WithEvents txtLocalCDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseFileLocalC As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalCobol As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalDDL As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalDTD As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalScript As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFileLocalModel As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents txtLocalScriptDir As System.Windows.Forms.TextBox
    Friend WithEvents txtLocalDBDDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseFieldLocalScript As System.Windows.Forms.Button
    Friend WithEvents cmdBrowseFieldlocalDBD As System.Windows.Forms.Button
    Friend WithEvents cmdGetDTD As System.Windows.Forms.Button
    Friend WithEvents cmdGetDDL As System.Windows.Forms.Button
    Friend WithEvents cmdGetCobol As System.Windows.Forms.Button
    Friend WithEvents cmdGetC As System.Windows.Forms.Button
    Friend WithEvents cmdPutScr As System.Windows.Forms.Button
    Friend WithEvents cmdGetDBD As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents txtLocalDMLDir As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowseLocalDMLDir As System.Windows.Forms.Button
    Friend WithEvents cmdGetDML As System.Windows.Forms.Button
    Friend WithEvents gbEnvProp As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
    Friend WithEvents gbScriptDir As System.Windows.Forms.GroupBox
    Friend WithEvents gbscriptFTP As System.Windows.Forms.GroupBox
    Friend WithEvents gbLocalDirs As System.Windows.Forms.GroupBox
    Friend WithEvents gbFTP As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtEnvironmentName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtEnvironmentDesc = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtLocalDTDDir = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtLocalDDLDir = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtLocalCobolDir = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtLocalCDir = New System.Windows.Forms.TextBox
        Me.cmdBrowseFileLocalC = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalCobol = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalDDL = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalDTD = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalScript = New System.Windows.Forms.Button
        Me.cmdBrowseFileLocalModel = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtLocalScriptDir = New System.Windows.Forms.TextBox
        Me.txtLocalDBDDir = New System.Windows.Forms.TextBox
        Me.cmdBrowseFieldLocalScript = New System.Windows.Forms.Button
        Me.cmdBrowseFieldlocalDBD = New System.Windows.Forms.Button
        Me.cmdGetDTD = New System.Windows.Forms.Button
        Me.cmdGetDDL = New System.Windows.Forms.Button
        Me.cmdGetCobol = New System.Windows.Forms.Button
        Me.cmdGetC = New System.Windows.Forms.Button
        Me.cmdPutScr = New System.Windows.Forms.Button
        Me.cmdGetDBD = New System.Windows.Forms.Button
        Me.Label11 = New System.Windows.Forms.Label
        Me.txtLocalDMLDir = New System.Windows.Forms.TextBox
        Me.cmdBrowseLocalDMLDir = New System.Windows.Forms.Button
        Me.cmdGetDML = New System.Windows.Forms.Button
        Me.gbEnvProp = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbScriptDir = New System.Windows.Forms.GroupBox
        Me.gbscriptFTP = New System.Windows.Forms.GroupBox
        Me.gbLocalDirs = New System.Windows.Forms.GroupBox
        Me.gbFTP = New System.Windows.Forms.GroupBox
        Me.Panel1.SuspendLayout()
        Me.gbEnvProp.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.gbScriptDir.SuspendLayout()
        Me.gbscriptFTP.SuspendLayout()
        Me.gbLocalDirs.SuspendLayout()
        Me.gbFTP.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(554, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 476)
        Me.GroupBox1.Size = New System.Drawing.Size(556, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(272, 499)
        Me.cmdOk.TabIndex = 21
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(368, 499)
        Me.cmdCancel.TabIndex = 22
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(464, 499)
        Me.cmdHelp.TabIndex = 23
        '
        'Label1
        '
        Me.Label1.Text = "Environment definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(478, 39)
        Me.Label2.Text = "Enter an environment name that is unique within a project. An environment is used" & _
            " to distinguish between a development, test, and production environments."
        '
        'txtEnvironmentName
        '
        Me.txtEnvironmentName.Location = New System.Drawing.Point(50, 19)
        Me.txtEnvironmentName.MaxLength = 128
        Me.txtEnvironmentName.Name = "txtEnvironmentName"
        Me.txtEnvironmentName.Size = New System.Drawing.Size(315, 20)
        Me.txtEnvironmentName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Name"
        '
        'txtEnvironmentDesc
        '
        Me.txtEnvironmentDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtEnvironmentDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtEnvironmentDesc.MaxLength = 1000
        Me.txtEnvironmentDesc.Multiline = True
        Me.txtEnvironmentDesc.Name = "txtEnvironmentDesc"
        Me.txtEnvironmentDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtEnvironmentDesc.Size = New System.Drawing.Size(356, 56)
        Me.txtEnvironmentDesc.TabIndex = 20
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(6, 166)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(28, 14)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "DTD"
        '
        'txtLocalDTDDir
        '
        Me.txtLocalDTDDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDTDDir.Location = New System.Drawing.Point(50, 163)
        Me.txtLocalDTDDir.MaxLength = 255
        Me.txtLocalDTDDir.Name = "txtLocalDTDDir"
        Me.txtLocalDTDDir.Size = New System.Drawing.Size(386, 20)
        Me.txtLocalDTDDir.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(28, 14)
        Me.Label5.TabIndex = 8
        Me.Label5.Text = "DDL"
        '
        'txtLocalDDLDir
        '
        Me.txtLocalDDLDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDDLDir.Location = New System.Drawing.Point(51, 85)
        Me.txtLocalDDLDir.MaxLength = 255
        Me.txtLocalDDLDir.Name = "txtLocalDDLDir"
        Me.txtLocalDDLDir.Size = New System.Drawing.Size(385, 20)
        Me.txtLocalDDLDir.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 36)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(39, 14)
        Me.Label6.TabIndex = 8
        Me.Label6.Text = "Cobol"
        '
        'txtLocalCobolDir
        '
        Me.txtLocalCobolDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalCobolDir.Location = New System.Drawing.Point(51, 33)
        Me.txtLocalCobolDir.MaxLength = 255
        Me.txtLocalCobolDir.Name = "txtLocalCobolDir"
        Me.txtLocalCobolDir.Size = New System.Drawing.Size(385, 20)
        Me.txtLocalCobolDir.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 114)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(15, 14)
        Me.Label8.TabIndex = 12
        Me.Label8.Text = "C"
        '
        'txtLocalCDir
        '
        Me.txtLocalCDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalCDir.Location = New System.Drawing.Point(51, 111)
        Me.txtLocalCDir.MaxLength = 255
        Me.txtLocalCDir.Name = "txtLocalCDir"
        Me.txtLocalCDir.Size = New System.Drawing.Size(385, 20)
        Me.txtLocalCDir.TabIndex = 11
        '
        'cmdBrowseFileLocalC
        '
        Me.cmdBrowseFileLocalC.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalC.Location = New System.Drawing.Point(442, 110)
        Me.cmdBrowseFileLocalC.Name = "cmdBrowseFileLocalC"
        Me.cmdBrowseFileLocalC.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalC.TabIndex = 12
        Me.cmdBrowseFileLocalC.Text = "..."
        '
        'cmdBrowseFileLocalCobol
        '
        Me.cmdBrowseFileLocalCobol.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalCobol.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalCobol.Location = New System.Drawing.Point(442, 32)
        Me.cmdBrowseFileLocalCobol.Name = "cmdBrowseFileLocalCobol"
        Me.cmdBrowseFileLocalCobol.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalCobol.TabIndex = 9
        Me.cmdBrowseFileLocalCobol.Text = "..."
        '
        'cmdBrowseFileLocalDDL
        '
        Me.cmdBrowseFileLocalDDL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalDDL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalDDL.Location = New System.Drawing.Point(442, 84)
        Me.cmdBrowseFileLocalDDL.Name = "cmdBrowseFileLocalDDL"
        Me.cmdBrowseFileLocalDDL.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalDDL.TabIndex = 6
        Me.cmdBrowseFileLocalDDL.Text = "..."
        '
        'cmdBrowseFileLocalDTD
        '
        Me.cmdBrowseFileLocalDTD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFileLocalDTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFileLocalDTD.Location = New System.Drawing.Point(442, 162)
        Me.cmdBrowseFileLocalDTD.Name = "cmdBrowseFileLocalDTD"
        Me.cmdBrowseFileLocalDTD.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFileLocalDTD.TabIndex = 3
        Me.cmdBrowseFileLocalDTD.Text = "..."
        '
        'cmdBrowseFileLocalScript
        '
        Me.cmdBrowseFileLocalScript.Location = New System.Drawing.Point(0, 0)
        Me.cmdBrowseFileLocalScript.Name = "cmdBrowseFileLocalScript"
        Me.cmdBrowseFileLocalScript.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowseFileLocalScript.TabIndex = 0
        '
        'cmdBrowseFileLocalModel
        '
        Me.cmdBrowseFileLocalModel.Location = New System.Drawing.Point(0, 0)
        Me.cmdBrowseFileLocalModel.Name = "cmdBrowseFileLocalModel"
        Me.cmdBrowseFileLocalModel.Size = New System.Drawing.Size(75, 23)
        Me.cmdBrowseFileLocalModel.TabIndex = 0
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(6, 32)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(39, 14)
        Me.Label9.TabIndex = 110
        Me.Label9.Text = "Script"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(6, 61)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(28, 14)
        Me.Label10.TabIndex = 111
        Me.Label10.Text = "DBD"
        '
        'txtLocalScriptDir
        '
        Me.txtLocalScriptDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalScriptDir.Location = New System.Drawing.Point(51, 29)
        Me.txtLocalScriptDir.Name = "txtLocalScriptDir"
        Me.txtLocalScriptDir.Size = New System.Drawing.Size(385, 20)
        Me.txtLocalScriptDir.TabIndex = 14
        '
        'txtLocalDBDDir
        '
        Me.txtLocalDBDDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDBDDir.Location = New System.Drawing.Point(51, 59)
        Me.txtLocalDBDDir.Name = "txtLocalDBDDir"
        Me.txtLocalDBDDir.Size = New System.Drawing.Size(385, 20)
        Me.txtLocalDBDDir.TabIndex = 17
        '
        'cmdBrowseFieldLocalScript
        '
        Me.cmdBrowseFieldLocalScript.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFieldLocalScript.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFieldLocalScript.Location = New System.Drawing.Point(442, 27)
        Me.cmdBrowseFieldLocalScript.Name = "cmdBrowseFieldLocalScript"
        Me.cmdBrowseFieldLocalScript.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFieldLocalScript.TabIndex = 15
        Me.cmdBrowseFieldLocalScript.Text = "..."
        '
        'cmdBrowseFieldlocalDBD
        '
        Me.cmdBrowseFieldlocalDBD.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseFieldlocalDBD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseFieldlocalDBD.Location = New System.Drawing.Point(442, 57)
        Me.cmdBrowseFieldlocalDBD.Name = "cmdBrowseFieldlocalDBD"
        Me.cmdBrowseFieldlocalDBD.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseFieldlocalDBD.TabIndex = 18
        Me.cmdBrowseFieldlocalDBD.Text = "..."
        '
        'cmdGetDTD
        '
        Me.cmdGetDTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetDTD.Location = New System.Drawing.Point(6, 149)
        Me.cmdGetDTD.Name = "cmdGetDTD"
        Me.cmdGetDTD.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetDTD.TabIndex = 4
        Me.cmdGetDTD.Text = "<-"
        '
        'cmdGetDDL
        '
        Me.cmdGetDDL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetDDL.Location = New System.Drawing.Point(6, 71)
        Me.cmdGetDDL.Name = "cmdGetDDL"
        Me.cmdGetDDL.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetDDL.TabIndex = 7
        Me.cmdGetDDL.Text = "<-"
        '
        'cmdGetCobol
        '
        Me.cmdGetCobol.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetCobol.Location = New System.Drawing.Point(6, 19)
        Me.cmdGetCobol.Name = "cmdGetCobol"
        Me.cmdGetCobol.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetCobol.TabIndex = 10
        Me.cmdGetCobol.Text = "<-"
        '
        'cmdGetC
        '
        Me.cmdGetC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetC.Location = New System.Drawing.Point(6, 97)
        Me.cmdGetC.Name = "cmdGetC"
        Me.cmdGetC.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetC.TabIndex = 13
        Me.cmdGetC.Text = "<-"
        '
        'cmdPutScr
        '
        Me.cmdPutScr.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPutScr.Location = New System.Drawing.Point(6, 19)
        Me.cmdPutScr.Name = "cmdPutScr"
        Me.cmdPutScr.Size = New System.Drawing.Size(32, 21)
        Me.cmdPutScr.TabIndex = 16
        Me.cmdPutScr.Text = "->"
        '
        'cmdGetDBD
        '
        Me.cmdGetDBD.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetDBD.Location = New System.Drawing.Point(6, 44)
        Me.cmdGetDBD.Name = "cmdGetDBD"
        Me.cmdGetDBD.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetDBD.TabIndex = 19
        Me.cmdGetDBD.Text = "<-"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(6, 140)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(31, 14)
        Me.Label11.TabIndex = 112
        Me.Label11.Text = "DML"
        '
        'txtLocalDMLDir
        '
        Me.txtLocalDMLDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLocalDMLDir.Location = New System.Drawing.Point(50, 137)
        Me.txtLocalDMLDir.MaxLength = 255
        Me.txtLocalDMLDir.Name = "txtLocalDMLDir"
        Me.txtLocalDMLDir.Size = New System.Drawing.Size(386, 20)
        Me.txtLocalDMLDir.TabIndex = 113
        '
        'cmdBrowseLocalDMLDir
        '
        Me.cmdBrowseLocalDMLDir.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowseLocalDMLDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseLocalDMLDir.Location = New System.Drawing.Point(442, 136)
        Me.cmdBrowseLocalDMLDir.Name = "cmdBrowseLocalDMLDir"
        Me.cmdBrowseLocalDMLDir.Size = New System.Drawing.Size(25, 21)
        Me.cmdBrowseLocalDMLDir.TabIndex = 114
        Me.cmdBrowseLocalDMLDir.Text = "..."
        '
        'cmdGetDML
        '
        Me.cmdGetDML.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdGetDML.Location = New System.Drawing.Point(6, 123)
        Me.cmdGetDML.Name = "cmdGetDML"
        Me.cmdGetDML.Size = New System.Drawing.Size(32, 21)
        Me.cmdGetDML.TabIndex = 115
        Me.cmdGetDML.Text = "<-"
        '
        'gbEnvProp
        '
        Me.gbEnvProp.Controls.Add(Me.gbDesc)
        Me.gbEnvProp.Controls.Add(Me.txtEnvironmentName)
        Me.gbEnvProp.Controls.Add(Me.Label3)
        Me.gbEnvProp.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbEnvProp.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbEnvProp.Location = New System.Drawing.Point(12, 74)
        Me.gbEnvProp.Name = "gbEnvProp"
        Me.gbEnvProp.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbEnvProp.Size = New System.Drawing.Size(374, 126)
        Me.gbEnvProp.TabIndex = 121
        Me.gbEnvProp.TabStop = False
        Me.gbEnvProp.Text = "Environment Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbDesc.Controls.Add(Me.txtEnvironmentDesc)
        Me.gbDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbDesc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbDesc.Location = New System.Drawing.Point(6, 45)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(362, 75)
        Me.gbDesc.TabIndex = 9
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'gbScriptDir
        '
        Me.gbScriptDir.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbScriptDir.Controls.Add(Me.gbscriptFTP)
        Me.gbScriptDir.Controls.Add(Me.Label9)
        Me.gbScriptDir.Controls.Add(Me.txtLocalScriptDir)
        Me.gbScriptDir.Controls.Add(Me.cmdBrowseFieldLocalScript)
        Me.gbScriptDir.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbScriptDir.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbScriptDir.Location = New System.Drawing.Point(12, 206)
        Me.gbScriptDir.Name = "gbScriptDir"
        Me.gbScriptDir.Size = New System.Drawing.Size(530, 67)
        Me.gbScriptDir.TabIndex = 122
        Me.gbScriptDir.TabStop = False
        Me.gbScriptDir.Text = "Local Script/Project Directory"
        '
        'gbscriptFTP
        '
        Me.gbscriptFTP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbscriptFTP.Controls.Add(Me.cmdPutScr)
        Me.gbscriptFTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbscriptFTP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbscriptFTP.Location = New System.Drawing.Point(473, 9)
        Me.gbscriptFTP.Name = "gbscriptFTP"
        Me.gbscriptFTP.Size = New System.Drawing.Size(45, 49)
        Me.gbscriptFTP.TabIndex = 111
        Me.gbscriptFTP.TabStop = False
        Me.gbscriptFTP.Text = "FTP"
        '
        'gbLocalDirs
        '
        Me.gbLocalDirs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbLocalDirs.Controls.Add(Me.gbFTP)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalCobolDir)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseFileLocalCobol)
        Me.gbLocalDirs.Controls.Add(Me.Label6)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseFileLocalDTD)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseLocalDMLDir)
        Me.gbLocalDirs.Controls.Add(Me.Label10)
        Me.gbLocalDirs.Controls.Add(Me.Label11)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalDMLDir)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalDBDDir)
        Me.gbLocalDirs.Controls.Add(Me.Label7)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseFieldlocalDBD)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalDTDDir)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseFileLocalDDL)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalDDLDir)
        Me.gbLocalDirs.Controls.Add(Me.cmdBrowseFileLocalC)
        Me.gbLocalDirs.Controls.Add(Me.Label5)
        Me.gbLocalDirs.Controls.Add(Me.txtLocalCDir)
        Me.gbLocalDirs.Controls.Add(Me.Label8)
        Me.gbLocalDirs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbLocalDirs.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbLocalDirs.Location = New System.Drawing.Point(12, 279)
        Me.gbLocalDirs.Name = "gbLocalDirs"
        Me.gbLocalDirs.Size = New System.Drawing.Size(530, 194)
        Me.gbLocalDirs.TabIndex = 123
        Me.gbLocalDirs.TabStop = False
        Me.gbLocalDirs.Text = "Local Description Directories"
        '
        'gbFTP
        '
        Me.gbFTP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFTP.Controls.Add(Me.cmdGetCobol)
        Me.gbFTP.Controls.Add(Me.cmdGetDBD)
        Me.gbFTP.Controls.Add(Me.cmdGetDDL)
        Me.gbFTP.Controls.Add(Me.cmdGetDTD)
        Me.gbFTP.Controls.Add(Me.cmdGetDML)
        Me.gbFTP.Controls.Add(Me.cmdGetC)
        Me.gbFTP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbFTP.ForeColor = System.Drawing.SystemColors.WindowText
        Me.gbFTP.Location = New System.Drawing.Point(473, 13)
        Me.gbFTP.Name = "gbFTP"
        Me.gbFTP.Size = New System.Drawing.Size(45, 175)
        Me.gbFTP.TabIndex = 0
        Me.gbFTP.TabStop = False
        Me.gbFTP.Text = "FTP"
        '
        'frmEnvironment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(554, 538)
        Me.Controls.Add(Me.gbLocalDirs)
        Me.Controls.Add(Me.gbScriptDir)
        Me.Controls.Add(Me.gbEnvProp)
        Me.MinimumSize = New System.Drawing.Size(479, 572)
        Me.Name = "frmEnvironment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Environment Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.gbEnvProp, 0)
        Me.Controls.SetChildIndex(Me.gbScriptDir, 0)
        Me.Controls.SetChildIndex(Me.gbLocalDirs, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbEnvProp.ResumeLayout(False)
        Me.gbEnvProp.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.gbScriptDir.ResumeLayout(False)
        Me.gbScriptDir.PerformLayout()
        Me.gbscriptFTP.ResumeLayout(False)
        Me.gbLocalDirs.ResumeLayout(False)
        Me.gbLocalDirs.PerformLayout()
        Me.gbFTP.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtEnvironmentName.TextChanged, txtEnvironmentDesc.TextChanged, txtLocalDTDDir.TextChanged, txtLocalDDLDir.TextChanged, txtLocalCobolDir.TextChanged, txtLocalCDir.TextChanged, txtLocalScriptDir.TextChanged, txtLocalDBDDir.TextChanged, txtLocalDMLDir.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False

    End Sub

    Private Sub SetDefaultName()

        Try
            Dim ProjObj As clsProject = objThis.Project
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = "Env_" & Count.ToString 'Strings.Left(ProjObj.Text, 6) & 
                For Each TestEnv As clsEnvironment In ProjObj.Environments
                    If TestEnv.EnvironmentName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtEnvironmentName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmEnv SetDefaultName")
        End Try

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

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

    '//Returns new Environment object
    Public Function NewObj() As clsEnvironment

        IsNewObj = True

        StartLoad()

        SetDefaultName()

        txtEnvironmentName.SelectAll()

        txtEnvironmentName.Focus()

        EndLoad()

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                NewObj = objThis
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                Return Nothing
        End Select

    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If objThis.ValidateNewObject(txtEnvironmentName.Text) = False Or ValidateNewName(txtEnvironmentName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.EnvironmentName = txtEnvironmentName.Text
            objThis.EnvironmentDescription = txtEnvironmentDesc.Text
            objThis.LocalDTDDir = txtLocalDTDDir.Text
            objThis.LocalDDLDir = txtLocalDDLDir.Text
            objThis.LocalCobolDir = txtLocalCobolDir.Text
            objThis.LocalCDir = txtLocalCDir.Text
            objThis.LocalDMLDir = txtLocalDMLDir.Text
            objThis.LocalScriptDir = txtLocalScriptDir.Text
            objThis.LocalDBDDir = txtLocalDBDDir.Text

            objThis.SetStructureDir()

            'LocalDTDFolderPath = txtLocalDTDDir.Text
            'LocalCFolderPath = txtLocalCDir.Text
            'LocalCOBOLFolderPath = txtLocalCobolDir.Text
            'LocalDDLFolderPath = txtLocalDDLDir.Text
            'LocalSCRIPTFolderPath = txtLocalScriptDir.Text
            'LocalModelFolderPath = txtLocalModelDir.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK

        Catch ex As Exception
            LogError(ex, "frmEnvironment cmdOk_Click")
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Add_an_Environment)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) 'Handles MyBase.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click(sender, e)
            Case Keys.Enter
                cmdOk_Click(sender, e)
            Case Keys.F1
                cmdHelp_Click(sender, e)
        End Select
    End Sub

    '//new by npatel on 8/13/05
    Private Sub cmdBrowseLocalDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFileLocalDTD.Click, cmdBrowseFileLocalDDL.Click, cmdBrowseFileLocalCobol.Click, cmdBrowseFileLocalC.Click, cmdBrowseFileLocalScript.Click, cmdBrowseFileLocalModel.Click, cmdBrowseLocalDMLDir.Click, cmdBrowseFieldlocalDBD.Click, cmdBrowseFieldLocalScript.Click


        Dim DefaultDir As String

        If objThis.DefaultStrDir <> "" Then
            DefaultDir = objThis.DefaultStrDir
        Else
            DefaultDir = ""
        End If
        '//new 8/15/05
        Select Case CType(sender, Button).Name
            Case "cmdBrowseFileLocalDTD"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseFileLocalC"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseFileLocalCobol"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseFileLocalDDL"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseLocalDMLDir"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseFileLocalScript"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "cmdBrowseFieldlocalDBD"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
            Case "btnRelDir"
                If dlgBrowseFolder.SelectedPath = Nothing Then
                    dlgBrowseFolder.SelectedPath = DefaultDir
                End If
        End Select

        If dlgBrowseFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Select Case CType(sender, Button).Name
                Case "cmdBrowseFileLocalDTD"
                    txtLocalDTDDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseFileLocalC"
                    txtLocalCDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseFileLocalCobol"
                    txtLocalCobolDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseFileLocalDDL"
                    txtLocalDDLDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseLocalDMLDir"
                    txtLocalDMLDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseFieldLocalScript"
                    txtLocalScriptDir.Text = dlgBrowseFolder.SelectedPath
                Case "cmdBrowseFieldlocalDBD"
                    txtLocalDBDDir.Text = dlgBrowseFolder.SelectedPath
                    'Case "btnRelDir"
                    '    txtRelDir.Text = dlgBrowseFolder.SelectedPath
            End Select
            objThis.DefaultStrDir = dlgBrowseFolder.SelectedPath
        End If

    End Sub

    'Private Sub cmdBrowseFieldLocalScript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFieldLocalScript.Click

    '    If dlgBrowseFolder.RootFolder = Nothing Then
    '        dlgBrowseFolder.SelectedPath = "" 'LocalSCRIPTFolderPath
    '    End If

    '    If dlgBrowseFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
    '        txtLocalScriptDir.Text = dlgBrowseFolder.SelectedPath
    '    End If

    'End Sub

    'Private Sub cmdBrowseFieldLocalModel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseFieldLocalModel.Click

    '    If dlgBrowseFolder.RootFolder = Nothing Then
    '        dlgBrowseFolder.SelectedPath = "" 'LocalModelFolderPath
    '    End If

    '    If dlgBrowseFolder.ShowDialog() = Windows.Forms.DialogResult.OK Then
    '        txtLocalModelDir.Text = dlgBrowseFolder.SelectedPath
    '    End If

    'End Sub

    Private Sub cmdGetDTD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetDTD.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalDTDDir.Text, ".dtd", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDTDDir = destination
            txtLocalDTDDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetDDL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetDDL.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalDDLDir.Text, ".ddl", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDDLDir = destination
            txtLocalDDLDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetCobol_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetCobol.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalCobolDir.Text, ".cob", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalCobolDir = destination
            txtLocalCobolDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetC.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalCDir.Text, ".h", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalCDir = destination
            txtLocalCDir.Text = destination
        End If

    End Sub

    Private Sub cmdGetDML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetDML.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalDMLDir.Text, ".dml", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalDMLDir = destination
            txtLocalDMLDir.Text = destination
        End If

    End Sub

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

    Private Sub cmdGetDBD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetDBD.Click

        Dim FTPClient As frmFTPClient = New frmFTPClient
        Dim destination As String

        destination = FTPClient.BrowseFile(Me.txtLocalDBDDir.Text, ".dbd", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

        If destination <> "" Then
            destination = GetCaseSensetivePath(destination)
            objThis.LocalScriptDir = destination
            txtLocalScriptDir.Text = destination
        End If

    End Sub

    'Private Sub btnRelDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelDir.Click

    'End Sub

    'Private Sub cmdPutRelDir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim FTPClient As frmFTPClient = New frmFTPClient
    '    Dim destination As String

    '    destination = FTPClient.BrowseFile(Me.txtRelDir.Text, ".*", modDeclares.enumCalledFrom.BY_ENVIRONMENT)

    '    If destination <> "" Then
    '        destination = GetCaseSensetivePath(destination)
    '        objThis.LocalScriptDir = destination
    '        txtLocalScriptDir.Text = destination
    '    End If
    'End Sub

End Class
