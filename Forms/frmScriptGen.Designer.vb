<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScriptGen
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
        Me.gbPath = New System.Windows.Forms.GroupBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.btnFTP = New System.Windows.Forms.Button
        Me.btnExpFolder = New System.Windows.Forms.Button
        Me.txtFolderPath = New System.Windows.Forms.TextBox
        Me.gbStudioFiles = New System.Windows.Forms.GroupBox
        Me.btnOpenINL = New System.Windows.Forms.Button
        Me.btnOpenSQD = New System.Windows.Forms.Button
        Me.txtINL = New System.Windows.Forms.TextBox
        Me.txtSQD = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.gbParseFiles = New System.Windows.Forms.GroupBox
        Me.btnOpenRPT = New System.Windows.Forms.Button
        Me.txtRPT = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.gbResults = New System.Windows.Forms.GroupBox
        Me.lblSummary = New System.Windows.Forms.Label
        Me.lblResult = New System.Windows.Forms.Label
        Me.gbScriptOptions = New System.Windows.Forms.GroupBox
        Me.btnParseOnlySQD = New System.Windows.Forms.Button
        Me.btnParseSQD = New System.Windows.Forms.Button
        Me.btnParseOnly = New System.Windows.Forms.Button
        Me.gbTarget = New System.Windows.Forms.GroupBox
        Me.rbAlltgt = New System.Windows.Forms.RadioButton
        Me.rbFieldtgt = New System.Windows.Forms.RadioButton
        Me.rbSelecttgt = New System.Windows.Forms.RadioButton
        Me.gbSource = New System.Windows.Forms.GroupBox
        Me.rbAllsrc = New System.Windows.Forms.RadioButton
        Me.rbFieldsrc = New System.Windows.Forms.RadioButton
        Me.rbSelectsrc = New System.Windows.Forms.RadioButton
        Me.Label9 = New System.Windows.Forms.Label
        Me.cbUID = New System.Windows.Forms.CheckBox
        Me.btnParse = New System.Windows.Forms.Button
        Me.cbDebugSrcData = New System.Windows.Forms.CheckBox
        Me.btnGenScript = New System.Windows.Forms.Button
        Me.btnSQData = New System.Windows.Forms.Button
        Me.Panel1.SuspendLayout()
        Me.gbPath.SuspendLayout()
        Me.gbStudioFiles.SuspendLayout()
        Me.gbParseFiles.SuspendLayout()
        Me.gbResults.SuspendLayout()
        Me.gbScriptOptions.SuspendLayout()
        Me.gbTarget.SuspendLayout()
        Me.gbSource.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(640, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 606)
        Me.GroupBox1.Size = New System.Drawing.Size(642, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(358, 629)
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(454, 629)
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(550, 629)
        '
        'Label1
        '
        Me.Label1.Text = "Script Generation Results"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(564, 39)
        Me.Label2.Text = "Choose to View, Save or Transfer each different type of generated script file."
        '
        'gbPath
        '
        Me.gbPath.Controls.Add(Me.Label8)
        Me.gbPath.Controls.Add(Me.btnFTP)
        Me.gbPath.Controls.Add(Me.btnExpFolder)
        Me.gbPath.Controls.Add(Me.txtFolderPath)
        Me.gbPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbPath.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbPath.Location = New System.Drawing.Point(12, 376)
        Me.gbPath.Name = "gbPath"
        Me.gbPath.Size = New System.Drawing.Size(616, 76)
        Me.gbPath.TabIndex = 59
        Me.gbPath.TabStop = False
        Me.gbPath.Text = "Scripts Directory"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(9, 20)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(202, 50)
        Me.Label8.TabIndex = 4
        Me.Label8.Text = "This is the Directory where your Scripts are written"
        '
        'btnFTP
        '
        Me.btnFTP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnFTP.Location = New System.Drawing.Point(507, 47)
        Me.btnFTP.Name = "btnFTP"
        Me.btnFTP.Size = New System.Drawing.Size(103, 23)
        Me.btnFTP.TabIndex = 3
        Me.btnFTP.Text = "FTP Scripts"
        Me.btnFTP.UseVisualStyleBackColor = True
        '
        'btnExpFolder
        '
        Me.btnExpFolder.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExpFolder.Location = New System.Drawing.Point(398, 47)
        Me.btnExpFolder.Name = "btnExpFolder"
        Me.btnExpFolder.Size = New System.Drawing.Size(103, 23)
        Me.btnExpFolder.TabIndex = 2
        Me.btnExpFolder.Text = "Explore Folder"
        Me.btnExpFolder.UseVisualStyleBackColor = True
        '
        'txtFolderPath
        '
        Me.txtFolderPath.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFolderPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtFolderPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFolderPath.Location = New System.Drawing.Point(217, 21)
        Me.txtFolderPath.Name = "txtFolderPath"
        Me.txtFolderPath.ReadOnly = True
        Me.txtFolderPath.Size = New System.Drawing.Size(393, 20)
        Me.txtFolderPath.TabIndex = 0
        '
        'gbStudioFiles
        '
        Me.gbStudioFiles.Controls.Add(Me.btnOpenINL)
        Me.gbStudioFiles.Controls.Add(Me.btnOpenSQD)
        Me.gbStudioFiles.Controls.Add(Me.txtINL)
        Me.gbStudioFiles.Controls.Add(Me.txtSQD)
        Me.gbStudioFiles.Controls.Add(Me.Label5)
        Me.gbStudioFiles.Controls.Add(Me.Label3)
        Me.gbStudioFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbStudioFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbStudioFiles.Location = New System.Drawing.Point(12, 458)
        Me.gbStudioFiles.Name = "gbStudioFiles"
        Me.gbStudioFiles.Size = New System.Drawing.Size(616, 82)
        Me.gbStudioFiles.TabIndex = 60
        Me.gbStudioFiles.TabStop = False
        Me.gbStudioFiles.Text = "Studio Generated Files"
        '
        'btnOpenINL
        '
        Me.btnOpenINL.Location = New System.Drawing.Point(507, 47)
        Me.btnOpenINL.Name = "btnOpenINL"
        Me.btnOpenINL.Size = New System.Drawing.Size(103, 23)
        Me.btnOpenINL.TabIndex = 8
        Me.btnOpenINL.Text = "Open INL"
        Me.btnOpenINL.UseVisualStyleBackColor = True
        '
        'btnOpenSQD
        '
        Me.btnOpenSQD.Location = New System.Drawing.Point(507, 17)
        Me.btnOpenSQD.Name = "btnOpenSQD"
        Me.btnOpenSQD.Size = New System.Drawing.Size(103, 23)
        Me.btnOpenSQD.TabIndex = 6
        Me.btnOpenSQD.Text = "Open SQD"
        Me.btnOpenSQD.UseVisualStyleBackColor = True
        '
        'txtINL
        '
        Me.txtINL.BackColor = System.Drawing.SystemColors.Window
        Me.txtINL.Location = New System.Drawing.Point(217, 49)
        Me.txtINL.Name = "txtINL"
        Me.txtINL.ReadOnly = True
        Me.txtINL.Size = New System.Drawing.Size(284, 20)
        Me.txtINL.TabIndex = 5
        '
        'txtSQD
        '
        Me.txtSQD.BackColor = System.Drawing.SystemColors.Window
        Me.txtSQD.Location = New System.Drawing.Point(217, 19)
        Me.txtSQD.Name = "txtSQD"
        Me.txtSQD.ReadOnly = True
        Me.txtSQD.Size = New System.Drawing.Size(284, 20)
        Me.txtSQD.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 52)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(128, 13)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Inline Script File (.inl)"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 13)
        Me.Label3.TabIndex = 0
        Me.Label3.Text = "Script File (.sqd)"
        '
        'gbParseFiles
        '
        Me.gbParseFiles.Controls.Add(Me.btnOpenRPT)
        Me.gbParseFiles.Controls.Add(Me.txtRPT)
        Me.gbParseFiles.Controls.Add(Me.Label6)
        Me.gbParseFiles.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbParseFiles.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbParseFiles.Location = New System.Drawing.Point(12, 546)
        Me.gbParseFiles.Name = "gbParseFiles"
        Me.gbParseFiles.Size = New System.Drawing.Size(616, 54)
        Me.gbParseFiles.TabIndex = 61
        Me.gbParseFiles.TabStop = False
        Me.gbParseFiles.Text = "Parser Generated File"
        '
        'btnOpenRPT
        '
        Me.btnOpenRPT.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnOpenRPT.Location = New System.Drawing.Point(507, 18)
        Me.btnOpenRPT.Name = "btnOpenRPT"
        Me.btnOpenRPT.Size = New System.Drawing.Size(103, 23)
        Me.btnOpenRPT.TabIndex = 5
        Me.btnOpenRPT.Text = "Open RPT"
        Me.btnOpenRPT.UseVisualStyleBackColor = True
        '
        'txtRPT
        '
        Me.txtRPT.BackColor = System.Drawing.SystemColors.Window
        Me.txtRPT.Location = New System.Drawing.Point(217, 20)
        Me.txtRPT.Name = "txtRPT"
        Me.txtRPT.ReadOnly = True
        Me.txtRPT.Size = New System.Drawing.Size(284, 20)
        Me.txtRPT.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(6, 23)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(140, 13)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Parser Report File (.rpt)"
        '
        'gbResults
        '
        Me.gbResults.Controls.Add(Me.lblSummary)
        Me.gbResults.Controls.Add(Me.lblResult)
        Me.gbResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbResults.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbResults.Location = New System.Drawing.Point(12, 248)
        Me.gbResults.Name = "gbResults"
        Me.gbResults.Size = New System.Drawing.Size(616, 122)
        Me.gbResults.TabIndex = 62
        Me.gbResults.TabStop = False
        Me.gbResults.Text = "Script Generation Results"
        '
        'lblSummary
        '
        Me.lblSummary.AutoSize = True
        Me.lblSummary.Location = New System.Drawing.Point(6, 38)
        Me.lblSummary.Name = "lblSummary"
        Me.lblSummary.Size = New System.Drawing.Size(57, 13)
        Me.lblSummary.TabIndex = 1
        Me.lblSummary.Text = "Summary"
        '
        'lblResult
        '
        Me.lblResult.AutoSize = True
        Me.lblResult.Location = New System.Drawing.Point(6, 16)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(43, 13)
        Me.lblResult.TabIndex = 0
        Me.lblResult.Text = "Result"
        '
        'gbScriptOptions
        '
        Me.gbScriptOptions.Controls.Add(Me.btnParseOnlySQD)
        Me.gbScriptOptions.Controls.Add(Me.btnParseSQD)
        Me.gbScriptOptions.Controls.Add(Me.btnParseOnly)
        Me.gbScriptOptions.Controls.Add(Me.gbTarget)
        Me.gbScriptOptions.Controls.Add(Me.gbSource)
        Me.gbScriptOptions.Controls.Add(Me.Label9)
        Me.gbScriptOptions.Controls.Add(Me.cbUID)
        Me.gbScriptOptions.Controls.Add(Me.btnParse)
        Me.gbScriptOptions.Controls.Add(Me.cbDebugSrcData)
        Me.gbScriptOptions.Controls.Add(Me.btnGenScript)
        Me.gbScriptOptions.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbScriptOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbScriptOptions.Location = New System.Drawing.Point(12, 74)
        Me.gbScriptOptions.Name = "gbScriptOptions"
        Me.gbScriptOptions.Size = New System.Drawing.Size(616, 168)
        Me.gbScriptOptions.TabIndex = 63
        Me.gbScriptOptions.TabStop = False
        Me.gbScriptOptions.Text = "Script Generation Options"
        '
        'btnParseOnlySQD
        '
        Me.btnParseOnlySQD.Location = New System.Drawing.Point(498, 44)
        Me.btnParseOnlySQD.Name = "btnParseOnlySQD"
        Me.btnParseOnlySQD.Size = New System.Drawing.Size(112, 23)
        Me.btnParseOnlySQD.TabIndex = 13
        Me.btnParseOnlySQD.Text = "Parse Only SQD"
        Me.btnParseOnlySQD.UseVisualStyleBackColor = True
        '
        'btnParseSQD
        '
        Me.btnParseSQD.Location = New System.Drawing.Point(382, 44)
        Me.btnParseSQD.Name = "btnParseSQD"
        Me.btnParseSQD.Size = New System.Drawing.Size(110, 23)
        Me.btnParseSQD.TabIndex = 12
        Me.btnParseSQD.Text = "Gen/Parse SQD"
        Me.btnParseSQD.UseVisualStyleBackColor = True
        '
        'btnParseOnly
        '
        Me.btnParseOnly.Location = New System.Drawing.Point(498, 14)
        Me.btnParseOnly.Name = "btnParseOnly"
        Me.btnParseOnly.Size = New System.Drawing.Size(112, 24)
        Me.btnParseOnly.TabIndex = 11
        Me.btnParseOnly.Text = "Parse Only INL"
        Me.btnParseOnly.UseVisualStyleBackColor = True
        '
        'gbTarget
        '
        Me.gbTarget.Controls.Add(Me.rbAlltgt)
        Me.gbTarget.Controls.Add(Me.rbFieldtgt)
        Me.gbTarget.Controls.Add(Me.rbSelecttgt)
        Me.gbTarget.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbTarget.Location = New System.Drawing.Point(305, 73)
        Me.gbTarget.Name = "gbTarget"
        Me.gbTarget.Size = New System.Drawing.Size(305, 89)
        Me.gbTarget.TabIndex = 10
        Me.gbTarget.TabStop = False
        Me.gbTarget.Text = "Target Mapping Level"
        '
        'rbAlltgt
        '
        Me.rbAlltgt.AutoSize = True
        Me.rbAlltgt.Location = New System.Drawing.Point(6, 19)
        Me.rbAlltgt.Name = "rbAlltgt"
        Me.rbAlltgt.Size = New System.Drawing.Size(276, 17)
        Me.rbAlltgt.TabIndex = 6
        Me.rbAlltgt.TabStop = True
        Me.rbAlltgt.Text = "Show ""Datastore.Description.Field"" in Procs"
        Me.rbAlltgt.UseVisualStyleBackColor = True
        '
        'rbFieldtgt
        '
        Me.rbFieldtgt.AutoSize = True
        Me.rbFieldtgt.Location = New System.Drawing.Point(6, 65)
        Me.rbFieldtgt.Name = "rbFieldtgt"
        Me.rbFieldtgt.Size = New System.Drawing.Size(202, 17)
        Me.rbFieldtgt.TabIndex = 8
        Me.rbFieldtgt.TabStop = True
        Me.rbFieldtgt.Text = "Show Field Name Only in Procs"
        Me.rbFieldtgt.UseVisualStyleBackColor = True
        '
        'rbSelecttgt
        '
        Me.rbSelecttgt.AutoSize = True
        Me.rbSelecttgt.Location = New System.Drawing.Point(6, 42)
        Me.rbSelecttgt.Name = "rbSelecttgt"
        Me.rbSelecttgt.Size = New System.Drawing.Size(210, 17)
        Me.rbSelecttgt.TabIndex = 7
        Me.rbSelecttgt.TabStop = True
        Me.rbSelecttgt.Text = "Show ""Desription.Field"" in Procs"
        Me.rbSelecttgt.UseVisualStyleBackColor = True
        '
        'gbSource
        '
        Me.gbSource.Controls.Add(Me.rbAllsrc)
        Me.gbSource.Controls.Add(Me.rbFieldsrc)
        Me.gbSource.Controls.Add(Me.rbSelectsrc)
        Me.gbSource.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbSource.Location = New System.Drawing.Point(12, 73)
        Me.gbSource.Name = "gbSource"
        Me.gbSource.Size = New System.Drawing.Size(287, 89)
        Me.gbSource.TabIndex = 9
        Me.gbSource.TabStop = False
        Me.gbSource.Text = "Source Mapping Level"
        '
        'rbAllsrc
        '
        Me.rbAllsrc.AutoSize = True
        Me.rbAllsrc.Location = New System.Drawing.Point(6, 19)
        Me.rbAllsrc.Name = "rbAllsrc"
        Me.rbAllsrc.Size = New System.Drawing.Size(276, 17)
        Me.rbAllsrc.TabIndex = 6
        Me.rbAllsrc.TabStop = True
        Me.rbAllsrc.Text = "Show ""Datastore.Description.Field"" in Procs"
        Me.rbAllsrc.UseVisualStyleBackColor = True
        '
        'rbFieldsrc
        '
        Me.rbFieldsrc.AutoSize = True
        Me.rbFieldsrc.Location = New System.Drawing.Point(6, 65)
        Me.rbFieldsrc.Name = "rbFieldsrc"
        Me.rbFieldsrc.Size = New System.Drawing.Size(202, 17)
        Me.rbFieldsrc.TabIndex = 8
        Me.rbFieldsrc.TabStop = True
        Me.rbFieldsrc.Text = "Show Field Name Only in Procs"
        Me.rbFieldsrc.UseVisualStyleBackColor = True
        '
        'rbSelectsrc
        '
        Me.rbSelectsrc.AutoSize = True
        Me.rbSelectsrc.Location = New System.Drawing.Point(6, 42)
        Me.rbSelectsrc.Name = "rbSelectsrc"
        Me.rbSelectsrc.Size = New System.Drawing.Size(210, 17)
        Me.rbSelectsrc.TabIndex = 7
        Me.rbSelectsrc.TabStop = True
        Me.rbSelectsrc.Text = "Show ""Desription.Field"" in Procs"
        Me.rbSelectsrc.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Arial Unicode MS", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label9.Location = New System.Drawing.Point(9, 54)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(269, 15)
        Me.Label9.TabIndex = 5
        Me.Label9.Text = "UID and PWD are needed to properly Parse the Script"
        '
        'cbUID
        '
        Me.cbUID.AutoSize = True
        Me.cbUID.Location = New System.Drawing.Point(12, 39)
        Me.cbUID.Name = "cbUID"
        Me.cbUID.Size = New System.Drawing.Size(243, 17)
        Me.cbUID.TabIndex = 4
        Me.cbUID.Text = "Use Username and Password In Script"
        Me.cbUID.UseVisualStyleBackColor = True
        '
        'btnParse
        '
        Me.btnParse.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnParse.Location = New System.Drawing.Point(382, 13)
        Me.btnParse.Name = "btnParse"
        Me.btnParse.Size = New System.Drawing.Size(110, 24)
        Me.btnParse.TabIndex = 3
        Me.btnParse.Text = "Gen/Parse INL"
        Me.btnParse.UseVisualStyleBackColor = True
        '
        'cbDebugSrcData
        '
        Me.cbDebugSrcData.AutoSize = True
        Me.cbDebugSrcData.Location = New System.Drawing.Point(12, 20)
        Me.cbDebugSrcData.Name = "cbDebugSrcData"
        Me.cbDebugSrcData.Size = New System.Drawing.Size(295, 17)
        Me.cbDebugSrcData.TabIndex = 1
        Me.cbDebugSrcData.Text = "Use Output Messages for Source in Procedures"
        Me.cbDebugSrcData.UseVisualStyleBackColor = True
        '
        'btnGenScript
        '
        Me.btnGenScript.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnGenScript.Location = New System.Drawing.Point(310, 13)
        Me.btnGenScript.Name = "btnGenScript"
        Me.btnGenScript.Size = New System.Drawing.Size(68, 54)
        Me.btnGenScript.TabIndex = 0
        Me.btnGenScript.Text = "Generate Script"
        Me.btnGenScript.UseVisualStyleBackColor = True
        '
        'btnSQData
        '
        Me.btnSQData.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSQData.Location = New System.Drawing.Point(175, 630)
        Me.btnSQData.Name = "btnSQData"
        Me.btnSQData.Size = New System.Drawing.Size(165, 23)
        Me.btnSQData.TabIndex = 64
        Me.btnSQData.Text = "Run SQData"
        Me.btnSQData.UseVisualStyleBackColor = True
        '
        'frmScriptGen
        '
        Me.ClientSize = New System.Drawing.Size(640, 668)
        Me.Controls.Add(Me.gbScriptOptions)
        Me.Controls.Add(Me.gbStudioFiles)
        Me.Controls.Add(Me.gbPath)
        Me.Controls.Add(Me.gbResults)
        Me.Controls.Add(Me.gbParseFiles)
        Me.Controls.Add(Me.btnSQData)
        Me.Name = "frmScriptGen"
        Me.Text = "SQData Studio V3"
        Me.Controls.SetChildIndex(Me.btnSQData, 0)
        Me.Controls.SetChildIndex(Me.gbParseFiles, 0)
        Me.Controls.SetChildIndex(Me.gbResults, 0)
        Me.Controls.SetChildIndex(Me.gbPath, 0)
        Me.Controls.SetChildIndex(Me.gbStudioFiles, 0)
        Me.Controls.SetChildIndex(Me.gbScriptOptions, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Panel1.ResumeLayout(False)
        Me.gbPath.ResumeLayout(False)
        Me.gbPath.PerformLayout()
        Me.gbStudioFiles.ResumeLayout(False)
        Me.gbStudioFiles.PerformLayout()
        Me.gbParseFiles.ResumeLayout(False)
        Me.gbParseFiles.PerformLayout()
        Me.gbResults.ResumeLayout(False)
        Me.gbResults.PerformLayout()
        Me.gbScriptOptions.ResumeLayout(False)
        Me.gbScriptOptions.PerformLayout()
        Me.gbTarget.ResumeLayout(False)
        Me.gbTarget.PerformLayout()
        Me.gbSource.ResumeLayout(False)
        Me.gbSource.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gbPath As System.Windows.Forms.GroupBox
    Friend WithEvents txtFolderPath As System.Windows.Forms.TextBox
    Friend WithEvents btnExpFolder As System.Windows.Forms.Button
    Friend WithEvents gbStudioFiles As System.Windows.Forms.GroupBox
    Friend WithEvents gbParseFiles As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnOpenINL As System.Windows.Forms.Button
    Friend WithEvents btnOpenSQD As System.Windows.Forms.Button
    Friend WithEvents txtINL As System.Windows.Forms.TextBox
    Friend WithEvents txtSQD As System.Windows.Forms.TextBox
    Friend WithEvents txtRPT As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents btnOpenRPT As System.Windows.Forms.Button
    Friend WithEvents btnFTP As System.Windows.Forms.Button
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents gbResults As System.Windows.Forms.GroupBox
    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents lblSummary As System.Windows.Forms.Label
    Friend WithEvents gbScriptOptions As System.Windows.Forms.GroupBox
    Friend WithEvents cbDebugSrcData As System.Windows.Forms.CheckBox
    Friend WithEvents btnGenScript As System.Windows.Forms.Button
    Friend WithEvents cbUID As System.Windows.Forms.CheckBox
    Friend WithEvents btnParse As System.Windows.Forms.Button
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents rbFieldsrc As System.Windows.Forms.RadioButton
    Friend WithEvents rbSelectsrc As System.Windows.Forms.RadioButton
    Friend WithEvents rbAllsrc As System.Windows.Forms.RadioButton
    Friend WithEvents gbSource As System.Windows.Forms.GroupBox
    Friend WithEvents gbTarget As System.Windows.Forms.GroupBox
    Friend WithEvents rbAlltgt As System.Windows.Forms.RadioButton
    Friend WithEvents rbFieldtgt As System.Windows.Forms.RadioButton
    Friend WithEvents rbSelecttgt As System.Windows.Forms.RadioButton
    Friend WithEvents btnParseOnly As System.Windows.Forms.Button
    Friend WithEvents btnParseOnlySQD As System.Windows.Forms.Button
    Friend WithEvents btnParseSQD As System.Windows.Forms.Button
    Friend WithEvents btnSQData As System.Windows.Forms.Button

End Class
