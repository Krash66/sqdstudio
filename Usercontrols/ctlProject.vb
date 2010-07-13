Public Class ctlProject
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)
    
    Dim objThis As New clsProject
    Friend WithEvents lblMetaVer As System.Windows.Forms.Label
    Friend WithEvents txtMetaVer As System.Windows.Forms.TextBox
   
    Dim IsNewObj As Boolean

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
    Friend WithEvents txtVersion As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtMetaDSN As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtProjectName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    Friend WithEvents gbProj As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCreateDate As System.Windows.Forms.TextBox
    Friend WithEvents gbV3 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLastUpd As System.Windows.Forms.TextBox
    Friend WithEvents txtCust As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtVersion = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtMetaDSN = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtProjectName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.gbProj = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCreateDate = New System.Windows.Forms.TextBox
        Me.gbV3 = New System.Windows.Forms.GroupBox
        Me.txtLastUpd = New System.Windows.Forms.TextBox
        Me.txtCust = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtMetaVer = New System.Windows.Forms.TextBox
        Me.lblMetaVer = New System.Windows.Forms.Label
        Me.gbProj.SuspendLayout()
        Me.gbV3.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtVersion
        '
        Me.txtVersion.BackColor = System.Drawing.SystemColors.Window
        Me.txtVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVersion.Location = New System.Drawing.Point(141, 123)
        Me.txtVersion.MaxLength = 128
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.ReadOnly = True
        Me.txtVersion.Size = New System.Drawing.Size(341, 20)
        Me.txtVersion.TabIndex = 3
        Me.txtVersion.Text = "1.0.0"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 126)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(93, 14)
        Me.Label5.TabIndex = 21
        Me.Label5.Text = "Project Version"
        '
        'txtMetaDSN
        '
        Me.txtMetaDSN.BackColor = System.Drawing.SystemColors.Window
        Me.txtMetaDSN.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMetaDSN.Location = New System.Drawing.Point(141, 71)
        Me.txtMetaDSN.MaxLength = 128
        Me.txtMetaDSN.Name = "txtMetaDSN"
        Me.txtMetaDSN.ReadOnly = True
        Me.txtMetaDSN.Size = New System.Drawing.Size(341, 20)
        Me.txtMetaDSN.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 74)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(81, 14)
        Me.Label6.TabIndex = 19
        Me.Label6.Text = "Metadata DSN"
        '
        'txtProjectName
        '
        Me.txtProjectName.Location = New System.Drawing.Point(141, 19)
        Me.txtProjectName.MaxLength = 128
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(341, 20)
        Me.txtProjectName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 14)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Project Name"
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(374, 434)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 4
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(454, 434)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 5
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 418)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(598, 7)
        Me.GroupBox1.TabIndex = 24
        Me.GroupBox1.TabStop = False
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(534, 434)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 6
        Me.cmdHelp.Text = "&Help"
        '
        'gbProj
        '
        Me.gbProj.Controls.Add(Me.lblMetaVer)
        Me.gbProj.Controls.Add(Me.txtMetaVer)
        Me.gbProj.Controls.Add(Me.Label1)
        Me.gbProj.Controls.Add(Me.txtCreateDate)
        Me.gbProj.Controls.Add(Me.txtProjectName)
        Me.gbProj.Controls.Add(Me.Label3)
        Me.gbProj.Controls.Add(Me.Label6)
        Me.gbProj.Controls.Add(Me.txtMetaDSN)
        Me.gbProj.Controls.Add(Me.txtVersion)
        Me.gbProj.Controls.Add(Me.Label5)
        Me.gbProj.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbProj.ForeColor = System.Drawing.Color.White
        Me.gbProj.Location = New System.Drawing.Point(3, 3)
        Me.gbProj.Name = "gbProj"
        Me.gbProj.Size = New System.Drawing.Size(491, 154)
        Me.gbProj.TabIndex = 25
        Me.gbProj.TabStop = False
        Me.gbProj.Text = "Project Properties"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(129, 13)
        Me.Label1.TabIndex = 23
        Me.Label1.Text = "Project Creation Date"
        '
        'txtCreateDate
        '
        Me.txtCreateDate.BackColor = System.Drawing.SystemColors.Window
        Me.txtCreateDate.Location = New System.Drawing.Point(141, 45)
        Me.txtCreateDate.Name = "txtCreateDate"
        Me.txtCreateDate.ReadOnly = True
        Me.txtCreateDate.Size = New System.Drawing.Size(341, 20)
        Me.txtCreateDate.TabIndex = 22
        '
        'gbV3
        '
        Me.gbV3.Controls.Add(Me.txtLastUpd)
        Me.gbV3.Controls.Add(Me.txtCust)
        Me.gbV3.Controls.Add(Me.Label4)
        Me.gbV3.Controls.Add(Me.Label2)
        Me.gbV3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbV3.ForeColor = System.Drawing.Color.White
        Me.gbV3.Location = New System.Drawing.Point(3, 163)
        Me.gbV3.Name = "gbV3"
        Me.gbV3.Size = New System.Drawing.Size(491, 71)
        Me.gbV3.TabIndex = 26
        Me.gbV3.TabStop = False
        '
        'txtLastUpd
        '
        Me.txtLastUpd.Location = New System.Drawing.Point(141, 39)
        Me.txtLastUpd.Name = "txtLastUpd"
        Me.txtLastUpd.Size = New System.Drawing.Size(341, 20)
        Me.txtLastUpd.TabIndex = 3
        '
        'txtCust
        '
        Me.txtCust.Location = New System.Drawing.Point(141, 13)
        Me.txtCust.Name = "txtCust"
        Me.txtCust.Size = New System.Drawing.Size(341, 20)
        Me.txtCust.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(6, 42)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(127, 13)
        Me.Label4.TabIndex = 1
        Me.Label4.Text = "Project Last Updated"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Customer Name"
        '
        'txtMetaVer
        '
        Me.txtMetaVer.Location = New System.Drawing.Point(141, 97)
        Me.txtMetaVer.Name = "txtMetaVer"
        Me.txtMetaVer.Size = New System.Drawing.Size(341, 20)
        Me.txtMetaVer.TabIndex = 24
        '
        'lblMetaVer
        '
        Me.lblMetaVer.AutoSize = True
        Me.lblMetaVer.Location = New System.Drawing.Point(6, 100)
        Me.lblMetaVer.Name = "lblMetaVer"
        Me.lblMetaVer.Size = New System.Drawing.Size(108, 13)
        Me.lblMetaVer.TabIndex = 25
        Me.lblMetaVer.Text = "MetaData Version"
        '
        'ctlProject
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Controls.Add(Me.gbV3)
        Me.Controls.Add(Me.gbProj)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlProject"
        Me.Size = New System.Drawing.Size(614, 466)
        Me.gbProj.ResumeLayout(False)
        Me.gbProj.PerformLayout()
        Me.gbV3.ResumeLayout(False)
        Me.gbV3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtProjectName.Enabled = True '//added on 7/24 to disable object name editing

        If objThis.ProjectMetaVersion = enumMetaVersion.V2 Then
            'txtMetaVer.Text = "Version 2 MetaData"
            gbV3.Visible = False
        Else
            'txtMetaVer.Text = "Version 3 MetaData"
            gbV3.Visible = True
        End If
        ClearControls(Me.Controls)

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False


    End Sub
    
    Private Sub OnChange(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMetaDSN.TextChanged, txtProjectName.TextChanged, txtVersion.TextChanged, txtCreateDate.TextChanged, txtLastUpd.TextChanged, txtCust.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

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

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Project)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtMetaDSN.KeyDown, txtProjectName.KeyDown, txtVersion.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select

    End Sub

#End Region

    Public Function EditObj(ByVal obj As INode) As clsProject
        
        IsNewObj = False

        objThis = obj
        objThis.LoadMe()

        StartLoad()

        UpdateFields()

        EndLoad()

        EditObj = Me.objThis

    End Function

    '//Set values from objThis to form controls
    Private Function UpdateFields() As Boolean

        txtProjectName.Text = objThis.ProjectName
        txtCreateDate.Text = objThis.ProjectCreationDate
        txtMetaDSN.Text = objThis.ProjectMetaDSN
        txtVersion.Text = objThis.ProjectVersion
        If objThis.ChangeVersion = True Then
            objThis.Save()
        End If
        If objThis.ProjectMetaVersion = enumMetaVersion.V3 Then
            txtCust.Text = objThis.ProjectCustomerName
            txtLastUpd.Text = objThis.ProjectLastUpdated
            txtMetaVer.Text = "Version 3 MetaData"
        Else
            txtMetaVer.Text = "Version 2 MetaData"
        End If

    End Function

    Public Function Save() As Boolean

        Try
            '// First Check Validity before Saving
            If ValidateNewName(txtProjectName.Text) = False Then
                Exit Function
            End If

            If objThis.ProjectName <> txtProjectName.Text Then
                objThis.IsRenamed = RenameProject(objThis, txtProjectName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtProjectName.Text = objThis.ProjectName
            Else
                objThis.ProjectName = txtProjectName.Text
            End If

            objThis.ProjectVersion = txtVersion.Text
            objThis.ProjectMetaDSN = txtMetaDSN.Text

            If objThis.ProjectMetaVersion = enumMetaVersion.V3 Then
                objThis.ProjectCustomerName = txtCust.Text
                '***this is handled when Project attr are saved
                'objThis.ProjectLastUpdated = Now.ToString
            End If


            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    MsgBox("Error during project save operation", MsgBoxStyle.Critical, MsgTitle)
                End If
            Else
                objThis.Save()
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
            LogError(ex)
        End Try

    End Function

End Class
