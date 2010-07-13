Public Class ctlVariable
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)
   
    Dim objThis As New clsVariable
    Friend WithEvents gbVar As System.Windows.Forms.GroupBox
    Friend WithEvents gbDesc As System.Windows.Forms.GroupBox
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
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtVariableDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtVariableName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtVariableSize As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtInitVal As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmdHelp As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtVariableDesc = New System.Windows.Forms.TextBox
        Me.txtVariableName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtVariableSize = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtInitVal = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.cmdHelp = New System.Windows.Forms.Button
        Me.gbVar = New System.Windows.Forms.GroupBox
        Me.gbDesc = New System.Windows.Forms.GroupBox
        Me.gbVar.SuspendLayout()
        Me.gbDesc.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdSave.Location = New System.Drawing.Point(333, 410)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(72, 24)
        Me.cmdSave.TabIndex = 5
        Me.cmdSave.Text = "&Save"
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(413, 410)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 394)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(557, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'txtVariableDesc
        '
        Me.txtVariableDesc.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtVariableDesc.Location = New System.Drawing.Point(3, 16)
        Me.txtVariableDesc.MaxLength = 1000
        Me.txtVariableDesc.Multiline = True
        Me.txtVariableDesc.Name = "txtVariableDesc"
        Me.txtVariableDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtVariableDesc.Size = New System.Drawing.Size(358, 66)
        Me.txtVariableDesc.TabIndex = 4
        '
        'txtVariableName
        '
        Me.txtVariableName.Location = New System.Drawing.Point(82, 19)
        Me.txtVariableName.MaxLength = 128
        Me.txtVariableName.Name = "txtVariableName"
        Me.txtVariableName.Size = New System.Drawing.Size(286, 20)
        Me.txtVariableName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 14)
        Me.Label3.TabIndex = 33
        Me.Label3.Text = "Name"
        '
        'txtVariableSize
        '
        Me.txtVariableSize.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVariableSize.Location = New System.Drawing.Point(82, 45)
        Me.txtVariableSize.MaxLength = 128
        Me.txtVariableSize.Name = "txtVariableSize"
        Me.txtVariableSize.Size = New System.Drawing.Size(286, 20)
        Me.txtVariableSize.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 48)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(30, 14)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Size"
        '
        'txtInitVal
        '
        Me.txtInitVal.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInitVal.Location = New System.Drawing.Point(82, 71)
        Me.txtInitVal.MaxLength = 128
        Me.txtInitVal.Name = "txtInitVal"
        Me.txtInitVal.Size = New System.Drawing.Size(286, 20)
        Me.txtInitVal.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(6, 74)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(70, 14)
        Me.Label6.TabIndex = 35
        Me.Label6.Text = "Initial Value"
        '
        'cmdHelp
        '
        Me.cmdHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHelp.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdHelp.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdHelp.Location = New System.Drawing.Point(493, 410)
        Me.cmdHelp.Name = "cmdHelp"
        Me.cmdHelp.Size = New System.Drawing.Size(72, 24)
        Me.cmdHelp.TabIndex = 7
        Me.cmdHelp.Text = "&Help"
        '
        'gbVar
        '
        Me.gbVar.Controls.Add(Me.gbDesc)
        Me.gbVar.Controls.Add(Me.txtVariableName)
        Me.gbVar.Controls.Add(Me.Label3)
        Me.gbVar.Controls.Add(Me.Label5)
        Me.gbVar.Controls.Add(Me.txtVariableSize)
        Me.gbVar.Controls.Add(Me.Label6)
        Me.gbVar.Controls.Add(Me.txtInitVal)
        Me.gbVar.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.gbVar.ForeColor = System.Drawing.Color.White
        Me.gbVar.Location = New System.Drawing.Point(3, 0)
        Me.gbVar.Name = "gbVar"
        Me.gbVar.Size = New System.Drawing.Size(376, 189)
        Me.gbVar.TabIndex = 40
        Me.gbVar.TabStop = False
        Me.gbVar.Text = "Variable Properties"
        '
        'gbDesc
        '
        Me.gbDesc.Controls.Add(Me.txtVariableDesc)
        Me.gbDesc.ForeColor = System.Drawing.Color.White
        Me.gbDesc.Location = New System.Drawing.Point(6, 97)
        Me.gbDesc.Name = "gbDesc"
        Me.gbDesc.Size = New System.Drawing.Size(364, 85)
        Me.gbDesc.TabIndex = 36
        Me.gbDesc.TabStop = False
        Me.gbDesc.Text = "Description"
        '
        'ctlVariable
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.gbVar)
        Me.Controls.Add(Me.cmdHelp)
        Me.Controls.Add(Me.cmdSave)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlVariable"
        Me.Size = New System.Drawing.Size(573, 442)
        Me.gbVar.ResumeLayout(False)
        Me.gbVar.PerformLayout()
        Me.gbDesc.ResumeLayout(False)
        Me.gbDesc.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtVariableName.Enabled = True '//added on 7/24 to disable object name editing
        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsVariable
        ClearControls(Me.Controls)

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

    Public Function Save() As Boolean
        Try
            Dim OldName As String = ""

            '// First Check Validity before Saving
            If ValidateNewName(txtVariableName.Text) = False Then
                Exit Function
            End If

            If objThis.VariableName <> txtVariableName.Text Then
                objThis.IsRenamed = RenameVariable(objThis, txtVariableName.Text)
            End If

            If objThis.IsRenamed = False Then
                txtVariableName.Text = objThis.VariableName
            Else
                OldName = objThis.VariableName
                objThis.VariableName = txtVariableName.Text
            End If

            objThis.VarSize = txtVariableSize.Text
            objThis.VarInitVal = txtInitVal.Text
            objThis.VariableDescription = txtVariableDesc.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then Exit Function
            Else
                If objThis.Save = False Then
                    Exit Function
                Else
                    If objThis.Engine Is Nothing Then
                        For Each sys As clsSystem In objThis.Environment.Systems
                            For Each eng As clsEngine In sys.Engines
                                For Each var As clsVariable In eng.Variables
                                    If OldName = "" Then
                                        OldName = objThis.VariableName
                                    End If
                                    If var.VariableName = OldName Then
                                        var.DeleteATTR()
                                        var.VariableName = objThis.VariableName
                                        If UpdateChildVars(var) = True Then
                                            If var.Save() = False Then
                                                MsgBox("Variable: " & var.Text & " failed to update correctly in Engine: " & _
                                                eng.Text, MsgBoxStyle.Information, "Problem saving variable")
                                            Else
                                                var.DeleteATTR()
                                                var.InsertATTR()
                                            End If
                                        Else
                                            MsgBox("Variable: " & var.Text & " failed to update correctly in Engine: " & _
                                            eng.Text, MsgBoxStyle.Information, "Problem saving variable")
                                        End If
                                    End If
                                Next
                            Next
                        Next
                    End If
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
            LogError(ex)
        End Try

    End Function

    Function UpdateChildVars(ByRef var As clsVariable) As Boolean

        Try
            var.VariableDescription = objThis.VariableDescription
            var.VariableType = objThis.VariableType
            var.VarInitVal = objThis.VarInitVal
            var.VarSize = objThis.VarSize

            Return True

        Catch ex As Exception
            LogError(ex, "ctlvariable UpdateChildVars")
            Return False
        End Try

    End Function

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

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInitVal.TextChanged, txtVariableDesc.TextChanged, txtVariableName.TextChanged, txtVariableSize.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Public Function EditObj(ByVal obj As INode) As clsVariable

        IsNewObj = False

        StartLoad()

        objThis = obj '//Load the form env object
        objThis.LoadMe()

        UpdateFields()

        EditObj = objThis

        EndLoad()

    End Function

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtVariableName.Text = objThis.VariableName
        txtVariableSize.Text = objThis.VarSize
        txtInitVal.Text = objThis.VarInitVal
        txtVariableDesc.Text = objThis.VariableDescription

    End Function

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click
        ShowHelp(modHelp.HHId.H_Variables)
    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles GroupBox1.KeyDown, txtInitVal.KeyDown, txtVariableDesc.KeyDown, txtVariableName.KeyDown, txtVariableSize.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

End Class
