Public Class frmVariable
    Inherits SQDStudio.frmBlank

    Public objThis As New clsVariable

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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtVariableName As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtVariableSize As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtInitVal As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtVariableDesc As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVariable))
        Me.txtVariableDesc = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtVariableName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtVariableSize = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtInitVal = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        'Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        'Me.PictureBox1.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstu" & _
        '    "dio\images\FormTop\sq_skyblue.jpg"
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.TabIndex = 6
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Text = "Variable definition"
        '
        'Label2
        '
        Me.Label2.Text = "Enter a unique variable name within an engine, and its attributes."
        '
        'txtVariableDesc
        '
        Me.txtVariableDesc.Location = New System.Drawing.Point(128, 208)
        Me.txtVariableDesc.MaxLength = 1000
        Me.txtVariableDesc.Multiline = True
        Me.txtVariableDesc.Name = "txtVariableDesc"
        Me.txtVariableDesc.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtVariableDesc.Size = New System.Drawing.Size(351, 56)
        Me.txtVariableDesc.TabIndex = 4
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(16, 208)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(112, 20)
        Me.Label4.TabIndex = 14
        Me.Label4.Text = "Description"
        '
        'txtVariableName
        '
        Me.txtVariableName.Location = New System.Drawing.Point(128, 88)
        Me.txtVariableName.MaxLength = 128
        Me.txtVariableName.Name = "txtVariableName"
        Me.txtVariableName.Size = New System.Drawing.Size(351, 20)
        Me.txtVariableName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(16, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 20)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "Variable  Name"
        '
        'txtVariableSize
        '
        Me.txtVariableSize.Location = New System.Drawing.Point(128, 128)
        Me.txtVariableSize.MaxLength = 128
        Me.txtVariableSize.Name = "txtVariableSize"
        Me.txtVariableSize.Size = New System.Drawing.Size(351, 20)
        Me.txtVariableSize.TabIndex = 2
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(16, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(112, 20)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Size"
        '
        'txtInitVal
        '
        Me.txtInitVal.Location = New System.Drawing.Point(128, 168)
        Me.txtInitVal.MaxLength = 128
        Me.txtInitVal.Name = "txtInitVal"
        Me.txtInitVal.Size = New System.Drawing.Size(351, 20)
        Me.txtInitVal.TabIndex = 3
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(16, 168)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(112, 20)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Initial Value"
        '
        'frmVariable
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(514, 335)
        Me.Controls.Add(Me.txtVariableDesc)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtVariableName)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtVariableSize)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtInitVal)
        Me.Controls.Add(Me.Label6)
        Me.Name = "frmVariable"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Variable Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtInitVal, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtVariableSize, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtVariableName, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtVariableDesc, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInitVal.TextChanged, txtVariableDesc.TextChanged, txtVariableName.TextChanged, txtVariableSize.TextChanged

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
            Dim EngObj As clsEngine = Nothing
            Dim EnvObj As clsEnvironment = Nothing
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            EnvObj = objThis.Environment
            If objThis.Engine IsNot Nothing Then
                EngObj = objThis.Engine
            End If

            While GoodName = False
                GoodName = True
                NewName = "V_Var" & Count.ToString '& "_" & Strings.Left(EngObj.Text, 9)

                If objThis.Engine IsNot Nothing Then
                    For Each TestVar As clsVariable In EngObj.Variables
                        If TestVar.VariableName = NewName Then
                            GoodName = False
                            Exit For
                        End If
                    Next
                End If

                For Each TestVar As clsVariable In EnvObj.Variables
                    If TestVar.VariableName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next

                Count += 1
            End While

            txtVariableName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmVar SetDefaultName")
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

    '//Returns new Variable object
    Public Function NewObj() As clsVariable

        IsNewObj = True

        StartLoad()

        SetDefaultName()

        txtVariableName.SelectAll()

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
            If objThis.ValidateNewObject(txtVariableName.Text) = False Or ValidateNewName(txtVariableName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.VariableName = txtVariableName.Text
            objThis.VarInitVal = txtInitVal.Text
            objThis.VarSize = IIf(txtVariableSize.Text = "", 0, Val(txtVariableSize.Text))
            objThis.VariableDescription = txtVariableDesc.Text

            If IsNewObj = True Then
                'If objThis.Engine IsNot Nothing Then
                '    Dim objenv As clsVariable = CType(objThis.Clone(objThis.Environment), clsVariable)
                '    objenv.Engine = Nothing
                '    If objenv.AddNew = False Then
                '        DialogResult = Windows.Forms.DialogResult.Retry
                '        Exit Sub
                '    Else
                '        If objThis.AddNew = False Then
                '            DialogResult = Windows.Forms.DialogResult.Retry
                '            Exit Sub
                '        End If
                '    End If
                'Else
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
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Add_Variables)

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

End Class
