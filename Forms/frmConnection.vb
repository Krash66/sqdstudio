Public Class frmConnection
    Inherits SQDStudio.frmBlank

    Public objThis As New clsConnection
    Dim mDSNTable As System.Data.DataTable        '/// Table of ODBC sources
    Private ODBCHelper As New CRODBCHelper()   '// gets ODBC list

    Dim IsNewObj As Boolean

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
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtConnectionDesc As System.Windows.Forms.TextBox
    Friend WithEvents cmbCnnType As System.Windows.Forms.ComboBox
    Friend WithEvents txtDatabase As System.Windows.Forms.TextBox
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserId As System.Windows.Forms.TextBox
    Friend WithEvents txtConnectionName As System.Windows.Forms.TextBox
    Friend WithEvents txtDateformat As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents cmbODBC As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtConnectionDesc = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.cmbCnnType = New System.Windows.Forms.ComboBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtDatabase = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.txtPassword = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtUserId = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.txtConnectionName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtDateformat = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.cmbODBC = New System.Windows.Forms.ComboBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(579, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 282)
        Me.GroupBox1.Size = New System.Drawing.Size(581, 7)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(297, 305)
        Me.cmdOk.TabIndex = 8
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(393, 305)
        Me.cmdCancel.TabIndex = 9
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(489, 305)
        '
        'Label1
        '
        Me.Label1.Text = "Connection definition"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(503, 39)
        Me.Label2.Text = "Enter a unique connection name within a system."
        '
        'txtConnectionDesc
        '
        Me.txtConnectionDesc.Location = New System.Drawing.Point(120, 205)
        Me.txtConnectionDesc.MaxLength = 1000
        Me.txtConnectionDesc.Multiline = True
        Me.txtConnectionDesc.Name = "txtConnectionDesc"
        Me.txtConnectionDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtConnectionDesc.Size = New System.Drawing.Size(416, 59)
        Me.txtConnectionDesc.TabIndex = 7
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(24, 205)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 17)
        Me.Label8.TabIndex = 38
        Me.Label8.Text = "Description"
        '
        'cmbCnnType
        '
        Me.cmbCnnType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCnnType.Location = New System.Drawing.Point(396, 85)
        Me.cmbCnnType.Name = "cmbCnnType"
        Me.cmbCnnType.Size = New System.Drawing.Size(171, 21)
        Me.cmbCnnType.TabIndex = 2
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(288, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(102, 20)
        Me.Label7.TabIndex = 36
        Me.Label7.Text = "Connection Type"
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(120, 125)
        Me.txtDatabase.MaxLength = 128
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(152, 20)
        Me.txtDatabase.TabIndex = 3
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(24, 128)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(96, 17)
        Me.Label5.TabIndex = 34
        Me.Label5.Text = "Database"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(396, 163)
        Me.txtPassword.MaxLength = 128
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(171, 20)
        Me.txtPassword.TabIndex = 6
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(288, 166)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 20)
        Me.Label6.TabIndex = 32
        Me.Label6.Text = "Password"
        '
        'txtUserId
        '
        Me.txtUserId.Location = New System.Drawing.Point(120, 163)
        Me.txtUserId.MaxLength = 128
        Me.txtUserId.Name = "txtUserId"
        Me.txtUserId.Size = New System.Drawing.Size(152, 20)
        Me.txtUserId.TabIndex = 5
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 166)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(96, 20)
        Me.Label4.TabIndex = 30
        Me.Label4.Text = "Username"
        '
        'txtConnectionName
        '
        Me.txtConnectionName.Location = New System.Drawing.Point(120, 85)
        Me.txtConnectionName.MaxLength = 128
        Me.txtConnectionName.Name = "txtConnectionName"
        Me.txtConnectionName.Size = New System.Drawing.Size(152, 20)
        Me.txtConnectionName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(24, 88)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 32)
        Me.Label3.TabIndex = 28
        Me.Label3.Text = "Connection Name"
        '
        'txtDateformat
        '
        Me.txtDateformat.Location = New System.Drawing.Point(396, 125)
        Me.txtDateformat.MaxLength = 128
        Me.txtDateformat.Name = "txtDateformat"
        Me.txtDateformat.Size = New System.Drawing.Size(171, 20)
        Me.txtDateformat.TabIndex = 4
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(288, 128)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(84, 17)
        Me.Label9.TabIndex = 40
        Me.Label9.Text = "Date Format"
        '
        'cmbODBC
        '
        Me.cmbODBC.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbODBC.FormattingEnabled = True
        Me.cmbODBC.Location = New System.Drawing.Point(120, 124)
        Me.cmbODBC.Name = "cmbODBC"
        Me.cmbODBC.Size = New System.Drawing.Size(152, 21)
        Me.cmbODBC.TabIndex = 58
        '
        'frmConnection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(579, 344)
        Me.Controls.Add(Me.cmbODBC)
        Me.Controls.Add(Me.txtDateformat)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtConnectionDesc)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.cmbCnnType)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtDatabase)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtPassword)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtUserId)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtConnectionName)
        Me.Controls.Add(Me.Label3)
        Me.MinimumSize = New System.Drawing.Size(564, 378)
        Me.Name = "frmConnection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Connection Properties"
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        Me.Controls.SetChildIndex(Me.Label3, 0)
        Me.Controls.SetChildIndex(Me.txtConnectionName, 0)
        Me.Controls.SetChildIndex(Me.Label4, 0)
        Me.Controls.SetChildIndex(Me.txtUserId, 0)
        Me.Controls.SetChildIndex(Me.Label6, 0)
        Me.Controls.SetChildIndex(Me.txtPassword, 0)
        Me.Controls.SetChildIndex(Me.Label5, 0)
        Me.Controls.SetChildIndex(Me.txtDatabase, 0)
        Me.Controls.SetChildIndex(Me.Label7, 0)
        Me.Controls.SetChildIndex(Me.cmbCnnType, 0)
        Me.Controls.SetChildIndex(Me.Label8, 0)
        Me.Controls.SetChildIndex(Me.txtConnectionDesc, 0)
        Me.Controls.SetChildIndex(Me.Label9, 0)
        Me.Controls.SetChildIndex(Me.txtDateformat, 0)
        Me.Controls.SetChildIndex(Me.cmbODBC, 0)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtConnectionDesc.TextChanged, txtConnectionName.TextChanged, txtDatabase.TextChanged, txtDateformat.TextChanged, txtPassword.TextChanged, txtUserId.TextChanged, cmbODBC.SelectedIndexChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub OnCnnTypeChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCnnType.SelectedIndexChanged

        If cmbCnnType.Text = "ODBC" Then
            txtDatabase.Visible = False
            cmbODBC.Visible = True
        Else
            txtDatabase.Visible = True
            cmbODBC.Visible = False
        End If
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
            Dim EnvObj As clsEnvironment = objThis.Environment
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = "Conn_" & Count.ToString '& "_" & Strings.Left(EnvObj.Text, 9)
                For Each TestConn As clsConnection In EnvObj.Connections
                    If TestConn.ConnectionName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtConnectionName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmConn SetDefaultName")
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
    Public Function NewObj() As clsConnection

        IsNewObj = True

        StartLoad()

        SetDefaultName()

        InitControls()

        txtConnectionName.SelectAll()

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
            If ValidateNewName(txtConnectionName.Text) = False Or objThis.ValidateNewObject(txtConnectionName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.ConnectionDescription = txtConnectionDesc.Text
            objThis.ConnectionType = cmbCnnType.Text
            objThis.ConnectionName = txtConnectionName.Text

            If objThis.ConnectionType <> "ODBC" Then
                objThis.Database = txtDatabase.Text
            Else
                objThis.Database = cmbODBC.Text
            End If

            objThis.UserId = txtUserId.Text
            objThis.Password = txtPassword.Text
            objThis.DateFormat = txtDateformat.Text

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    DialogResult = Windows.Forms.DialogResult.Retry
                    Exit Sub
                End If
            Else
                objThis.Save()
            End If

            DialogResult = Windows.Forms.DialogResult.OK
            Me.Close()

        Catch ex As Exception
            DialogResult = Windows.Forms.DialogResult.Retry
            LogError(ex)
        End Try

    End Sub

    Sub InitControls()

        cmbCnnType.Items.Clear()
        cmbCnnType.Items.Add(New Mylist("ODBC", "ODBC"))
        cmbCnnType.Items.Add(New Mylist("ODBCwithSubstitutionVariable", "ODBCwithSubstitutionVariable"))
        cmbCnnType.Items.Add(New Mylist("DB2", "DB2"))
        cmbCnnType.Items.Add(New Mylist("OCI", "OCI"))
        cmbCnnType.Items.Add(New Mylist("NATIVEORI", "NATIVEORI"))
        cmbCnnType.SelectedIndex = 0

        SetComboConn()

        txtConnectionName.Focus()

    End Sub

    Function SetComboConn() As Boolean

        cmbODBC.Items.Clear()

        Dim tempDBName As String

        Try
            If getDSNlist() = True Then
                For Each dr As DataRow In mDSNTable.Rows
                    If dr(0).ToString <> "" Then
                        tempDBName = dr(0).ToString
                        cmbODBC.Items.Add(New Mylist(tempDBName, tempDBName))
                    End If
                Next
            End If

            cmbODBC.SelectedIndex = 0

            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLInfo SetComboConn")
            Return False
        End Try

    End Function

    Private Function getDSNlist() As Boolean
        'GET THE DSN LIST
        Dim ret As Boolean

        ret = ODBCHelper.BuildDSNDataTable()
        If ret = True Then
            mDSNTable = ODBCHelper.DSNDataTable
            Return True
        Else
            Return False
        End If

    End Function

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        ShowHelp(modHelp.HHId.H_Add_Connections)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Select Case e.KeyCode
            Case Keys.Escape
                Call cmdCancel_Click(sender, e)
            Case Keys.Enter
                Call cmdOk_Click(sender, e)
            Case Keys.F1
                Call cmdHelp_Click(sender, e)
        End Select
    End Sub

End Class
