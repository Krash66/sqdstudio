Public Class frmStructSelection
    Inherits SQDStudio.frmBlank

    Public objThis As New clsStructureSelection

    Dim IsNewObj As Boolean
    Dim lastTreeviewSearchText As String
    Dim lastTreeviewSearchNode As TreeNode
    Dim colSkipNodes As New ArrayList

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

#Region " Windows Form Designer generated code "

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub


    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents pnlFields As System.Windows.Forms.Panel
    Friend WithEvents lblFieldName As System.Windows.Forms.Label
    Friend WithEvents cmdSearch As System.Windows.Forms.Button
    Friend WithEvents txtSearchField As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tvFields As System.Windows.Forms.TreeView
    Friend WithEvents lvFieldAttrs As System.Windows.Forms.ListView
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents pnlSelect As System.Windows.Forms.Panel
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtSelectionDesc As System.Windows.Forms.TextBox
    Friend WithEvents txtSelectionName As System.Windows.Forms.TextBox
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents column1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents column2 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmStructSelection))
        Me.pnlFields = New System.Windows.Forms.Panel
        Me.lblFieldName = New System.Windows.Forms.Label
        Me.cmdSearch = New System.Windows.Forms.Button
        Me.txtSearchField = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.tvFields = New System.Windows.Forms.TreeView
        Me.lvFieldAttrs = New System.Windows.Forms.ListView
        Me.column1 = New System.Windows.Forms.ColumnHeader
        Me.column2 = New System.Windows.Forms.ColumnHeader
        Me.Label4 = New System.Windows.Forms.Label
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.pnlSelect = New System.Windows.Forms.Panel
        Me.txtSelectionDesc = New System.Windows.Forms.TextBox
        Me.lblFile = New System.Windows.Forms.Label
        Me.txtSelectionName = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlFields.SuspendLayout()
        Me.pnlSelect.SuspendLayout()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.ImageLocation = "C:\Documents and Settings\tkarasc\My Documents\Visual Studio 2005\Projects\sqdstu" & _
            "dio\images\FormTop\sq_skyblue.jpg"
        '
        'Panel1
        '
        Me.Panel1.Size = New System.Drawing.Size(592, 68)
        '
        'GroupBox1
        '
        Me.GroupBox1.Location = New System.Drawing.Point(1, 433)
        Me.GroupBox1.Size = New System.Drawing.Size(578, 10)
        '
        'cmdOk
        '
        Me.cmdOk.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdOk.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdOk.Location = New System.Drawing.Point(322, 449)
        Me.cmdOk.TabIndex = 5
        '
        'cmdCancel
        '
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdCancel.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdCancel.Location = New System.Drawing.Point(409, 449)
        Me.cmdCancel.TabIndex = 6
        '
        'cmdHelp
        '
        Me.cmdHelp.FlatAppearance.BorderColor = System.Drawing.Color.Silver
        Me.cmdHelp.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.cmdHelp.Location = New System.Drawing.Point(499, 449)
        Me.cmdHelp.TabIndex = 7
        '
        'Label1
        '
        Me.Label1.Text = "Structure Subset Properties"
        '
        'Label2
        '
        Me.Label2.Size = New System.Drawing.Size(516, 39)
        Me.Label2.Text = "Select a Subset of your structure "
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
        Me.pnlFields.Controls.Add(Me.tvFields)
        Me.pnlFields.Controls.Add(Me.lvFieldAttrs)
        Me.pnlFields.Controls.Add(Me.Label4)
        Me.pnlFields.Location = New System.Drawing.Point(12, 166)
        Me.pnlFields.Name = "pnlFields"
        Me.pnlFields.Padding = New System.Windows.Forms.Padding(5)
        Me.pnlFields.Size = New System.Drawing.Size(568, 261)
        Me.pnlFields.TabIndex = 34
        '
        'lblFieldName
        '
        Me.lblFieldName.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFieldName.ForeColor = System.Drawing.Color.Blue
        Me.lblFieldName.Location = New System.Drawing.Point(322, 10)
        Me.lblFieldName.Name = "lblFieldName"
        Me.lblFieldName.Size = New System.Drawing.Size(170, 16)
        Me.lblFieldName.TabIndex = 38
        '
        'cmdSearch
        '
        Me.cmdSearch.Location = New System.Drawing.Point(197, 7)
        Me.cmdSearch.Name = "cmdSearch"
        Me.cmdSearch.Size = New System.Drawing.Size(30, 20)
        Me.cmdSearch.TabIndex = 37
        Me.cmdSearch.Text = "Go"
        '
        'txtSearchField
        '
        Me.txtSearchField.Location = New System.Drawing.Point(55, 7)
        Me.txtSearchField.MaxLength = 128
        Me.txtSearchField.Name = "txtSearchField"
        Me.txtSearchField.Size = New System.Drawing.Size(136, 20)
        Me.txtSearchField.TabIndex = 36
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(8, 10)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(46, 12)
        Me.Label5.TabIndex = 35
        Me.Label5.Text = "Search"
        '
        'tvFields
        '
        Me.tvFields.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tvFields.CheckBoxes = True
        Me.tvFields.Location = New System.Drawing.Point(8, 32)
        Me.tvFields.Name = "tvFields"
        Me.tvFields.Size = New System.Drawing.Size(219, 221)
        Me.tvFields.TabIndex = 3
        '
        'lvFieldAttrs
        '
        Me.lvFieldAttrs.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvFieldAttrs.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.column1, Me.column2})
        Me.lvFieldAttrs.FullRowSelect = True
        Me.lvFieldAttrs.Location = New System.Drawing.Point(229, 32)
        Me.lvFieldAttrs.Name = "lvFieldAttrs"
        Me.lvFieldAttrs.Size = New System.Drawing.Size(331, 221)
        Me.lvFieldAttrs.TabIndex = 4
        Me.lvFieldAttrs.UseCompatibleStateImageBehavior = False
        Me.lvFieldAttrs.View = System.Windows.Forms.View.Details
        '
        'column1
        '
        Me.column1.Text = "Attribute"
        Me.column1.Width = 160
        '
        'column2
        '
        Me.column2.Text = "Value"
        Me.column2.Width = 150
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(243, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(81, 12)
        Me.Label4.TabIndex = 32
        Me.Label4.Text = "Attributes"
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
        Me.pnlSelect.Controls.Add(Me.txtSelectionDesc)
        Me.pnlSelect.Controls.Add(Me.lblFile)
        Me.pnlSelect.Controls.Add(Me.txtSelectionName)
        Me.pnlSelect.Controls.Add(Me.Label3)
        Me.pnlSelect.Location = New System.Drawing.Point(12, 74)
        Me.pnlSelect.Name = "pnlSelect"
        Me.pnlSelect.Size = New System.Drawing.Size(568, 86)
        Me.pnlSelect.TabIndex = 35
        '
        'txtSelectionDesc
        '
        Me.txtSelectionDesc.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSelectionDesc.Location = New System.Drawing.Point(109, 32)
        Me.txtSelectionDesc.MaxLength = 1000
        Me.txtSelectionDesc.Multiline = True
        Me.txtSelectionDesc.Name = "txtSelectionDesc"
        Me.txtSelectionDesc.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtSelectionDesc.Size = New System.Drawing.Size(383, 51)
        Me.txtSelectionDesc.TabIndex = 2
        '
        'lblFile
        '
        Me.lblFile.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFile.Location = New System.Drawing.Point(8, 35)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(78, 20)
        Me.lblFile.TabIndex = 65
        Me.lblFile.Text = "Description"
        '
        'txtSelectionName
        '
        Me.txtSelectionName.Location = New System.Drawing.Point(109, 7)
        Me.txtSelectionName.MaxLength = 128
        Me.txtSelectionName.Name = "txtSelectionName"
        Me.txtSelectionName.Size = New System.Drawing.Size(119, 20)
        Me.txtSelectionName.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 10)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(90, 20)
        Me.Label3.TabIndex = 55
        Me.Label3.Text = "Subset  Name"
        '
        'frmStructSelection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 485)
        Me.Controls.Add(Me.pnlSelect)
        Me.Controls.Add(Me.pnlFields)
        Me.MinimumSize = New System.Drawing.Size(600, 472)
        Me.Name = "frmStructSelection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Structure Subset Properties"
        Me.Controls.SetChildIndex(Me.pnlFields, 0)
        Me.Controls.SetChildIndex(Me.cmdHelp, 0)
        Me.Controls.SetChildIndex(Me.Panel1, 0)
        Me.Controls.SetChildIndex(Me.GroupBox1, 0)
        Me.Controls.SetChildIndex(Me.cmdOk, 0)
        Me.Controls.SetChildIndex(Me.cmdCancel, 0)
        'Me.Controls.SetChildIndex(Me.PictureBox1, 0)
        Me.Controls.SetChildIndex(Me.pnlSelect, 0)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlFields.ResumeLayout(False)
        Me.pnlFields.PerformLayout()
        Me.pnlSelect.ResumeLayout(False)
        Me.pnlSelect.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSelectionDesc.TextChanged, txtSelectionName.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

    Private Sub SetDefaultName()

        Try
            Dim StrObj As clsStructure = objThis.ObjStructure
            Dim NewName As String = ""
            Dim Count As Integer = 1
            Dim GoodName As Boolean = False

            While GoodName = False
                GoodName = True
                NewName = StrObj.Text & "_Sub" & Count.ToString
                For Each TestSS As clsStructureSelection In StrObj.StructureSelections
                    If TestSS.SelectionName = NewName Then
                        GoodName = False
                        Exit For
                    End If
                Next
                Count += 1
            End While

            txtSelectionName.Text = NewName

        Catch ex As Exception
            LogError(ex, "frmStrSel SetDefaultName")
        End Try

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

    '//Returns new SS object
    Friend Function NewObj(Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN) As clsStructureSelection

        Try
            IsNewObj = True

            StartLoad()

            SetDefaultName()

            InitControls()

            ShowHideFieldAttr()

            txtSelectionName.SelectAll()

            EndLoad()

doAgain:
            Select Case Me.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    NewObj = objThis
                Case Windows.Forms.DialogResult.Retry
                    GoTo doAgain
                Case Else
                    NewObj = Nothing
            End Select

        Catch ex As Exception
            LogError(ex, "frmStructSelect NewObj")
            NewObj = Nothing
        End Try

    End Function

    Public Overrides Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If objThis.ValidateNewObject(txtSelectionName.Text) = False Or ValidateNewName(txtSelectionName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            objThis.SelectionName = txtSelectionName.Text
            objThis.SelectionDescription = txtSelectionDesc.Text

            '//save current selected fields before we do add/save
            LoadFldArrFromTreeview(objThis.ObjSelectionFields, tvFields, True)

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

    Function InitControls() As Boolean

        '//Add listview columns
        tvFields.HideSelection = False

        lvFieldAttrs.SmallImageList = imgListSmall

        pnlSelect.Visible = True
        pnlFields.Visible = True

    End Function

    Private Function ShowHideFieldAttr() As Boolean

        Try
            '//Dont reload and reprocess it once its loaded
            If tvFields.GetNodeCount(True) > 0 Then Exit Function
            'End If

            tvFields.BackColor = Color.LightBlue
            lvFieldAttrs.BackColor = Color.LightBlue
            lblFieldName.ForeColor = Color.Red
            lblFieldName.Text = "********* Loading *********"

            Me.Refresh()

            '//FirstLoad all field of parent structure and then check only selected for this selection
            Call FillFieldAttr()

            tvFields.BackColor = Color.White
            lvFieldAttrs.BackColor = Color.White
            lblFieldName.Text = ""
            lblFieldName.ForeColor = Color.Blue

            '//select first itm
            If tvFields.GetNodeCount(True) > 0 Then
                tvFields.SelectedNode = tvFields.Nodes.Item(0)
            End If

        Catch ex As Exception
            LogError(ex, "frmSS ShowHideFieldAttr")
        End Try

    End Function

    '//This will load treeview of fields
    Function FillFieldAttr() As Boolean

        Try
            tvFields.BeginUpdate()
            lvFieldAttrs.BeginUpdate()
            tvFields.Nodes.Clear()

            objThis.ObjStructure.LoadItems()

            '//Load all fields from structure
            AddFieldsToTreeView(objThis.ObjStructure, , tvFields)

            '//Now load selected fields and check if object is already created and user is editing it
            If IsNewObj = False Then objThis.LoadMe()
            CheckSelectedFields()

            Return True

        Catch ex As Exception
            LogError(ex, "frmSS FillFieldAttr")
            Return False
        Finally
            tvFields.EndUpdate()
            lvFieldAttrs.EndUpdate()
        End Try

    End Function

    Function CheckSelectedFields() As Boolean
        '//TODO
        Dim i As Integer
        Dim cnt As Integer = tvFields.GetNodeCount(True)

        If cnt <= 0 Then Exit Function
        Try
            tvFields.BeginUpdate()
            '//Looop all selected field and check that field in treeview
            For i = 0 To objThis.ObjSelectionFields.Count - 1
                SelectFirstMatchingNode(tvFields, "", True, CType(objThis.ObjSelectionFields(i), clsField).Key)
            Next
            Return True

        Catch ex As Exception
            LogError(ex, "frmSS ChkSelectedFlds")
            Return False
        Finally
            tvFields.EndUpdate()
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
            lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY))


            'lvFieldAttrs.Items.Add(modDeclares.TXT_ISKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_ISKEY))
            'lvFieldAttrs.Items.Add(modDeclares.TXT_FKEY).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_FKEY)) '// added by TK and KS 11/6/2006
            lvFieldAttrs.Items.Add(modDeclares.TXT_INITVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INITVAL))
            lvFieldAttrs.Items.Add(modDeclares.TXT_RETYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_RETYPE))

            lvFieldAttrs.Items.Add(modDeclares.TXT_EXTTYPE).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_EXTTYPE))
            lvFieldAttrs.Items.Add(modDeclares.TXT_INVALID).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_INVALID))

            lvFieldAttrs.Items.Add(modDeclares.TXT_DATEFORMAT).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_DATEFORMAT))
            lvFieldAttrs.Items.Add(modDeclares.TXT_LABEL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_LABEL))
            lvFieldAttrs.Items.Add(modDeclares.TXT_IDENTVAL).SubItems.Add(CType(cNode.Tag, clsField).GetFieldAttr(modDeclares.enumFieldAttributes.ATTR_IDENTVAL))

            Dim f As New System.Drawing.Font(lvFieldAttrs.Font, FontStyle.Bold)
            Dim i As Integer

            For i = 0 To lvFieldAttrs.Items.Count - 1
                lvFieldAttrs.Items(i).UseItemStyleForSubItems = False
                lvFieldAttrs.Items(i).SubItems(0).Font = f
                lvFieldAttrs.Items(i).SubItems(0).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).SubItems(1).ForeColor = Color.Gray
                lvFieldAttrs.Items(i).ImageIndex = 30
            Next

            lblFieldName.Text = cNode.Text

            lvFieldAttrs.EndUpdate()
            Return True

        Catch ex As Exception
            LogError(ex, "frmSS ShowFieldAttr")
            Return False
        End Try

    End Function

    Private Sub tvFields_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterSelect

        ShowFieldAttributes(e.Node)

    End Sub

    Private Sub txtSearchField_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSearchField.TextChanged

        SelectFirstMatchingNode(tvFields, txtSearchField.Text)

    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        If SelectFirstMatchingNode(tvFields, colSkipNodes, txtSearchField.Text) = False Then
            MsgBox("No matching node found for entered text", MsgBoxStyle.Critical, MsgTitle)
        End If

    End Sub

    Private Sub tvFields_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFields.AfterCheck

        '//Only process this logic if action using mouse or keyboard other wise this event can fire on and on ....
        If IsEventFromCode = False Then
            CheckUncheckNodes(e.Node)
        End If

    End Sub

    Public Overrides Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Add_Structures)

    End Sub

    Public Overrides Sub MyForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

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
