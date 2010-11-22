Public Class ctlFolderNode
    Inherits System.Windows.Forms.UserControl

    Dim objThis As New clsFolderNode
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
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lv As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmdClose = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lv = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.SuspendLayout()
        '
        'cmdClose
        '
        Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.cmdClose.Location = New System.Drawing.Point(432, 296)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(72, 24)
        Me.cmdClose.TabIndex = 31
        Me.cmdClose.Text = "&Close"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(8, 280)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(496, 7)
        Me.GroupBox1.TabIndex = 30
        Me.GroupBox1.TabStop = False
        '
        'lv
        '
        Me.lv.BackColor = System.Drawing.Color.White
        Me.lv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lv.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.lv.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lv.Location = New System.Drawing.Point(0, 0)
        Me.lv.Name = "lv"
        Me.lv.Size = New System.Drawing.Size(512, 328)
        Me.lv.TabIndex = 32
        Me.lv.UseCompatibleStateImageBehavior = False
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = ""
        Me.ColumnHeader1.Width = 271
        '
        'ctlFolderNode
        '
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.Controls.Add(Me.lv)
        Me.Controls.Add(Me.cmdClose)
        Me.Controls.Add(Me.GroupBox1)
        Me.ForeColor = System.Drawing.Color.White
        Me.Name = "ctlFolderNode"
        Me.Size = New System.Drawing.Size(512, 328)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub StartLoad(ByVal cNode As TreeNode)
        '//Unload old object before we load new object
        objThis = Nothing
        objThis = New clsFolderNode(CType(cNode.Tag, INode).Text, CType(cNode.Tag, INode).Type)

        ClearControls(Me.Controls)
    End Sub

    Private Sub EndLoad()
        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False
    End Sub

    Public Sub cmdClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdClose.Click
        Try
            Me.Visible = False
        Catch ex As Exception
            LogError(ex)
        End Try

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles lv.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.Enter
                lv_DoubleClick(sender, New EventArgs)
            Case Keys.F1
                If lv.SelectedItems.Count > 0 Then
                    Try
                        Dim currentNode As TreeNode = CType(lv.SelectedItems(0).Tag, TreeNode)
                        Dim helpInode As INode = currentNode.Tag
                        Select Case helpInode.Type
                            Case NODE_PROJECT
                                ShowHelp(modHelp.HHId.H_Project)
                            Case NODE_ENVIRONMENT, NODE_FO_ENVIRONMENT
                                ShowHelp(modHelp.HHId.H_Environment)
                            Case NODE_STRUCT
                                Dim StructType As enumStructure
                                StructType = CType(currentNode.Tag, clsStructure).StructureType
                                Select Case StructType
                                    Case enumStructure.STRUCT_COBOL
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case enumStructure.STRUCT_C
                                        ShowHelp(modHelp.HHId.H_C_Structure)
                                    Case enumStructure.STRUCT_COBOL_IMS
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case enumStructure.STRUCT_IMS
                                        ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                                    Case enumStructure.STRUCT_REL_DDL
                                        ShowHelp(modHelp.HHId.H_Relational_DDL)
                                    Case enumStructure.STRUCT_REL_DML
                                        ShowHelp(modHelp.HHId.H_Relational_DML)
                                    Case enumStructure.STRUCT_XMLDTD
                                        ShowHelp(modHelp.HHId.H_XML_DTD)
                                    Case enumStructure.STRUCT_UNKNOWN
                                        ShowHelp(modHelp.HHId.H_Structures)
                                    Case Else
                                        ShowHelp(modHelp.HHId.H_Structures)
                                End Select
                            Case NODE_FO_STRUCT
                                Dim FolderType As String = helpInode.Text
                                Select Case FolderType
                                    Case "COBOL"
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case "CHeader"
                                        ShowHelp(modHelp.HHId.H_C_Structure)
                                    Case "COBOLIMS"
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case "IMS"
                                        ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                                    Case "DDL"
                                        ShowHelp(modHelp.HHId.H_Relational_DDL)
                                    Case "DML"
                                        ShowHelp(modHelp.HHId.H_Relational_DML)
                                    Case "XMLDTD"
                                        ShowHelp(modHelp.HHId.H_XML_DTD)
                                    Case "Unknown"
                                        ShowHelp(modHelp.HHId.H_Structures)
                                    Case Else
                                        ShowHelp(modHelp.HHId.H_Structures)
                                End Select
                            Case NODE_STRUCT_FLD
                                ShowHelp(modHelp.HHId.H_Structure_Field_Attributes)
                            Case NODE_STRUCT_SEL, NODE_FO_STRUCT_SEL
                                Dim ParentStructType As enumStructure
                                ParentStructType = CType(currentNode.Tag, clsStructure).Parent.Type
                                Select Case ParentStructType
                                    Case enumStructure.STRUCT_COBOL
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case enumStructure.STRUCT_C
                                        ShowHelp(modHelp.HHId.H_C_Structure)
                                    Case enumStructure.STRUCT_COBOL_IMS
                                        ShowHelp(modHelp.HHId.H_COBOL_Copybook__IMS_DBD)
                                    Case enumStructure.STRUCT_IMS
                                        ShowHelp(modHelp.HHId.H_Stru_IMS_Segment)
                                    Case enumStructure.STRUCT_REL_DDL
                                        ShowHelp(modHelp.HHId.H_Relational_DDL)
                                    Case enumStructure.STRUCT_REL_DML
                                        ShowHelp(modHelp.HHId.H_Relational_DML)
                                    Case enumStructure.STRUCT_XMLDTD
                                        ShowHelp(modHelp.HHId.H_XML_DTD)
                                    Case enumStructure.STRUCT_UNKNOWN
                                        ShowHelp(modHelp.HHId.H_Structures)
                                    Case Else
                                        ShowHelp(modHelp.HHId.H_Structures)
                                End Select
                            Case NODE_SYSTEM, NODE_FO_SYSTEM
                                ShowHelp(modHelp.HHId.H_Systems)
                            Case NODE_CONNECTION, NODE_FO_CONNECTION
                                ShowHelp(modHelp.HHId.H_Connections)
                            Case NODE_ENGINE, NODE_FO_ENGINE
                                ShowHelp(modHelp.HHId.H_Engines)
                            Case NODE_SOURCEDATASTORE, NODE_FO_SOURCEDATASTORE, NODE_SOURCEDSSEL
                                ShowHelp(modHelp.HHId.H_Sources)
                            Case NODE_TARGETDATASTORE, NODE_FO_TARGETDATASTORE, NODE_TARGETDSSEL
                                ShowHelp(modHelp.HHId.H_Targets)
                            Case NODE_VARIABLE, NODE_FO_VARIABLE
                                ShowHelp(modHelp.HHId.H_Variables)
                            Case NODE_PROC, NODE_FO_PROC
                                ShowHelp(modHelp.HHId.H_Add_Mapping_Procedures)
                            Case NODE_GEN, NODE_FO_JOIN
                                ShowHelp(modHelp.HHId.H_General_Procedure)
                            Case NODE_LOOKUP, NODE_FO_LOOKUP
                                ShowHelp(modHelp.HHId.H_Lookups)
                            Case NODE_MAIN, NODE_FO_MAIN
                                ShowHelp(modHelp.HHId.H_Main_Procedure)
                            Case NODE_MAPPING, NODE_FO_MAPPING
                                ShowHelp(modHelp.HHId.H_MAP)
                            Case NODE_FUN, NODE_FO_FUNCTION, NODE_FO_FUNCTION_RECENT
                                ShowHelp(modHelp.HHId.H_Function_Reference)
                            Case NODE_TEMPLATE, NODE_FO_TEMPLATE
                                ShowHelp(modHelp.HHId.H_Reference)
                            Case Else
                                ShowHelp(modHelp.HHId.H_DATETIME)
                        End Select
                    Catch ex As Exception
                        LogError(ex)
                        Exit Sub
                    End Try
                End If
        End Select
    End Sub

    Public Function EditObj(ByVal cNode As TreeNode, ByVal obj As INode) As clsFolderNode

        IsNewObj = False

        StartLoad(cNode)

        objThis = obj '//Load the form env object

        LoadList(cNode)

        EditObj = objThis

        EndLoad()

    End Function

    Public Function Clear() As Boolean

        lv.Items.Clear()

    End Function

    '//Set values from objProject to form controls
    Private Function LoadList(ByVal cNode As TreeNode) As Boolean

        Dim nd As TreeNode
        lv.Items.Clear()

        lv.SmallImageList = imgListSmall
        lv.LargeImageList = imgListSmall

        For Each nd In cNode.Nodes
            With lv.Items.Add(CType(nd.Tag, INode).Text)
                If CType(nd.Tag, INode).Type = NODE_SOURCEDATASTORE Then
                    .ImageIndex = ImgIdxFromName(CType(nd.Tag, INode).Type, CType(nd.Tag, INode))
                Else
                    .ImageIndex = ImgIdxFromName(CType(nd.Tag, INode).Type)
                End If
                .Tag = nd
            End With
        Next

    End Function

    Private Sub lv_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lv.SelectedIndexChanged

    End Sub

    Private Sub lv_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lv.DoubleClick

        Try
            If lv.SelectedItems.Count > 0 Then
                CType(lv.SelectedItems(0).Tag, TreeNode).TreeView.SelectedNode = lv.SelectedItems(0).Tag
                CType(lv.SelectedItems(0).Tag, TreeNode).EnsureVisible()
            End If

        Catch ex As Exception
        End Try

    End Sub

End Class
