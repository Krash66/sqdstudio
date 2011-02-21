Public Class ctlMain
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    Dim objThis As New clsTask
    Dim IsNewObj As Boolean
    Dim curEditType As enumDirection
    Dim firstTime As Boolean = False

    '//Object pointing to the SQFunction which is being edited
    Dim objCurMap As clsMapping
    Dim objCurFunct As clsSQFunction
    Dim ObjCurFld As clsField
    Dim IsCodeEditorOnTop As Boolean
    Dim IsDescEditorOnTop As Boolean

    '/// Colors for highlighting and text
    Dim MAPPED_ITEM_COLOR As Color = Color.SkyBlue
    Dim DL_mapped_Item_Color As Color = Color.LavenderBlush
    Dim Maybe_mapped_Item_Color As Color = Color.Lavender
    Dim MISSING_DEP_ITEM_COLOR As Color = Color.Yellow
    'Dim SELECTED_CELL_COLOR As Color = Color.Green
    Dim BACK_COLOR As Color = Color.White
    Dim HAS_DESC_COLOR As Color = Color.Blue  '// text color if mapping has a mapping description

    '/// for listview navigation
    Dim PrevCol As Integer = 0
    Dim PrevRow As Integer = 0
    Dim CurCol As Integer = 0
    Dim CurRow As Integer = 0
    Dim MousePos As Point

    Dim IsClipboardAvail As Boolean
    Dim objClip As Object
    Dim TempMap As clsMapping
    Dim lastMappingSearchPos As Integer = 0
    Dim lastMappingSearchText As String = ""
    Dim lastTreeviewSearchText As String
    Dim lastTreeviewSearchNode As TreeNode
    Dim DoSetTaskType As Boolean


    Dim colSkipNodes As New ArrayList


#Region "form events"

    Private Sub StartLoad()

        Me.Visible = False
        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
        txtTaskName.Enabled = True '//added on 7/24 to disable object name editing

        ClearControls(Me.Controls, modDeclares.enumControlTypes.ControlTreeView)
        'HideScriptEditor(True) '//8/16/05
        'HideDescEditor(True)  '//4/10/07
        LoadFunctionFromXML(tvFunctions)

        gbName.ForeColor = Color.White
        gbFun.ForeColor = Color.White
        'gbTgt.ForeColor = Color.White
        gbSrc.ForeColor = Color.White
        gbMap.ForeColor = Color.White
        'gbFldAttr.ForeColor = Color.White
        objCurMap = New clsMapping
        objCurFunct = New clsSQFunction
        DoSetTaskType = False

    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        IsEventFromCode = False

    End Sub

    'Private Sub ctlTask_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize

    '    'pnlScript.Size = pnlSelect.Size
    '    'pnlDesc.Size = pnlSelect.Size
    '    'pnlScript.Location = pnlSelect.Location
    '    'pnlDesc.Location = pnlSelect.Location

    'End Sub

    'Private Sub ctlTask_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

    '    'lvMappings.SmallImageList = imgListSmall

    'End Sub

    'Private Sub Panel1_Resize(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlSourceTarget.Resize

    '    'tvSource.Width = (pnlSourceTarget.Width / 2) - 5
    '    'tvTarget.Left = (pnlSourceTarget.Width / 2) + 5
    '    'tvTarget.Width = (pnlSourceTarget.Width / 2) - 5
    '    'lblTargetCaption.Left = tvTarget.Left

    'End Sub

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTaskDesc.TextChanged, txtTaskName.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    'Private Sub OnCodeChange(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If IsEventFromCode = True Then Exit Sub

    '    objCurMap.IsModified = True
    '    objThis.IsModified = True
    '    cmdSave.Enabled = True

    '    OnChange(Me, New EventArgs)

    'End Sub

    'Private Sub OnDescChange(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '
    '    If IsEventFromCode = True Then Exit Sub
    '
    '    objCurMap.IsModified = True
    '    objThis.IsModified = True
    '    cmdSave.Enabled = True
    '
    '    RaiseEvent Modified(Me, objThis)
    '
    'End Sub

    Public Function EditObj(ByVal cNode As TreeNode, ByVal obj As INode) As clsTask

        IsNewObj = False
        StartLoad()

        objThis = obj '//Load the form env object

        'If objThis.Engine IsNot Nothing Then
        '    cbGroupItems.Checked = objThis.Engine.MapGroupItems
        'End If

        objThis.ObjTreeNode = cNode

        SetTaskTitle(objThis.TaskType)

        'If objThis.TaskType = modDeclares.enumTaskType.TASK_JOIN Or objThis.TaskType = modDeclares.enumTaskType.TASK_LOOKUP Then
        '    'tvTarget.Enabled = False
        '    tvTarget.Visible = False
        '    lblTargetCaption.Visible = False
        'Else
        '    lblTargetCaption.Visible = True
        '    tvTarget.Visible = True
        'End If

        UpdateFields()

        EditObj = objThis

        FillMappings()

        EndLoad()

        If CType(cNode.Tag, INode).Type = NODE_ENGINE Then
            FillSourceTarget(cNode)
        Else
            FillSrcTgtTop(cNode)
        End If


        'the following if statement is to relode the source and target for the first time 
        'If firstTime = True Then
        '    FillSourceTarget(objThis.ObjTreeNode)
        '    firstTime = False
        'End If

        If tvSource.Nodes.Count <> 0 Then
            tvSource.CollapseAll()
            tvSource.TopNode.Expand()
        End If
        'If tvTarget.Nodes.Count <> 0 Then
        '    tvTarget.CollapseAll()
        '    tvTarget.TopNode.Expand()
        'End If

        '/// Now get Last Mapped Fields
        'If SetLastMapped() = True Then
        '    HiLiteLastSrcTgtFlds(tvSource.TopNode, objThis.LastSrcFld, True, tvSource)
        '    HiLiteLastSrcTgtFlds(tvTarget.TopNode, objThis.LastTgtFld, True, tvTarget)
        'End If

        'Me.Refresh()

        Me.Visible = True

    End Function

    'Private Sub OngbMap_resize(ByVal sender As Object, ByVal e As EventArgs) Handles gbMap.Resize

    '    'lvMappings.Columns(0).Width = (lvMappings.Width / 2) - 12
    '    'lvMappings.Columns(1).Width = (lvMappings.Width / 2) - 12

    'End Sub

    Sub SetTaskTitle(ByVal Type As enumTaskType)

        Try
            Select Case Type
                Case enumTaskType.TASK_GEN
                    gbName.Text = "General Procedure"
                Case enumTaskType.TASK_LOOKUP
                    gbName.Text = "LookUp"
                Case enumTaskType.TASK_MAIN
                    gbName.Text = "Main Procedure"
                Case enumTaskType.TASK_PROC
                    gbName.Text = "Mapping Procedure"
            End Select

        Catch ex As Exception
            LogError(ex, "ctlMain setTaskTitle")
        End Try

    End Sub

#End Region

#Region "Save"

    Public Function Save() As Boolean

        Try
            Dim OldName As String = ""

            '// First Check Validity before Saving
            If ValidateNewName(txtTaskName.Text) = False Then
                Exit Function
            End If
            '// check to see if renamed
            If objThis.TaskName <> txtTaskName.Text Then
                objThis.IsRenamed = RenameTask(objThis, txtTaskName.Text)
            End If
            '// now set the task name
            If objThis.IsRenamed = False Then
                txtTaskName.Text = objThis.TaskName
            Else
                objThis.TaskName = txtTaskName.Text
            End If
            '// save all of the task properties from form to object
            If SaveCurrentSelection() = False Then
                Exit Function
            End If
            '/// Added save as main or General Procedure
            If DoSetTaskType = True Then
                If objThis.SaveTaskType() = False Then
                    Exit Function
                End If
            End If
            'If SaveLastMapped() = False Then
            '    Exit Function
            'End If
            '// add if it's new, save if not new
            If IsNewObj = True Then
                If objThis.AddNew = False Then Exit Function
            Else
                If objThis.Save() = False Then
                    Exit Function
                Else
                    If objThis.UpdateTaskDesc() = False Then
                        Exit Function
                        'Else
                        '    If objThis.Engine Is Nothing Then
                        '        For Each sys As clsSystem In objThis.Environment.Systems
                        '            For Each eng As clsEngine In sys.Engines
                        '                For Each task As clsTask In eng.Procs
                        '                    If OldName = "" Then
                        '                        OldName = objThis.TaskName
                        '                    End If
                        '                    If task.TaskName = OldName Then
                        '                        task.TaskName = objThis.TaskName
                        '                        task = objThis.Clone(eng)  'UpdateChildDS(src)
                        '                        If task IsNot Nothing Then
                        '                            If task.Save() = False Then
                        '                                MsgBox("Procedure: " & task.Text & " failed to update correctly in Engine: " & _
                        '                                eng.Text, MsgBoxStyle.Information, "Problem saving Procedure")
                        '                            End If
                        '                        Else
                        '                            MsgBox("Procedure: " & task.Text & " failed to update correctly in Engine: " & _
                        '                            eng.Text, MsgBoxStyle.Information, "Problem saving Procedure")
                        '                        End If
                        '                    End If
                        '                Next
                        '            Next
                        '        Next
                        '    End If
                    End If
                End If
            End If

            '// raise event so main form can update itself accordingly
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

    '//write form controls values to object
    Function SaveCurrentSelection() As Boolean

        Try

            Dim objMap As New clsMapping


            '// save structure selection
            'If IsCodeEditorOnTop = True Then
            '    'SaveCurrentScript()
            'End If

            'If IsDescEditorOnTop = True Then
            '    SaveMapDesc()
            'End If

            objThis.TaskDescription = txtTaskDesc.Text
            objCurFunct.SQFunctionWithInnerText = txtCodeEditor.Text
            objCurFunct.SQFunctionName = objCurFunct.SQFunctionWithInnerText

            '//before we clear old selection save it to another array so later on while saving this object we can compare for any new addition in selection list
            objThis.OldObjMappings.Clear()
            objThis.OldObjMappings = objThis.ObjMappings.Clone

            ''//clear previous selection and store new selection
            If objThis.ObjMappings.Count > 0 Then
                objThis.ObjMappings.Clear()
                objCurMap.MappingSource = objCurFunct
                objCurMap.SourceType = enumMappingType.MAPPING_TYPE_FUN
                objCurMap.SourceParent = objThis.TaskName
                objThis.ObjMappings.Add(objCurMap)
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlTask SaveCurrentSelection")
            Return False
        End Try

    End Function

    'Function SaveLastMapped() As Boolean

    '    Dim ErrorMsg As String = ""
    '    Try
    '        If objThis.SaveLastFlds() = False Then
    '            ErrorMsg = "An Error occured while saving Last Mapped Fields to a file."
    '        End If

    '        If ErrorMsg <> "" Then
    '            MsgBox(ErrorMsg, MsgBoxStyle.OkOnly, MsgTitle)
    '            Return False
    '            Exit Function
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask SaveLastMapped")
    '        Return False
    '    End Try

    'End Function

    '////// Modify Here. !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'Function SaveCurrentScript() As Boolean

    '    Dim fieldType As modDeclares.enumMappingType

    '    Try
    '        If curEditType = modDeclares.enumDirection.DI_SOURCE Then
    '            fieldType = GetTypeFromText(txtCodeEditor.Text, modDeclares.enumDirection.DI_SOURCE)
    '            If txtCodeEditor.Text <> "" And fieldType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                fieldType = modDeclares.enumMappingType.MAPPING_TYPE_FUN  '/// mod 12/6/07  TK
    '            End If
    '            Select Case fieldType
    '                Case modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT
    '                    Dim c As New clsVariable(txtCodeEditor.Text, modDeclares.enumVariableType.VTYPE_CONST)

    '                    objCurMap.MappingSource = Nothing
    '                    objCurMap.MappingSource = c
    '                    objCurMap.SourceType = fieldType
    '                    'lvMappings.Items(objCurMap.SeqNo).Text = txtCodeEditor.Text
    '                    'lvMappings.Items(objCurMap.SeqNo).ImageIndex = ImgIdxFromName(objCurMap.MappingSource.Type)

    '                    objCurMap.IsModified = True
    '                    If objCurMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And objCurMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                        MapItem(objCurMap.SeqNo)
    '                    End If
    '                Case modDeclares.enumMappingType.MAPPING_TYPE_FUN
    '                    Dim sqFun As New clsSQFunction

    '                    sqFun.SQFunctionName = txtCodeEditor.Text
    '                    sqFun.SQFunctionWithInnerText = txtCodeEditor.Text
    '                    objCurMap.SourceType = fieldType
    '                    objCurMap.MappingSource = Nothing
    '                    objCurMap.MappingSource = sqFun

    '                    objCurMap.IsModified = True
    '                    'lvMappings.Items(objCurMap.SeqNo).Text = txtCodeEditor.Text
    '                    'lvMappings.Items(objCurMap.SeqNo).ImageIndex = ImgIdxFromName(objCurMap.MappingSource.Type)
    '                    If objCurMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And objCurMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                        MapItem(objCurMap.SeqNo)
    '                    End If
    '                Case modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                    'Dim fldObj As New clsField

    '                    ObjCurFld.CorrectedFieldName = TruncatedFieldName(txtCodeEditor.Text)
    '                    'lvMappings.Items(objCurMap.SeqNo).Text = txtCodeEditor.Text
    '                    ObjCurFld.Text = txtCodeEditor.Text
    '                    objCurMap.SourceType = fieldType
    '                    objCurMap.MappingSource = ObjCurFld
    '                    objCurMap.IsModified = True


    '                Case modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    'lvMappings.Items(objCurMap.SeqNo).Text = txtCodeEditor.Text
    '                    objCurMap.MappingSource = Nothing
    '                    objCurMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    'lvMappings.Items(objCurMap.SeqNo).ImageIndex = -1
    '                    objCurMap.IsModified = True
    '            End Select
    '        Else
    '            fieldType = GetTypeFromText(txtCodeEditor.Text, modDeclares.enumDirection.DI_TARGET)
    '            If txtCodeEditor.Text <> "" And fieldType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                fieldType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR
    '            End If
    '            If (txtCodeEditor.Text).IndexOf(".") = -1 And fieldType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                fieldType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR
    '            End If
    '            Select Case fieldType
    '                Case modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                    'Dim fldObj As New clsField

    '                    ObjCurFld.CorrectedFieldName = TruncatedFieldName(txtCodeEditor.Text)
    '                    objCurMap.IsModified = True
    '                    'lvMappings.Items(objCurMap.SeqNo).SubItems(1).Text = txtCodeEditor.Text
    '                    ObjCurFld.Text = txtCodeEditor.Text
    '                    objCurMap.MappingTarget = ObjCurFld
    '                    'OnChange(lvMappings, New EventArgs)



    '                Case modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR, enumMappingType.MAPPING_TYPE_VAR
    '                    Dim objVar As New clsVariable

    '                    objVar.Text = txtCodeEditor.Text
    '                    objVar.VariableName = txtCodeEditor.Text
    '                    objVar.CorrectedVariableName = txtCodeEditor.Text
    '                    objCurMap.MappingTarget = objVar

    '                    If objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                        objVar.VariableName = CType(objCurMap.MappingTarget, clsField).FieldName
    '                    ElseIf objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
    '                        objVar.VariableName = CType(objCurMap.MappingTarget, clsVariable).VariableName
    '                    End If

    '                    'lvMappings.Items(objCurMap.SeqNo).SubItems(1).Text = txtCodeEditor.Text
    '                    objCurMap.IsModified = True
    '                    '//fire change event
    '                    'OnChange(lvMappings, New EventArgs)
    '                    If (Not objCurMap.MappingSource Is Nothing) And (Not objCurMap.MappingTarget Is Nothing) Then
    '                        MapItem(objCurMap.SeqNo)
    '                    End If
    '                Case modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    Dim emptyObj As New clsField

    '                    If objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Or _
    '                       objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                        emptyObj.FieldName = CType(objCurMap.MappingTarget, clsField).FieldName
    '                    ElseIf objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Then
    '                        emptyObj.FieldName = CType(objCurMap.MappingTarget, clsVariable).VariableName
    '                    End If
    '                    objCurMap.MappingTarget = emptyObj
    '                    objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    txtCodeEditor.Text = ""
    '                    'lvMappings.Items(objCurMap.SeqNo).SubItems(1).Text = txtCodeEditor.Text
    '                    objCurMap.IsModified = True
    '                    '//fire change event
    '                    'OnChange(lvMappings, New EventArgs)
    '                Case Else
    '                    MsgBox("Destination can only be a Field or Variable!!", MsgBoxStyle.Critical, MsgTitle)
    '                    Return False
    '            End Select
    '        End If

    '        Return True
    '    Catch ex As Exception
    '        LogError(ex, "ctlTask SaveCurrentScript")
    '        Return False
    '    End Try

    'End Function

    'Function SaveMapDesc() As Boolean

    '    If txtMapDesc.Text.Length > 255 Then
    '        MsgBox("The length of your description is too long. Please Limit your descriptions to under 255 Characters", MsgBoxStyle.Information, MsgTitle)
    '        Return False
    '        Exit Function
    '    End If

    '    objCurMap.MappingDesc = Strings.Replace(txtMapDesc.Text, "'", "''")
    '    Return True

    'End Function

#End Region

#Region "Help"

    Private Sub cmdHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Dim helpObj As String = Me.objThis.Type

        Select Case helpObj
            Case "MAN"
                ShowHelp(modHelp.HHId.H_Main_Procedure)
            Case "PRC"
                ShowHelp(modHelp.HHId.H_Mapping_Procedures)
            Case "GEN"
                ShowHelp(modHelp.HHId.H_General_Procedure)
            Case "LOK"
                ShowHelp(modHelp.HHId.H_Lookups)
            Case Else
                ShowHelp(modHelp.HHId.H_Mapping_Procedures)
        End Select

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaskName.KeyDown

        Dim TypeTask As enumTaskType
        TypeTask = objThis.TaskType
        Select Case e.KeyCode
            Case Keys.Escape
                cmdClose_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    Public Sub lvMappings_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvSource.KeyDown  'lvMappings.KeyDown, ToolBar1.KeyDown,

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Source_to_Target_Data_Mapping)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

    Public Sub tvFunctions_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tvFunctions.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Function_Tree)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If
    End Sub

    Public Sub txtCodeEditor_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtCodeEditor.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Editor_Window)
        ElseIf e.KeyCode = Keys.Escape Then
            cmdClose_Click(sender, New EventArgs)
        End If

    End Sub

#End Region

#Region "Load Functions"

    '//Set values from objProject to form controls
    Private Function UpdateFields() As Boolean

        txtTaskName.Text = objThis.TaskName
        txtTaskDesc.Text = objThis.TaskDescription

    End Function

    Function FillMappings() As Boolean
        '//Now load selected mappings
        If IsNewObj = False Then
            objThis.LoadDatastores()
            objThis.LoadMappings(True)
            LoadCodeEditor()
        End If

    End Function

    Function LoadCodeEditor() As Boolean

        Try
            If objThis.TaskType = enumTaskType.TASK_GEN Then
                cbIsMain.Checked = False
            Else
                cbIsMain.Checked = True
            End If

            If objThis.ObjMappings.Count > 0 Then
                objCurMap = CType(objThis.ObjMappings(0), clsMapping)
                objCurFunct = CType(objCurMap.MappingSource, clsSQFunction)
                txtCodeEditor.Text = objCurFunct.SQFunctionWithInnerText
                objCurMap.IsMapped = "1"
                'If objCurMap.IsMapped = "1" Then
                '    cbIsMain.Checked = True
                'End If
            Else
                objCurFunct.SQFunctionName = objThis.TaskName
                objCurFunct.SQFunctionWithInnerText = ""
                txtCodeEditor.Text = objCurFunct.SQFunctionWithInnerText
                objCurMap.IsMapped = "0"
                objCurMap.MappingSource = objCurFunct
                objCurMap.IsModified = True
                objCurMap.SourceType = enumMappingType.MAPPING_TYPE_FUN
                objCurMap.TargetType = enumMappingType.MAPPING_TYPE_NONE
                objThis.ObjMappings.Add(objCurMap)
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "ctlMain LoadCodeEditor")
            Return False
        End Try
    End Function

    Function FillSrcTgtTop(ByVal cNode As TreeNode) As Boolean

        Dim nd As TreeNode
        Dim ndSource As TreeNode
        'Dim ndTarget As TreeNode
        Dim LastSrcNode As String = ""
        'Dim LastTgtNode As String = ""
        Dim tempNode As TreeNode
        'Dim srcChldNode As TreeNode
        'Dim tgtChldNode As TreeNode

        Try
            tvSource.Nodes.Clear()
            'tvTarget.Nodes.Clear()

            tvSource.ImageList = imgListSmall
            'tvTarget.ImageList = imgListSmall

            tvSource.ImageIndex = 32
            'tvTarget.ImageIndex = 32
            tvSource.SelectedImageIndex = 32
            'tvTarget.SelectedImageIndex = 32

            tvSource.CheckBoxes = False
            'tvTarget.CheckBoxes = False

            If objThis.Engine Is Nothing Then
                objThis.LoadDatastores()
            End If

            For Each nd In cNode.Nodes
                If CType(nd.Tag, INode).Type = NODE_FO_DATASTORE Then
                    'For Each ndSrc As TreeNode In nd.Nodes
                    '    If CType(ndSrc.Tag, INode).Type = NODE_SOURCEDATASTORE Then
                    '        ndSource = ndSrc.Clone  '/clone the datastore node
                    '        ndSource.Nodes.Clear()  '/now clear all the clone's child nodes
                    '        tvSource.Nodes.Add(ndSource) '/add it to the tree
                    '    End If
                    'Next

                    Dim DSinode As INode
                    For Each ndDSFolder As TreeNode In nd.Nodes
                        DSinode = ndDSFolder.Tag
                        If Strings.Left(DSinode.Type, 5) = "FO_DS" Then

                            'build source tree from Main tree
                            tempNode = ndDSFolder.Clone
                            tempNode.Nodes.Clear()
                            For Each ndDS As TreeNode In ndDSFolder.Nodes
                                'srcChldNode = ndDS.Clone
                                'srcChldNode.Nodes.Clear()
                                If objThis.ObjSources.Contains(CType(ndDS.Tag, clsDatastore)) = True Then
                                    ndSource = ndDS.Clone
                                    ndSource.Nodes.Clear()
                                    tvSource.Nodes.Add(ndSource)
                                End If
                            Next
                            'tvSource.Nodes.Add(tempNode)
                            'tvSource.ExpandAll()

                            'build Target tree from Main Tree
                            'tempNode = ndDSFolder.Clone
                            'tempNode.Nodes.Clear()
                            'For Each ndDS As TreeNode In ndDSFolder.Nodes
                            '    'tgtChldNode = ndDS.Clone
                            '    'tgtChldNode.Nodes.Clear()
                            '    If objThis.ObjTargets.Contains(CType(ndDS.Tag, clsDatastore)) = True Then
                            '        ndTarget = ndDS.Clone
                            '        ndTarget.Nodes.Clear()
                            '        tvTarget.Nodes.Add(ndTarget)
                            '    End If

                            'Next
                            'tvTarget.Nodes.Add(tempNode)
                            'tvTarget.ExpandAll()
                        End If
                    Next
                ElseIf CType(nd.Tag, INode).Type = NODE_FO_VARIABLE Then
                    'Add Variables to source and Target trees
                    Dim NewNode As TreeNode
                    For Each ndVar As TreeNode In nd.Nodes
                        If Not (objThis.Text = ndVar.Text) Then
                            NewNode = AddTreeNode(tvSource, NODE_VARIABLE, ndVar.Tag)
                            'NewNode = AddNode(tvTarget.Nodes, NODE_VARIABLE, ndVar.Tag)
                        End If
                    Next
                End If

            Next

            'tvSource.Show()
            'tvTarget.Show()


            '//Now load fields for selected source/target items and check if object is already created and user is editing it on seperate dialogbox (not on user control)
            'If IsNewObj = False Then
            '    If objThis.Engine IsNot Nothing Then
            '        objThis.LoadDatastores()
            '        CheckSelectedDatastores()
            '        RemoveUnSelectedDatastores()
            '    End If

            '    '/now add all the selections and fields to each datastore node
            'LoadCheckedItems(tvSource)
            'LoadCheckedItems(tvTarget)
            'End If
            For Each srcDS As TreeNode In tvSource.Nodes
                LoadItemSubTree(srcDS)
            Next
            'For Each tgtDS As TreeNode In tvTarget.Nodes
            '    LoadItemSubTree(tgtDS)
            'Next


            'tvSource.ExpandAll()
            'tvTarget.ExpandAll()

            Return True

        Catch ex As Exception
            LogError(ex, "ctlTask FillSourceTarget")
            Return False
        End Try

    End Function

    Function FillSourceTarget(ByVal cNode As TreeNode) As Boolean

        Dim nd As TreeNode
        Dim ndSource As TreeNode
        'Dim ndTarget As TreeNode
        Dim LastSrcNode As String = ""
        Dim LastTgtNode As String = ""
        'Dim tempNode As TreeNode
        'Dim srcChldNode As TreeNode
        'Dim tgtChldNode As TreeNode

        Try
            tvSource.Nodes.Clear()
            'tvTarget.Nodes.Clear()

            tvSource.ImageList = imgListSmall
            'tvTarget.ImageList = imgListSmall

            tvSource.ImageIndex = 32
            'tvTarget.ImageIndex = 32
            tvSource.SelectedImageIndex = 32
            'tvTarget.SelectedImageIndex = 32

            tvSource.CheckBoxes = False
            'tvTarget.CheckBoxes = False


            objThis.LoadDatastores()


            For Each nd In cNode.Nodes
                If CType(nd.Tag, INode).Type = NODE_FO_SOURCEDATASTORE Then
                    '//load sources and their field in source treeview
                    For Each ndSrc As TreeNode In nd.Nodes
                        If CType(ndSrc.Tag, INode).Type = NODE_SOURCEDATASTORE Then
                            CType(ndSrc.Tag, INode).LoadItems()
                            ndSource = ndSrc.Clone  '/clone the datastore node
                            ndSource.Nodes.Clear()  '/now clear all the clone's child nodes
                            tvSource.Nodes.Add(ndSource) '/add it to the tree
                        End If
                    Next
                    'ElseIf CType(nd.Tag, INode).Type = NODE_FO_TARGETDATASTORE Then
                    '    '//load targets and their field in target treeview
                    '    For Each ndTgt As TreeNode In nd.Nodes
                    '        If CType(ndTgt.Tag, INode).Type = NODE_TARGETDATASTORE Then
                    '            CType(ndTgt.Tag, INode).LoadItems()
                    '            ndTarget = ndTgt.Clone  '/clone the datastore node
                    '            ndTarget.Nodes.Clear()  '/now clear all the clone's child nodes
                    '            tvTarget.Nodes.Add(ndTarget)
                    '        End If
                    '    Next
                ElseIf CType(nd.Tag, INode).Type = NODE_FO_VARIABLE Then
                    'Add Variables to source and Target trees
                    Dim NewNode As TreeNode
                    For Each ndVar As TreeNode In nd.Nodes
                        If Not (objThis.Text = ndVar.Text) Then
                            NewNode = AddTreeNode(tvSource, NODE_VARIABLE, ndVar.Tag)
                            'NewNode = AddNode(tvTarget.Nodes, NODE_VARIABLE, ndVar.Tag)
                        End If
                    Next
                    'ElseIf CType(nd.Tag, INode).Type = NODE_FO_DATASTORE Then
                    '    'For Each ndSrc As TreeNode In nd.Nodes
                    '    '    If CType(ndSrc.Tag, INode).Type = NODE_SOURCEDATASTORE Then
                    '    '        ndSource = ndSrc.Clone  '/clone the datastore node
                    '    '        ndSource.Nodes.Clear()  '/now clear all the clone's child nodes
                    '    '        tvSource.Nodes.Add(ndSource) '/add it to the tree
                    '    '    End If
                    '    'Next

                    '    Dim DSinode As INode
                    '    For Each ndDSFolder As TreeNode In nd.Nodes
                    '        DSinode = ndDSFolder.Tag
                    '        If Strings.Left(DSinode.Type, 5) = "FO_DS" Then

                    '            'build source tree from Main tree
                    '            tempNode = ndDSFolder.Clone
                    '            tempNode.Nodes.Clear()
                    '            For Each ndDS As TreeNode In ndDSFolder.Nodes
                    '                srcChldNode = ndDS.Clone
                    '                'srcChldNode.Nodes.Clear()
                    '                tvSource.Nodes.Add(srcChldNode)
                    '            Next
                    '            'tvSource.Nodes.Add(tempNode)
                    '            tvSource.ExpandAll()

                    '            'build Target tree from Main Tree
                    '            tempNode = ndDSFolder.Clone
                    '            tempNode.Nodes.Clear()
                    '            For Each ndDS As TreeNode In ndDSFolder.Nodes
                    '                tgtChldNode = ndDS.Clone
                    '                'tgtChldNode.Nodes.Clear()
                    '                tvTarget.Nodes.Add(tgtChldNode)
                    '            Next
                    '            'tvTarget.Nodes.Add(tempNode)
                    '            tvTarget.ExpandAll()
                    '        End If
                    '    Next
                    'End If
                    'ElseIf CType(nd.Tag, INode).Type = NODE_FO_JOIN Then
                    '    '//load joins and their field in source treeview coz joins can be only source
                    '    Dim ndJon As TreeNode
                    '    Dim newNode As TreeNode
                    '    Dim objJoin As clsTask
                    '    Dim objMap As clsMapping
                    '    Dim i As Integer
                    '    For Each ndJon In nd.Nodes
                    '        'tvSource.Nodes.Add(ndJon.Clone)
                    '        If Not (objThis.Text = ndJon.Text) Then
                    '            newNode = AddNode(tvSource.Nodes, NODE_JOIN, ndJon.Tag)
                    '            objJoin = ndJon.Tag
                    '            objJoin.LoadMappings()
                    '            For i = 0 To objJoin.ObjMappings.Count - 1
                    '                objMap = objJoin.ObjMappings(i)
                    '                If (objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD) Then
                    '                    '//TODO :
                    '                    If Not objMap.MappingSource Is Nothing Then
                    '                        AddNode(newNode.Nodes, NODE_STRUCT_FLD, objMap.MappingSource, , CType(objMap.MappingSource, clsField).FieldName)
                    '                    End If
                    '                End If
                    '            Next
                    '        End If
                    '    Next

                    'ElseIf CType(nd.Tag, INode).Type = NODE_FO_LOOKUP Then
                    '    '//load lookups and their field in source treeview coz lookups can be only source
                    '    Dim ndLkp As TreeNode
                    '    Dim newNode As TreeNode
                    '    Dim objLook As clsTask
                    '    Dim objMap As clsMapping
                    '    Dim i As Integer
                    '    For Each ndLkp In nd.Nodes
                    '        'tvSource.Nodes.Add(ndJon.Clone)
                    '        If Not (objThis.Text = ndLkp.Text) Then
                    '            newNode = AddNode(tvSource.Nodes, NODE_LOOKUP, ndLkp.Tag)
                    '            objLook = ndLkp.Tag
                    '            objLook.LoadMappings()
                    '            For i = 0 To objLook.ObjMappings.Count - 1
                    '                objMap = objLook.ObjMappings(i)
                    '                If (objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD) Then
                    '                    If Not objMap.MappingSource Is Nothing Then
                    '                        AddNode(newNode.Nodes, NODE_STRUCT_FLD, objMap.MappingSource, , CType(objMap.MappingSource, clsField).FieldName)
                    '                    End If
                    '                End If
                    '            Next
                    '        End If
                    '    Next
                End If
            Next

            'tvSource.Show()
            'tvTarget.Show()


            '//Now load fields for selected source/target items and check if object is already created and user is editing it on seperate dialogbox (not on user control)
            If IsNewObj = False Then
                If objThis.Engine IsNot Nothing Then
                    objThis.LoadDatastores()
                    CheckSelectedDatastores()
                    RemoveUnSelectedDatastores()
                End If

                '/now add all the selections and fields to each datastore node
                LoadCheckedItems(tvSource)
                'LoadCheckedItems(tvTarget)
            End If

            'tvSource.ExpandAll()
            'tvTarget.ExpandAll()

            Return True

        Catch ex As Exception
            LogError(ex, "ctlTask FillSourceTarget")
            Return False
        End Try

    End Function

    Function CheckSelectedDatastores() As Boolean

        Dim i As Integer
        Dim cnt1 As Integer = tvSource.GetNodeCount(True)
        'Dim cnt2 As Integer = tvTarget.GetNodeCount(True)
        Dim nd As TreeNode

        If cnt1 <= 0 Then Exit Function ' + cnt2

        tvSource.BeginUpdate()
        'tvTarget.BeginUpdate()

        '//Looop all selected field and check that field in treeview
        For i = 0 To objThis.ObjSources.Count - 1
            nd = SelectFirstMatchingNode(tvSource, "", True, CType(objThis.ObjSources(i), clsDatastore).Key)
        Next

        'For i = 0 To objThis.ObjTargets.Count - 1
        '    nd = SelectFirstMatchingNode(tvTarget, "", True, CType(objThis.ObjTargets(i), clsDatastore).Key)
        'Next

        tvSource.EndUpdate()
        'tvTarget.EndUpdate()

    End Function

    Function RemoveUnSelectedDatastores() As Boolean

        Dim i As Integer
        Dim nd As TreeNode
        Dim bFound As Boolean
        Dim arrNodes As New ArrayList
        Dim cnt1 As Integer = tvSource.GetNodeCount(False)
        'Dim cnt2 As Integer = tvTarget.GetNodeCount(False)

        If cnt1 <= 0 Then Exit Function '+ cnt2

        tvSource.BeginUpdate()
        'tvTarget.BeginUpdate()

        arrNodes.Clear()

        '//Looop all selected field and check that field in treeview
        For Each nd In tvSource.Nodes
            'nd = tvSource.Nodes(i)
            If CType(nd.Tag, INode).Type = NODE_SOURCEDATASTORE And nd.Checked = False Then
                Try
                    bFound = False
                    For i = 0 To objThis.ObjSources.Count - 1
                        If CType(nd.Tag, INode).Key = objThis.ObjSources(i).key Then
                            bFound = True
                            Exit For
                        End If
                    Next

                    If bFound = False Then
                        'nd.Remove()
                        arrNodes.Add(nd)
                        Console.WriteLine(nd.Text & " will be removed")
                    End If
                Catch ex As Exception
                End Try
            End If
        Next

        For i = 0 To arrNodes.Count - 1
            tvSource.Nodes.Remove(arrNodes(i))
        Next

        'arrNodes.Clear()
        'For Each nd In tvTarget.Nodes
        '    'nd = tvTarget.Nodes(i)
        '    If CType(nd.Tag, INode).Type = NODE_TARGETDATASTORE Then
        '        bFound = False
        '        For i = 0 To objThis.ObjTargets.Count - 1
        '            If CType(nd.Tag, INode).Key = objThis.ObjTargets(i).key Then
        '                bFound = True
        '                Exit For
        '            End If
        '        Next
        '        If bFound = False Then
        '            'nd.Remove()
        '            'tvTarget.Nodes.Remove(nd)
        '            arrNodes.Add(nd)
        '            Console.WriteLine(nd.Text & " will be removed")
        '        End If
        '    End If
        'Next

        'For i = 0 To arrNodes.Count - 1
        '    tvTarget.Nodes.Remove(arrNodes(i))
        'Next

        tvSource.EndUpdate()
        'tvTarget.EndUpdate()

    End Function

    Function LoadCheckedItems(ByVal tv As TreeView) As Boolean

        Dim ndDs As TreeNode
        Dim NodeType As String

        For Each ndDs In tv.Nodes
            NodeType = CType(ndDs.Tag, INode).Type
            If ndDs.Checked = True Or objThis.Engine Is Nothing Then 'And (NodeType = NODE_SOURCEDATASTORE Or NodeType = NODE_TARGETDATASTORE)
                '//Now add subtree (field selection + fields)under this
                LoadItemSubTree(ndDs)
            End If
        Next

        Return True

    End Function

    Function LoadItemSubTree(ByVal ndDs As TreeNode) As Boolean

        Dim objDs As clsDatastore
        Dim objDsSel As clsDSSelection
        Dim i As Integer
        Dim ndSel As TreeNode

        If CType(ndDs.Tag, INode).Type = "SDS" Or CType(ndDs.Tag, INode).Type = "TDS" Then
            objDs = ndDs.Tag
            objDs.LoadItems(True)

            If objDs.ObjSelections.Count > 0 Then
                '//Add all selection of this DS
                For i = 0 To objDs.ObjSelections.Count - 1
                    objDsSel = objDs.ObjSelections(i)
                    objDsSel.LoadMe()
                    '// First add a node in the tree for this DSSelection
                    ndSel = AddNode(ndDs, objDsSel.Type, objDsSel, True)
                    '//Add all fields of this DSSelection
                    ndSel.Tag = objDsSel
                    '//load field list for this DSSelection
                    'objDsSel.LoadItems()
                    '// Add the fields to the DSSelection Node based on Field Level
                    AddFieldsToTreeView(objDsSel, ndSel)
                    '//Dont show all fields, its too much mess in one screen
                    'ndSel.Collapse()
                Next
            End If
        End If

        Return True

    End Function

    '//Loads tasks in Listview
    'Function LoadTaskMappings() As Boolean

    '    Dim objMap As clsMapping

    '    lvMappings.Items.Clear()

    '    For Each objMap In objThis.ObjMappings
    '        AddMapping(objMap, objMap.SeqNo)
    '    Next

    '    lvMappings.SelectedItems.Clear()

    '    If lvMappings.Items.Count > 0 Then
    '        lvMappings.Items(0).EnsureVisible()
    '    End If

    'End Function

#End Region

#Region "TreeView and ListView Functions"

    Private Sub tvFunctions_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvFunctions.AfterSelect

        If CType(e.Node.Tag, INode).Type = NODE_FUN Then
            lblFunName.Text = e.Node.Text
            'ToolTip1.SetToolTip(lblFunName, lblFunName.Text)
            lblFunSyntax.Text = CType(e.Node.Tag, clsSQFunction).SQFunctionSyntax
            'ToolTip1.SetToolTip(lblFunSyntax, lblFunSyntax.Text)
            lblFunDesc.Text = CType(e.Node.Tag, clsSQFunction).SQFunctionDescription
            'ToolTip1.SetToolTip(lblFunDesc, lblFunDesc.Text)
        Else
            lblFunName.Text = "<" & e.Node.Text & ">"
            'ToolTip1.SetToolTip(lblFunName, "Category : " & lblFunName.Text)
            lblFunSyntax.Text = ""
            'ToolTip1.SetToolTip(lblFunSyntax, "")
            lblFunDesc.Text = ""
            'ToolTip1.SetToolTip(lblFunDesc, "")
        End If

    End Sub

    Private Sub tvSource_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvSource.AfterSelect

        tvSource.HideSelection = False
        'tvSource.BackColor = Color.Wheat
        'tvTarget.BackColor = BACK_COLOR

        'tvTarget.HideSelection = True
        'tvTarget.SelectedNode = Nothing

    End Sub

    Private Sub tvSource_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvSource.NodeMouseClick

        Try
            If e.Button = Windows.Forms.MouseButtons.Right Then
                MousePos = e.Location
                'ContextMenuStrip1.Show(tvSource.PointToScreen(e.Location), ToolStripDropDownDirection.BelowRight)
                'OpenStructureFileToolStripMenuItem_Click(sender, New EventArgs)
            End If

        Catch ex As Exception
            LogError(ex, "ctlTask tvSrc_Click")
        End Try

    End Sub

    'Private Sub tvTarget_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs)

    '    tvSource.HideSelection = True
    '    'tvTarget.BackColor = Color.Wheat
    '    'tvSource.BackColor = BACK_COLOR

    '    'tvTarget.HideSelection = False
    '    tvSource.SelectedNode = Nothing

    'End Sub

    'Private Sub tvTarget_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseClickEventArgs)

    '    Try
    '        If e.Button = Windows.Forms.MouseButtons.Right Then
    '            MousePos = e.Location
    '            'ContextMenuStrip1.Show(tvTarget.PointToScreen(e.Location), ToolStripDropDownDirection.BelowRight)
    '            'OpenStructureFileToolStripMenuItem_Click(sender, New EventArgs)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask tvTgt_Click")
    '    End Try

    'End Sub

    Private Sub tvSourceTarget_Hover(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeNodeMouseHoverEventArgs) Handles tvSource.NodeMouseHover

        Try
            Dim Hnode As TreeNode
            Dim obj As INode
            Dim fld As clsField

            Hnode = e.Node
            If Hnode IsNot Nothing Then
                If Hnode.Tag IsNot Nothing Then
                    obj = CType(Hnode.Tag, INode)
                    If obj.Type = NODE_STRUCT_FLD Then
                        fld = CType(obj, clsField)
                        'txtLength.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString
                        'txtDataType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
                        'txtReType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE).ToString
                        'txtExtType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE).ToString
                        'txtInitVal.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL).ToString
                        'txtInvalid.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID).ToString
                        'txtCanNull.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL).ToString
                    End If
                End If
            End If

        Catch ex As Exception
            LogError(ex, "ctlTask tvsourceHover")
        End Try

    End Sub

    'Private Sub lvMappings_Hover(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ListViewItemMouseHoverEventArgs) Handles lvMappings.ItemMouseHover

    '    Try
    '        Dim targetPoint As Point = Control.MousePosition
    '        Dim row, col As Integer
    '        Dim lvItm As ListViewItem
    '        Dim map As clsMapping
    '        Dim fld As clsField = Nothing

    '        targetPoint = lvMappings.PointToClient(targetPoint)
    '        lvItm = lvMappings.GetItemAt(targetPoint.X, targetPoint.Y)
    '        map = CType(lvItm.Tag, clsMapping)

    '        If GetListSubItemFromPoint(lvMappings, targetPoint.X, targetPoint.Y, row, col) = True Then
    '            Select Case col
    '                Case 0
    '                    If map.MappingSource IsNot Nothing Then
    '                        lvItm.ToolTipText = CType(map.MappingSource, INode).Text
    '                        If CType(map.MappingSource, INode).Type = NODE_STRUCT_FLD Or _
    '                        CType(map.MappingSource, INode).Type = MAPPING_TYPE_FIELD Then
    '                            fld = CType(map.MappingSource, clsField)
    '                        End If
    '                    End If
    '                Case 1
    '                    If map.MappingTarget IsNot Nothing Then
    '                        lvItm.ToolTipText = CType(map.MappingTarget, INode).Text
    '                        If CType(map.MappingTarget, INode).Type = NODE_STRUCT_FLD Or _
    '                        CType(map.MappingTarget, INode).Type = MAPPING_TYPE_FIELD Then
    '                            fld = CType(map.MappingTarget, clsField)
    '                        End If
    '                    End If
    '            End Select
    '            If fld IsNot Nothing Then
    '                txtLength.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH).ToString
    '                txtDataType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE).ToString
    '                txtReType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE).ToString
    '                txtExtType.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE).ToString
    '                txtInitVal.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL).ToString
    '                txtInvalid.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID).ToString
    '                txtCanNull.Text = fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL).ToString
    '            End If
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask tvsourceHover")
    '    End Try

    'End Sub

#End Region

#Region "Drag and Drop"

    Private Sub tvDrag_ItemDrag(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles tvFunctions.ItemDrag, tvSource.ItemDrag

        DoDragDrop(e.Item, DragDropEffects.Copy)

    End Sub

    'Private Sub lvMappings_ItemDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemDragEventArgs) Handles lvMappings.ItemDrag

    '    DoDragDrop(e.Item, DragDropEffects.Move)

    'End Sub

    'Private Sub lvMappings_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvMappings.DragEnter

    '    'See if there is a TreeNode being dragged
    '    If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
    '        'TreeNode found allow copy effect
    '        e.Effect = e.AllowedEffect
    '    ElseIf e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", True) Then
    '        e.Effect = DragDropEffects.Move
    '    End If

    'End Sub

    'Private Sub lvMappings_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvMappings.DragOver

    '    lvMappings.MultiSelect = False

    '    If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
    '        '//Drag drop treeview node handle here
    '        OnDragOverFromTreeview(sender, e)
    '    ElseIf e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", True) Then
    '        '//Drag drop listview items handle here 
    '        e.Effect = DragDropEffects.Move
    '        OnDragOverFromListView(sender, e)
    '    End If

    'End Sub

    'Sub OnDragOverFromListView(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)

    '    Try
    '        Dim targetPoint As Point = lvMappings.PointToClient(New Point(e.X, e.Y))
    '        Dim draggedItem As ListViewItem = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
    '        Dim lvItm As ListViewItem

    '        lvItm = lvMappings.GetItemAt(targetPoint.X, targetPoint.Y)

    '        '//just highlight the selection
    '        If Not (lvItm Is Nothing) Then
    '            lvItm.Focused = True
    '        Else
    '            Exit Sub
    '        End If

    '        Select Case draggedItem.ListView.Name
    '            Case "lvMappings"
    '                '//allowed only from above treeviews
    '            Case Else
    '                '//if dragged node came from any other tree except above treevies then dont accept
    '                e.Effect = DragDropEffects.None
    '                Exit Sub
    '        End Select

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask OnDragOverFromListView")
    '    End Try

    'End Sub

    'Sub OnDragOverFromTreeview(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs)

    '    Try
    '        'we only drop if drop on source or target column
    '        Dim targetPoint As Point = lvMappings.PointToClient(New Point(e.X, e.Y))
    '        Dim row, col As Integer
    '        Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
    '        Dim lvItm As ListViewItem

    '        lvItm = lvMappings.GetItemAt(targetPoint.X, targetPoint.Y)
    '        '//just highlight the selection
    '        If Not (lvItm Is Nothing) Then
    '            lvItm.Focused = True
    '            lvItm.Selected = True
    '        End If

    '        Select Case draggedNode.TreeView.Name
    '            Case "tvTarget", "tvSource", "tvFunctions", "tvExplorer"
    '                '//allowed only from above treeviews
    '            Case Else
    '                '//if dragged node came from any other tree except above treevies then dont accept
    '                e.Effect = DragDropEffects.None
    '                Exit Sub
    '        End Select

    '        If GetListSubItemFromPoint(lvMappings, targetPoint.X, targetPoint.Y, row, col) = True Then
    '            Select Case col
    '                Case 0 '//dragging on source

    '                    '//dont allow target to be dropped on source
    '                    If draggedNode.TreeView.Name = "tvTarget" Then
    '                        e.Effect = DragDropEffects.None
    '                        Exit Sub
    '                    End If

    '                    '//Source can take Field, FieldSelection, Variables, functions, 
    '                    '//lookup, Join, variables
    '                    '//only allow the following type of node to be drag/drop on source column
    '                    If CType(draggedNode.Tag, INode).Type = NODE_STRUCT_SEL Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_STRUCT_FLD Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_SOURCEDSSEL Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_JOIN Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_LOOKUP Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_FUN Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_TEMPLATE Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_VARIABLE Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_LOOKUP Then

    '                        e.Effect = DragDropEffects.Copy
    '                    Else
    '                        e.Effect = DragDropEffects.None
    '                    End If

    '                Case 1 '//dragging on target
    '                    '//dont allow source to be dropped on target
    '                    If draggedNode.TreeView.Name = "tvSource" Then
    '                        e.Effect = DragDropEffects.None
    '                        Exit Sub
    '                    End If

    '                    '//Target can take Field, FieldSelection and Variables 
    '                    '//only allow the following type of node to be drag/drop on source column
    '                    If CType(draggedNode.Tag, INode).Type = NODE_STRUCT_SEL Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_STRUCT_FLD Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_TARGETDSSEL Or _
    '                        CType(draggedNode.Tag, INode).Type = NODE_VARIABLE Then

    '                        e.Effect = DragDropEffects.Copy
    '                    Else
    '                        e.Effect = DragDropEffects.None
    '                    End If
    '            End Select
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask OnDragOverFromTreeView")
    '    End Try

    'End Sub

    '    Private Sub lvMappings_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles lvMappings.DragDrop

    '        ' Retrieve the client coordinates of the drop location.
    '        Dim targetPoint As Point = lvMappings.PointToClient(New Point(e.X, e.Y))
    '        Dim lvItm As ListViewItem
    '        'Retrieve the node that was dragged.
    '        Dim draggedNode As TreeNode = Nothing
    '        Dim draggedListItm As ListViewItem

    '        Dim DraggedObj As INode
    '        Dim ObjDragOver As INode
    '        Dim DraggedMap As clsMapping
    '        Dim DestMap As clsMapping

    '        Try
    '            lvMappings.BeginUpdate()

    '            lvItm = lvMappings.GetItemAt(targetPoint.X, targetPoint.Y)

    '            If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
    '                '//Drag drop treeview node handle here
    '                draggedNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
    '            ElseIf e.Data.GetDataPresent("System.Windows.Forms.ListViewItem", True) Then
    '                If Not (lvItm Is Nothing) Then
    '                    draggedListItm = CType(e.Data.GetData(GetType(ListViewItem)), ListViewItem)
    '                    '/// If mapping item is target and dropped on item with no target, then map items together
    '                    '/// and vise-versa  OR insert mapping item in new position.
    '                    DraggedObj = CType(draggedListItm.Tag, INode)
    '                    ObjDragOver = CType(lvItm.Tag, INode)
    '                    If DraggedObj.Type = NODE_MAPPING And ObjDragOver.Type = NODE_MAPPING Then
    '                        DraggedMap = CType(DraggedObj, clsMapping)
    '                        DestMap = CType(ObjDragOver, clsMapping)
    '                        If DraggedMap.SourceType = enumMappingType.MAPPING_TYPE_NONE And _
    '                        DestMap.TargetType = enumMappingType.MAPPING_TYPE_NONE Then
    '                            '/// Add Dragged Target Field to Destination Mapping object
    '                            '// Then delete the mapping that was dragged
    '                            If DestMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD And _
    '                            DraggedMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                                '/// Update the mapping Object  
    '                                DestMap.MappingTarget = CType(DraggedMap.MappingTarget, clsField)
    '                                DestMap.IsMapped = "1"
    '                                DestMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD
    '                                DestMap.TargetParent = DraggedMap.TargetParent
    '                                DestMap.TargetDataStore = DraggedMap.TargetDataStore
    '                                '/// Update the ListView
    '                                lvItm.SubItems(1) = draggedListItm.SubItems(1)
    '                                lvMappings.Items.RemoveAt(draggedListItm.Index)
    '                                ResetMappingSeqNo()
    '                                '//Fire change event
    '                                OnChange(Me, New EventArgs)
    '                                Exit Try
    '                            Else
    '                                GoTo fallThru
    '                            End If
    '                        ElseIf DraggedMap.TargetType = enumMappingType.MAPPING_TYPE_NONE And _
    '                        DestMap.SourceType = enumMappingType.MAPPING_TYPE_NONE Then
    '                            '/// Add Dragged Source Field to Destination Mapping Object
    '                            '// Then delete the mapping that was dragged
    '                            If DestMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD And _
    '                            DraggedMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                                '/// Update the Mappings
    '                                DestMap.MappingSource = CType(DraggedMap.MappingSource, clsField)
    '                                DestMap.IsMapped = "1"
    '                                DestMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD
    '                                DestMap.SourceParent = DraggedMap.SourceParent
    '                                DestMap.SourceDataStore = DraggedMap.SourceDataStore
    '                                '/// Update the Listview
    '                                lvItm.SubItems(0) = draggedListItm.SubItems(0)
    '                                lvMappings.Items.RemoveAt(draggedListItm.Index)
    '                                ResetMappingSeqNo()
    '                                '//Fire change event
    '                                OnChange(Me, New EventArgs)
    '                                Exit Try
    '                            Else
    '                                GoTo fallThru
    '                            End If
    '                        Else
    '                            GoTo fallThru
    '                        End If
    '                    Else
    'fallThru:               lvMappings.MultiSelect = True '// now again turnmulti select on
    '                        '//Note: SetMappingPosition function takes new position of first item in selection. 
    '                        '//If user dragged any item other than first in the selection then calculate new position 
    '                        '//of first item based on move offset of entire selection
    '                        Dim moveOffset As Integer
    '                        moveOffset = lvItm.Index - draggedListItm.Index
    '                        SetMappingPosition(lvMappings.SelectedItems(0).Index + moveOffset)
    '                        Exit Try
    '                    End If
    '                End If
    '            End If

    '            'we only drop if drop on source or target column
    '            Dim row, col As Integer

    '            If GetListSubItemFromPoint(lvMappings, targetPoint.X, targetPoint.Y, row, col) = True Then
    '                Select Case col
    '                    Case 0 '//dropped on source
    '                        DoDropOperation(lvItm, draggedNode, modDeclares.enumDirection.DI_SOURCE, draggedNode.Tag.Type)
    '                    Case 1 '//dropped on target
    '                        DoDropOperation(lvItm, draggedNode, modDeclares.enumDirection.DI_TARGET, draggedNode.Tag.Type)
    '                End Select
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "ctlTask lvMappings.DragDrop")
    '        Finally
    '            lvMappings.EndUpdate()
    '        End Try

    '    End Sub

    '//lvItm : is target item, if lvItm is nothing then we will add dragged item at the end of all list items
    '//draggedNode : is node dragged from tvSource,tvTarget or tvFunction treeview control
    '//              we perform different operation depending on type of node
    'Function DoDropOperation(ByRef lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal nType As String = NODE_STRUCT_FLD) As Boolean

    '    Select Case nType
    '        Case NODE_FUN, NODE_TEMPLATE '//if SQFunction is dropped on listview
    '            OnSQFunctionDrop(lvItm, draggedNode, dirType)
    '        Case NODE_STRUCT_SEL, NODE_SOURCEDSSEL, NODE_TARGETDSSEL
    '            '//if selection is dropped on listview
    '            OnFldSelectionDrop(lvItm, draggedNode, dirType)
    '        Case NODE_STRUCT_FLD '//if field is dropped on listview
    '            OnFieldDrop(lvItm, draggedNode, dirType)
    '        Case NODE_JOIN '//if join is dropped on listview
    '            OnJoinDrop(lvItm, draggedNode, dirType)
    '        Case NODE_LOOKUP '//if lookup is dropped on listview
    '            OnLookupDrop(lvItm, draggedNode, dirType)
    '        Case NODE_VARIABLE '//if variable is dropped on listview
    '            OnVariableDrop(lvItm, draggedNode, dirType)
    '    End Select

    'End Function

    '//call if SQData function is dropped on listview
    'Function OnSQFunctionDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE) As Boolean

    '    If CType(draggedNode.Tag, clsSQFunction).IsTemplate = True Then
    '        UpdateTemplateScript(draggedNode.Tag)
    '    End If

    '    '//Fix on 8/12/05 by npatel : removed draggedNode.tag and replaced with draggedNode
    '    ManualMapping(lvItm, draggedNode, dirType)
    '    AddToRecentlyUsedFunctionList(draggedNode)

    'End Function

    ''//call if variable is dropped on listview
    'Function OnVariableDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE) As Boolean

    '    ManualMapping(lvItm, draggedNode, dirType)

    'End Function

    '//call if join is dropped on listview
    'Function OnJoinDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE) As Boolean

    '    ManualMapping(lvItm, draggedNode, dirType)

    'End Function

    ''//call if lookup is dropped on listview
    'Function OnLookupDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE) As Boolean

    '    ManualMapping(lvItm, draggedNode, dirType)

    'End Function

    '//call if selection is dropped on listview
    'Function OnFldSelectionDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal pass As Integer = 1) As Boolean

    '    Dim objSel As clsDSSelection
    '    Dim objFld As clsField
    '    Dim DataStoreName As String
    '    Dim duplicatecnt As Integer = 0

    '    Try
    '        DataStoreName = GetParentDSForThisNode(draggedNode).Text
    '        objSel = draggedNode.Tag

    '        For Each objFld In objSel.DSSelectionFields
    '            If Not (objFld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" And cbGroupItems.Checked = False) Then
    '                If IsDuplicateItem(objFld, dirType, False) = False Then
    '                    '//TODO Add field and create new mapping
    '                    If objFld.FieldName <> "" Then
    '                        objFld.Text = objFld.ParentStructureName & "." & objFld.FieldName 'GetParentDSForThisNode(draggedNode).Text & "." & _
    '                    ElseIf objFld.OrgName <> "" Then
    '                        objFld.Text = objFld.ParentStructureName & "." & objFld.CorrectedFieldName 'GetParentDSForThisNode(draggedNode).Text & "." & _
    '                    Else
    '                        objFld.Text = objFld.ParentStructureName & "." & objFld.OrgName 'GetParentDSForThisNode(draggedNode).Text & "." & _
    '                    End If
    '                    AutoFieldMapping(objFld, dirType, DataStoreName, pass)
    '                Else
    '                    '//skipping field coz its already in the list
    '                    If pass < 1 Then
    '                        duplicatecnt = duplicatecnt + 1
    '                    End If
    '                End If
    '            End If
    '        Next
    '        pass = pass + 1
    '        If pass < 6 Then
    '            Call OnFldSelectionDrop(lvItm, draggedNode, dirType, pass)
    '        End If

    '        If duplicatecnt = 1 Then
    '            MsgBox("[" & duplicatecnt & "] duplicate item was skipped", MsgBoxStyle.Exclamation, MsgTitle)
    '        ElseIf duplicatecnt > 1 Then
    '            MsgBox("[" & duplicatecnt & "] duplicate items were skipped", MsgBoxStyle.Exclamation, MsgTitle)
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex)
    '        Return False
    '    End Try

    'End Function

    '//call if single field is dropped on listview
    'Function OnFieldDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE) As Boolean

    '    Try
    '        Dim objfld As clsField = draggedNode.Tag
    '        Dim GrpItmFlag As Boolean = False

    '        If (objfld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" And objfld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE) = "") Then
    '            GroupItemDrop(lvItm, draggedNode, dirType)
    '        Else
    '            If dirType = modDeclares.enumDirection.DI_SOURCE Then
    '                ManualMapping(lvItm, draggedNode, dirType, True)
    '            Else
    '                If IsDuplicateItem(CType(draggedNode.Tag, clsField), dirType, False) = False Then
    '                    ManualMapping(lvItm, draggedNode, dirType, True)
    '                Else
    '                    MsgBox("Duplicate field was found", MsgBoxStyle.Exclamation, MsgTitle)
    '                End If
    '            End If
    '        End If

    '        'If objfld.Parent.Type = NODE_SOURCEDSSEL Or objfld.Parent.Type = NODE_TARGETDSSEL Then
    '        '    CType(objfld.Parent, clsDSSelection).IsMapped = True
    '        'End If

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex)
    '        Return False
    '    End Try

    'End Function

    '// call if groupitem field is dropped on listview
    'Function GroupItemDrop(ByVal lvItm As ListViewItem, ByVal draggedNode As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal pass As Integer = 1) As Boolean

    '    Dim FldNodeCol As TreeNodeCollection = draggedNode.Nodes
    '    Dim curNode As TreeNode
    '    Dim objfld As clsField = draggedNode.Tag
    '    Dim duplicatecnt As Integer = 0
    '    Dim DataStoreName As String

    '    Try
    '        DataStoreName = GetParentDSForThisNode(draggedNode).Text
    '        If DataStoreName = "" Then
    '            DataStoreName = GetParentDSForField(draggedNode)
    '        End If

    '        '// First do Group Item Field ... Map if Group Items box checked
    '        If Not (objfld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" And cbGroupItems.Checked = False) Then
    '            If IsDuplicateItem(objfld, dirType, False) = False Then
    '                '//TODO Add field and create new mapping
    '                If objfld.CorrectedFieldName <> "" Then
    '                    objfld.Text = objfld.ParentStructureName & "." & objfld.CorrectedFieldName ' GetParentDSForThisNode(draggedNode).Text & "." & 
    '                ElseIf objfld.OrgName <> "" Then
    '                    objfld.Text = objfld.ParentStructureName & "." & objfld.OrgName 'GetParentDSForThisNode(draggedNode).Text & "." & 
    '                Else
    '                    objfld.Text = objfld.ParentStructureName & "." & objfld.FieldName 'GetParentDSForThisNode(draggedNode).Text & "." & 
    '                End If
    '                AutoFieldMapping(objfld, dirType, DataStoreName, pass)
    '            Else
    '                '//skipping field coz its already in the list
    '                duplicatecnt = duplicatecnt + 1
    '            End If
    '        Else
    '            '/// Now do the Group Item's Child fields
    '            For Each curNode In FldNodeCol
    '                objfld = curNode.Tag
    '                If objfld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE) = "GROUPITEM" Then
    '                    '/// recurse if Child field is a Group Item
    '                    GroupItemDrop(lvItm, curNode, dirType)
    '                Else
    '                    If IsDuplicateItem(objfld, dirType, False) = False Then
    '                        '//TODO Add field and create new mapping
    '                        If objfld.CorrectedFieldName <> "" Then
    '                            objfld.Text = objfld.ParentStructureName & "." & objfld.CorrectedFieldName 'GetParentDSForThisNode(draggedNode).Text & "." & 
    '                        ElseIf objfld.OrgName <> "" Then
    '                            objfld.Text = objfld.ParentStructureName & "." & objfld.OrgName ' GetParentDSForThisNode(draggedNode).Text & "." & 
    '                        Else
    '                            objfld.Text = objfld.ParentStructureName & "." & objfld.FieldName 'GetParentDSForThisNode(draggedNode).Text & "." & 
    '                        End If
    '                        AutoFieldMapping(objfld, dirType, DataStoreName, pass)
    '                    Else
    '                        '//skipping field coz its already in the list
    '                        duplicatecnt = duplicatecnt + 1
    '                    End If
    '                End If
    '            Next

    '        End If

    '        pass = pass + 1
    '        If pass < 6 Then
    '            Call GroupItemDrop(lvItm, draggedNode, dirType, pass)
    '        End If

    '        'If duplicatecnt = 1 Then
    '        '    MsgBox("[" & duplicatecnt & "] duplicate item was skipped", MsgBoxStyle.Exclamation)
    '        'ElseIf duplicatecnt > 1 Then
    '        '    MsgBox("[" & duplicatecnt & "] duplicate items were skipped", MsgBoxStyle.Exclamation)
    '        'End If

    '        Return True
    '    Catch ex As Exception
    '        LogError(ex)
    '        Return False
    '    End Try

    'End Function

    Private Sub txtCodeEditor_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtCodeEditor.DragEnter

        Try
            'See if there is a TreeNode or text being dragged
            If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Or e.Data.GetDataPresent(DataFormats.Text) Then
                'TreeNode found allow copy effect
                e.Effect = DragDropEffects.Copy
            End If

        Catch ex As Exception
            LogError(ex, "ctlMain txtCodeEditor_DragEnter")
        End Try

    End Sub

    Private Sub txtCodeEditor_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtCodeEditor.DragOver

        Try
            Dim pt As Point
            Dim ret As Integer

            pt = New Point(e.X, e.Y)
            pt = txtCodeEditor.PointToClient(pt)
            txtCodeEditor.Focus()
            ret = GetCharFromPos(txtCodeEditor, pt)
            txtCodeEditor.Select(ret, 0)

        Catch ex As Exception
            LogError(ex, "ctlMain txtCodeEditor_DragOver")
        End Try

    End Sub

    Private Sub txtCodeEditor_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles txtCodeEditor.DragDrop

        Dim prevSel As Integer
        Dim txt As String
        Dim obj As INode

        Try
            prevSel = txtCodeEditor.SelectionStart
            If e.Data.GetDataPresent("System.Windows.Forms.TreeNode", True) Then
                Dim draggedNode As TreeNode = CType(e.Data.GetData(GetType(TreeNode)), TreeNode)
                '//Only accept nodes with INOde interface
                If Not (draggedNode.Tag.GetType.GetInterface("INode") Is Nothing) Then
                    '//handle function differently. We will show expanded text with para place holder
                    If draggedNode.Tag.GetType.Name = GetType(clsSQFunction).Name Then
                        If draggedNode.Text = "Route" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetMainText) 'GetScriptForMainTemplateV3
                        ElseIf draggedNode.Text = "Procedure" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetScriptForProcV3)
                        ElseIf draggedNode.Text = "LOOK" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetScriptForLOOK)
                        ElseIf draggedNode.Text = "CASE" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetScriptForCASE)
                        ElseIf draggedNode.Text = "SetImage" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetSetImageTemplate)
                        ElseIf draggedNode.Text = "CURRENTDATE" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetScriptForCurrentDate)
                        ElseIf draggedNode.Text = "MAPBEFOREkeyChng" Then
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, GetScriptForKeyChange)
                        Else
                            txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, _
                            CType(draggedNode.Tag, clsSQFunction).SQFunctionWithInnerText)
                        End If

                        AddToRecentlyUsedFunctionList(draggedNode)
                    Else
                        obj = draggedNode.Tag
                        If obj.Type = NODE_STRUCT_FLD Then
                            If CType(obj, clsField).CorrectedFieldName <> "" Then
                                txt = GetParentDSForThisNode(draggedNode).Text & "." & _
                                CType(obj, clsField).ParentStructureName & "." & _
                                CType(obj, clsField).CorrectedFieldName
                            ElseIf CType(obj, clsField).OrgName <> "" Then
                                txt = GetParentDSForThisNode(draggedNode).Text & "." & _
                                CType(obj, clsField).ParentStructureName & "." & _
                                CType(obj, clsField).OrgName
                            Else
                                txt = GetParentDSForThisNode(draggedNode).Text & "." & _
                                CType(obj, clsField).ParentStructureName & "." & _
                                CType(obj, clsField).FieldName
                            End If
                        Else
                            txt = obj.Text
                        End If
                        txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, txt)
                    End If

                    txtCodeEditor.SelectionStart = prevSel
                End If
            ElseIf e.Data.GetDataPresent(DataFormats.Text) Then
                txtCodeEditor.Text = txtCodeEditor.Text.Insert(prevSel, e.Data.GetData(DataFormats.Text))
                txtCodeEditor.SelectionStart = prevSel
            End If

        Catch ex As Exception
            LogError(ex, "ctlTask DragDrop")
        End Try

    End Sub

#End Region

#Region "Menu and Click Events"

    'Public Sub cbGroupItems_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbGroupItems.CheckedChanged

    '    If IsEventFromCode = True Then Exit Sub

    '    objThis.IsModified = True

    '    cmdSave.Enabled = True

    '    If objThis.Engine IsNot Nothing Then
    '        If cbGroupItems.Checked = True Then
    '            objThis.Engine.MapGroupItems = True
    '        Else
    '            objThis.Engine.MapGroupItems = False
    '        End If
    '    End If

    '    RaiseEvent Modified(Me, objThis)

    'End Sub

    'Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    'HideScriptEditor()
    '    'HideDescEditor(True)

    'End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        pnlScript.Visible = False
        'pnlDesc.Visible = False
        'pnlSourceTarget.Visible = True
        'pnlSourceTarget.BringToFront()
        IsCodeEditorOnTop = False
        IsDescEditorOnTop = False

    End Sub

    'Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If Not (tvSource.SelectedNode Is Nothing) Then
    '        If SelectFirstMatchingNode(tvSource, colSkipNodes, txtSearchField.Text) = False Then
    '            MsgBox("No matching node found for entered text", MsgBoxStyle.Critical)
    '        End If
    '    ElseIf Not (tvTarget.SelectedNode Is Nothing) Then
    '        If SelectFirstMatchingNode(tvTarget, colSkipNodes, txtSearchField.Text) = False Then
    '            MsgBox("No matching node found for entered text", MsgBoxStyle.Critical)
    '        End If
    '    Else
    '        MsgBox("Please click on the treeview you want to search")
    '    End If

    'End Sub

    Public Sub cmdSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        objThis.CallFromUsercontrol = True '//8/15/05
        Save()
        objThis.CallFromUsercontrol = False '//8/15/05

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
                    objThis.IsDisturbed = True
                ElseIf ret = MsgBoxResult.Cancel Then
                    Exit Sub
                End If
            End If

            Me.Visible = False
            RaiseEvent Closed(Me, objThis)

        Catch ex As Exception
            LogError(ex, "ctlTask cmdClose_Click")
        End Try

    End Sub

    'Private Sub cmdSaveDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    'HideScriptEditor(True)
    '    HideDescEditor()

    '    'If CType(lvMappings.Items(CurRow).Tag, clsMapping).MappingDesc <> "" Then
    '    '    lvMappings.Items(CurRow).ForeColor = HAS_DESC_COLOR
    '    'Else
    '    '    lvMappings.Items(CurRow).ForeColor = Color.Black
    '    'End If

    'End Sub

    Private Sub cmdCancelEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        pnlScript.Visible = False
        ' pnlDesc.Visible = False
        'pnlSourceTarget.Visible = True
        'pnlSourceTarget.BringToFront()
        IsCodeEditorOnTop = False
        IsDescEditorOnTop = False

    End Sub

    'Private Sub lvMappings_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvMappings.DoubleClick

    '    Dim row As Integer
    '    Dim col As Integer
    '    Dim objMap As clsMapping

    '    Try
    '        row = CurRow
    '        col = CurCol

    '        '//if user double click on function then we show script edit control
    '        If row >= 0 Then
    '            objMap = lvMappings.Items(row).Tag

    '            If Not (objMap Is Nothing) Then
    '                If objMap.HasMissingDependency = True Then
    '                    MsgBox("**** Sorry you can not edit this mapping. *****" & vbCrLf & vbCrLf & objMap.Commment, MsgBoxStyle.Critical, MsgTitle)
    '                    Exit Sub
    '                End If

    '                '// functions can be only entered in source
    '                If col = 0 Then
    '                    curEditType = modDeclares.enumDirection.DI_SOURCE
    '                Else
    '                    curEditType = modDeclares.enumDirection.DI_TARGET
    '                End If

    '                ShowScriptEditor(objMap)
    '            End If '//If Not (objMap Is Nothing) Then ...
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask lvMappings_DblClick")
    '    End Try

    'End Sub

    '    Private Sub lvMappings_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvMappings.MouseDown

    '        Dim row, col As Integer
    '        Dim ItemClicked As Boolean
    '        Dim lvItm As ListViewItem

    '        ItemClicked = GetListSubItemFromPoint(lvMappings, e.X, e.Y, row, col)
    '        If ItemClicked = False Then
    '            For Each lvItm In lvMappings.SelectedItems
    '                lvItm.Selected = False
    '                lvItm.Focused = False
    '            Next
    '        End If
    '        If Not lvMappings.FocusedItem Is Nothing Then
    '            lvMappings.FocusedItem.Selected = False
    '            lvMappings.FocusedItem.Focused = False
    '        End If

    '    End Sub

    '    Private Sub lvMappings_Mouseup(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lvMappings.MouseUp

    '        Dim row As Integer, col As Integer
    '        Dim ItemClicked As Boolean
    '        Dim flagOutside As Boolean
    '        Dim objMap As clsMapping

    '        Try
    '            ItemClicked = GetListSubItemFromPoint(lvMappings, e.X, e.Y, row, col)

    '            CurRow = row
    '            CurCol = col
    '            flagOutside = (row = -1)

    '            If e.Button = Windows.Forms.MouseButtons.Right Then
    '                If ItemClicked = True Then
    '                    mnuDelItem.Enabled = Not flagOutside
    '                    mnuDelSource.Enabled = Not flagOutside
    '                    mnuDelTarget.Enabled = Not flagOutside

    '                    mnuCopyMapping.Enabled = Not flagOutside
    '                    mnuCutMapping.Enabled = Not flagOutside

    '                    mnuPaste.Enabled = IsClipboardAvail
    '                    mnuEdit.Enabled = Not flagOutside
    '                    mnuMapDesc.Enabled = Not flagOutside

    '                    If row >= 0 Then

    '                        objMap = lvMappings.Items(row).Tag
    '                        If Not (objMap Is Nothing) Then

    '                            If lvMappings.Items(row).Selected = False Then
    '                                lvMappings.Items(row).Selected = True
    '                                lvMappings.MultiSelect = False
    '                            End If


    '                            If col = 0 And objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                                mnuDelSource.Enabled = False
    '                            End If
    '                            If col = 1 And objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                                mnuDelTarget.Enabled = False
    '                            End If
    '                        End If

    '                        If lvMappings.SelectedItems.Count > 1 Then
    '                            lvMappings.MultiSelect = True

    '                            '//Allow Del Source/Target/row for muli select if applicable 
    '                            mnuDelItem.Enabled = mnuDelItem.Enabled And True
    '                            mnuDelSource.Enabled = mnuDelSource.Enabled And True
    '                            mnuDelTarget.Enabled = mnuDelTarget.Enabled And True

    '                            '//Do not allow Cut/Copy for muli select
    '                            mnuCutMapping.Enabled = mnuCutMapping.Enabled And False
    '                            mnuCopyMapping.Enabled = mnuCopyMapping.Enabled And False
    '                            mnuPaste.Enabled = mnuPaste.Enabled And False
    '                        Else
    '                            lvMappings.MultiSelect = False
    '                        End If
    '                    End If
    '                Else '//Left button up
    '                    Debug.Write(Now)
    '                End If

    '                mnuPopup.Show(lvMappings, New Point(e.X, e.Y))
    '            Else
    '                lvMappings.MultiSelect = True
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "ctlTask lvMappings_MouseUp")
    '        End Try

    '    End Sub

    '    Private Sub mnuCopyMapping_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCopyMapping.Click

    '        If CurRow >= 0 Then
    '            objClip = GetObjCopy(objThis, lvMappings.Items(CurRow).Tag)
    '            IsClipboardAvail = True
    '        End If

    '    End Sub

    '    Private Sub mnuCopySingle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCopySingle.Click

    '        Try
    '            If CurRow >= 0 Then
    '                Dim objMap As clsMapping
    '                TempMap = New clsMapping

    '                objMap = lvMappings.Items(CurRow).Tag
    '                If CurCol = 0 Then
    '                    '/// build clip object
    '                    TempMap = objMap.Clone(objThis)
    '                    TempMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    TempMap.MappingTarget = Nothing
    '                    TempMap.TargetDataStore = ""
    '                    TempMap.IsMapped = "0"
    '                    '/// add to clipboard
    '                    objClip = TempMap
    '                Else
    '                    '/// build clip object
    '                    TempMap = objMap.Clone(objThis)
    '                    TempMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    TempMap.MappingSource = Nothing
    '                    TempMap.SourceDataStore = ""
    '                    TempMap.IsMapped = "0"
    '                    '/// add to clipboard
    '                    objClip = TempMap
    '                End If

    '                IsClipboardAvail = True
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "ctlTask mnuCopySingle")
    '        End Try

    '    End Sub

    '    Private Sub mnuCutSingle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuCutSingle.Click

    '        Try
    '            If CurRow >= 0 Then
    '                Dim objMap As clsMapping
    '                TempMap = New clsMapping

    '                objMap = lvMappings.Items(CurRow).Tag
    '                If CurCol = 0 Then
    '                    '/// build clip object
    '                    TempMap = objMap.Clone(objThis)
    '                    TempMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    TempMap.MappingTarget = Nothing
    '                    TempMap.TargetDataStore = ""
    '                    TempMap.IsMapped = "0"
    '                    '/// add to clipboard
    '                    objClip = TempMap
    '                    '/// update remaining object
    '                    lvMappings.Items(CurRow).SubItems(0).Text = " "
    '                    objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    objMap.MappingSource = Nothing
    '                    objMap.SourceDataStore = ""
    '                    objMap.IsMapped = "0"
    '                Else
    '                    '/// build clip object
    '                    TempMap = objMap.Clone(objThis)
    '                    TempMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    TempMap.MappingSource = Nothing
    '                    TempMap.SourceDataStore = ""
    '                    TempMap.IsMapped = "0"
    '                    '/// add to clipboard
    '                    objClip = TempMap
    '                    '/// update remaining object
    '                    lvMappings.Items(CurRow).SubItems(1).Text = " "
    '                    objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                    objMap.MappingTarget = Nothing
    '                    objMap.TargetDataStore = ""
    '                    objMap.IsMapped = "0"
    '                End If
    '                lvMappings.Items(CurRow).Tag = objMap
    '                IsClipboardAvail = True
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "ctlTask mnuCutSingle")
    '        End Try

    '    End Sub

    '    Private Sub mnuCutMapping_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuCutMapping.Click

    '        Try
    '            If CurRow >= 0 Then
    '                Dim objMap As clsMapping

    '                objMap = lvMappings.Items(CurRow).Tag

    '                objClip = Nothing
    '                objClip = objMap.Clone(objThis)         'new by TK 3/2/2007
    '                lvMappings.Items.RemoveAt(CurRow)

    '                IsClipboardAvail = True
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "ctlTask mnuCutMapping")
    '        End Try

    '    End Sub

    '    Private Sub mnuPaste_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuPaste.Click

    '        Dim CutCopyMap As clsMapping
    '        Dim DestMap As clsMapping

    '        Try
    '            lvMappings.BeginUpdate()

    '            If IsClipboardAvail = False Then Exit Sub

    '            If CType(objClip, INode).Type = NODE_MAPPING Then
    '                If CType(lvMappings.Items(CurRow).Tag, clsMapping).SourceType = enumMappingType.MAPPING_TYPE_NONE And _
    '                CType(objClip, clsMapping).TargetType = enumMappingType.MAPPING_TYPE_NONE Then
    '                    DestMap = CType(lvMappings.Items(CurRow).Tag, clsMapping)
    '                    CutCopyMap = CType(objClip, clsMapping)
    '                    If DestMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD And _
    '                    CutCopyMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                        '/// Update the mapping Object  
    '                        DestMap.MappingSource = CType(CutCopyMap.MappingSource, clsField)
    '                        DestMap.IsMapped = "1"
    '                        DestMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD
    '                        DestMap.SourceParent = CutCopyMap.SourceParent
    '                        DestMap.SourceDataStore = CutCopyMap.SourceDataStore
    '                        '/// Update the ListView
    '                        lvMappings.Items(CurRow).SubItems(0).Text = CType(CutCopyMap.MappingSource, clsField).Text
    '                        'ResetMappingSeqNo()
    '                        '//Fire change event
    '                        OnChange(Me, New EventArgs)
    '                        Exit Try
    '                    Else
    '                        GoTo fallthru2
    '                    End If
    '                ElseIf CType(lvMappings.Items(CurRow).Tag, clsMapping).TargetType = enumMappingType.MAPPING_TYPE_NONE And _
    '                CType(objClip, clsMapping).SourceType = enumMappingType.MAPPING_TYPE_NONE Then
    '                    DestMap = CType(lvMappings.Items(CurRow).Tag, clsMapping)
    '                    CutCopyMap = CType(objClip, clsMapping)
    '                    If DestMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD And _
    '                    CutCopyMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                        '/// Update the mapping Object  
    '                        DestMap.MappingTarget = CType(CutCopyMap.MappingTarget, clsField)
    '                        DestMap.IsMapped = "1"
    '                        DestMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD
    '                        DestMap.TargetParent = CutCopyMap.TargetParent
    '                        DestMap.TargetDataStore = CutCopyMap.TargetDataStore
    '                        '/// Update the ListView
    '                        lvMappings.Items(CurRow).SubItems(1).Text = CType(CutCopyMap.MappingTarget, clsField).Text
    '                        'ResetMappingSeqNo()
    '                        '//Fire change event
    '                        OnChange(Me, New EventArgs)
    '                        Exit Try
    '                    Else
    '                        GoTo fallthru2
    '                    End If
    '                Else
    'fallthru2:          AddMapping(objClip, CurRow)
    '                    IsClipboardAvail = True
    '                End If
    '            Else
    '                If CurRow >= 0 Then
    '                    Dim objMap As clsMapping
    '                    objMap = lvMappings.Items(CurRow).Tag
    '                    If CurCol = 0 Then
    '                        lvMappings.Items(CurRow).SubItems(0).Text = CType(objClip, INode).Text
    '                        objMap.SourceType = GetSourceTypeFromNodeType(CType(objClip, INode).Type)
    '                        objMap.MappingSource = GetObjCopy(objThis, objClip)

    '                    Else
    '                        lvMappings.Items(CurRow).SubItems(1).Text = CType(objClip, INode).Text
    '                        objMap.TargetType = GetSourceTypeFromNodeType(CType(objClip, INode).Type)
    '                        objMap.MappingTarget = GetObjCopy(objThis, objClip)

    '                    End If

    '                    IsClipboardAvail = True
    '                End If
    '            End If


    '        Catch ex As Exception
    '            LogError(ex, "ctlTask mnuPastClick")
    '        Finally
    '            lvMappings.EndUpdate()
    '        End Try

    '    End Sub

    'Private Sub mnuSelectMappedItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim itm As ListViewItem
    '    Dim map As clsMapping

    '    lvMappings.SelectedItems.Clear()
    '    lvMappings.MultiSelect = True
    '    For Each itm In lvMappings.Items
    '        map = itm.Tag
    '        If map.IsMapped = "1" Or map.IsMapped = "2" Or map.IsMapped = "3" Then
    '            itm.Selected = True
    '        End If
    '    Next

    'End Sub

    'Private Sub mnuSelectUnmappedItems_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim itm As ListViewItem
    '    Dim map As clsMapping

    '    lvMappings.SelectedItems.Clear()
    '    lvMappings.MultiSelect = True

    '    For Each itm In lvMappings.Items
    '        map = itm.Tag
    '        '//according to Kam select all unmapped items where target is blank
    '        If map.IsMapped = "0" And (map.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE) And (map.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE) Then
    '            itm.Selected = True
    '        End If
    '    Next

    'End Sub

    'Private Sub cmdSearchMapping_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim itm As ListViewItem
    '    Dim map As clsMapping
    '    Dim i As Integer

    '    '//check for new search
    '    If lastMappingSearchText <> txtSearchMapping.Text Or lastMappingSearchPos = 0 Then
    '        lvMappings.SelectedItems.Clear()
    '        lastMappingSearchPos = 0
    '    End If

    '    If lastMappingSearchPos > 0 And lastMappingSearchPos < lvMappings.Items.Count - 1 Then
    '        lastMappingSearchPos = lastMappingSearchPos + 1 '// start from next item from last found item
    '    Else
    '        lastMappingSearchPos = 0
    '    End If

    '    lvMappings.MultiSelect = True
    '    lastMappingSearchText = txtSearchMapping.Text

    '    For i = lastMappingSearchPos To lvMappings.Items.Count - 1
    '        itm = lvMappings.Items(i)
    '        map = itm.Tag
    '        If itm.SubItems(0).Text.ToLower.IndexOf(txtSearchMapping.Text.ToLower) >= 0 Then
    '            itm.Selected = True
    '            lvMappings.EnsureVisible(itm.Index)
    '            lastMappingSearchPos = itm.Index
    '            lvMappings.Focus()
    '            Exit For
    '        ElseIf itm.SubItems(1).Text.ToLower.IndexOf(txtSearchMapping.Text.ToLower) >= 0 Then
    '            itm.Selected = True
    '            lvMappings.EnsureVisible(itm.Index)
    '            lastMappingSearchPos = itm.Index
    '            lvMappings.Focus()
    '            Exit For
    '        End If
    '        If i >= lvMappings.Items.Count - 1 Then
    '            If lastMappingSearchPos = 0 Then
    '                MsgBox("No item found matching with your search criteria.", MsgBoxStyle.Information)
    '            Else
    '                MsgBox("You have reached at the end.", MsgBoxStyle.Information)
    '            End If

    '            lastMappingSearchPos = 0
    '            Exit For
    '        End If
    '    Next

    'End Sub

    'Private Sub mnuEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuEdit.Click
    '    lvMappings_DoubleClick(sender, e)
    'End Sub

    'Private Sub mnuMapDesc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMapDesc.Click

    '    Dim row As Integer
    '    Dim col As Integer
    '    Dim objMap As clsMapping

    '    Try
    '        row = CurRow
    '        col = CurCol

    '        If lvMappings.SelectedItems.Count <> 1 Then
    '            MsgBox("Please select a single mapping to edit a description", MsgBoxStyle.Information, MsgTitle)
    '            Exit Sub
    '        End If
    '        '//if user double click on function then we show script edit control
    '        If row >= 0 Then
    '            objMap = lvMappings.Items(row).Tag

    '            If Not (objMap Is Nothing) Then
    '                If objMap.HasMissingDependency = True Then
    '                    MsgBox("**** Sorry you can not edit this mapping. *****" & vbCrLf & vbCrLf & objMap.Commment, MsgBoxStyle.Critical, MsgTitle)
    '                    Exit Sub
    '                End If

    '                '// functions can be only entered in source
    '                If col = 0 Then
    '                    curEditType = modDeclares.enumDirection.DI_SOURCE
    '                Else
    '                    curEditType = modDeclares.enumDirection.DI_TARGET
    '                End If

    '                ShowDescEditor(objMap)
    '                'lvMappings.Items(row).BackColor = HAS_DESC_COLOR
    '            End If '//If Not (objMap Is Nothing) Then ...
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask mnuMapDesc_Click")
    '    End Try

    'End Sub

    'Private Sub mnuMapList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMapList.Click

    '    Try
    '        Dim SrcArr As New ArrayList
    '        Dim TgtArr As New ArrayList

    '        SrcArr.Clear()
    '        TgtArr.Clear()

    '        If lvMappings.SelectedItems.Count > 0 Then
    '            For Each lvItm As ListViewItem In lvMappings.SelectedItems
    '                If lvItm.Tag IsNot Nothing Then
    '                    Dim map As clsMapping = lvItm.Tag
    '                    If map.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                        SrcArr.Add(CType(map.MappingSource, clsField).FieldName)
    '                    End If
    '                    If map.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                        TgtArr.Add(CType(map.MappingTarget, clsField).FieldName)
    '                    End If
    '                End If
    '            Next
    '            Dim frm As New frmMapList
    '            Call frm.NewOrOpen(objThis.Project, SrcArr, TgtArr)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask mnuMapListClick")
    '    End Try

    'End Sub

    'Private Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick

    '    Dim ActionType As Object = e.Button.Tag

    '    Try
    '        Select Case ActionType.ToString
    '            Case "Edit"
    '                mnuEdit.PerformClick()
    '            Case "Cut"
    '                mnuCutMapping.PerformClick()
    '            Case "Copy"
    '                mnuCopyMapping.PerformClick()
    '            Case "Paste"
    '                mnuPaste.PerformClick()
    '            Case "Delete"
    '                DelMapping(sender, e)
    '            Case "DelSrc"
    '                mnuDelSource.PerformClick()
    '            Case "DelTgt"
    '                mnuDelTarget.PerformClick()
    '            Case "InsUp"
    '                mnuInsertMapping.PerformClick()
    '            Case "InsDown"
    '                AddNewMapping(sender, e)
    '            Case "Up"
    '                MoveUp(sender, e)
    '            Case "Down"
    '                MoveDown(sender, e)
    '            Case "Desc"
    '                mnuMapDesc.PerformClick()
    '            Case "Src"
    '                ShowHideSrc()
    '            Case "Tgt"
    '                ShowHideTgt()
    '        End Select

    '    Catch ex As Exception
    '        Log("ToolBar1_ButtonClick=>" & ex.Message)
    '    End Try

    'End Sub

    'Private Sub ShowHideSrc()

    '    scSrc.Panel1Collapsed = Not tbShowSrc.Pushed

    'End Sub

    'Private Sub ShowHideTgt()

    '    scTgt.Panel2Collapsed = Not tbShowTgt.Pushed

    'End Sub

    'Private Sub OnInsertMapping(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuInsertMapping.Click

    '    Dim flagOutside As Boolean
    '    Dim pt As Point = lvMappings.PointToClient(New Point(MousePosition.X, MousePosition.Y))

    '    Try
    '        flagOutside = (CurRow = -1)

    '        If flagOutside = True Then
    '            '//add new item at the end
    '            AddMapping()
    '        Else
    '            '//add new item above the selected item
    '            AddMapping(CurRow)

    '            MarkAsModifedBelowThis(CurRow)

    '            OnChange(Me, New EventArgs)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask mnuInsMapping_Click")
    '    End Try

    'End Sub

    'Private Sub OnDelItem(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelItem.Click

    '    DelMapping(Me, New EventArgs)

    'End Sub

    'Private Sub OnDelSource(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelSource.Click

    '    Dim i As Integer
    '    Dim objMap As clsMapping

    '    Try
    '        If lvMappings.SelectedItems.Count > 0 Then
    '            Dim curIndex As Integer
    '            lvMappings.BeginUpdate()

    '            For i = (lvMappings.SelectedItems.Count - 1) To 0 Step -1
    '                curIndex = lvMappings.SelectedItems(i).Index

    '                objMap = lvMappings.Items(curIndex).Tag
    '                objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                objMap.MappingSource = Nothing
    '                objMap.IsMapped = "0"

    '                lvMappings.SelectedItems(i).SubItems(0).Text = " "
    '                lvMappings.SelectedItems(i).BackColor = BACK_COLOR
    '                lvMappings.SelectedItems(i).Font = New Font(lvMappings.Font, FontStyle.Regular)
    '            Next

    '            '//Fire change event
    '            OnChange(Me, New EventArgs)

    '        ElseIf lvMappings.SelectedItems.Count = 0 Then
    '            MsgBox("No item to delete", MsgBoxStyle.Exclamation, MsgTitle)
    '        Else
    '            MsgBox("Please select an item from the list", MsgBoxStyle.Exclamation, MsgTitle)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex)
    '    Finally
    '        lvMappings.EndUpdate()
    '    End Try

    'End Sub

    'Private Sub OnDelTarget(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuDelTarget.Click

    '    Dim i As Integer
    '    Dim objMap As clsMapping

    '    Try
    '        If lvMappings.SelectedItems.Count > 0 Then
    '            Dim curIndex As Integer
    '            lvMappings.BeginUpdate()

    '            For i = (lvMappings.SelectedItems.Count - 1) To 0 Step -1
    '                curIndex = lvMappings.SelectedItems(i).Index

    '                objMap = lvMappings.Items(curIndex).Tag
    '                objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE
    '                objMap.MappingTarget = Nothing
    '                objMap.IsMapped = "0"

    '                lvMappings.SelectedItems(i).SubItems(1).Text = " "
    '                lvMappings.SelectedItems(i).BackColor = BACK_COLOR
    '                lvMappings.SelectedItems(i).Font = New Font(lvMappings.Font, FontStyle.Regular)
    '            Next

    '            '//Fire change event
    '            OnChange(Me, New EventArgs)

    '        ElseIf lvMappings.SelectedItems.Count = 0 Then
    '            MsgBox("No item to delete", MsgBoxStyle.Exclamation, MsgTitle)
    '        Else
    '            MsgBox("Please select an item from the list", MsgBoxStyle.Exclamation, MsgTitle)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask mnuDelTarget_click")
    '    Finally
    '        lvMappings.EndUpdate()
    '    End Try

    'End Sub

#End Region

#Region "Mapping"

    '//This function will add/update mapping when new variable,lookup,join added 
    '//Added on 8/8/05 nd as TreeNode
    'Function ManualMapping(ByVal lvItm As ListViewItem, ByVal nd As TreeNode, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal CheckForAutoMap As Boolean = False) As Boolean

    '    Dim objMap As clsMapping
    '    Dim obj As INode '//Added on 8/8
    '    Dim objParentDS As INode '//added on 9/27/05
    '    Dim msgreturn As MsgBoxResult

    '    Try
    '        obj = nd.Tag
    '        '//if item dropped on exising list item (source or target ) which is blank then
    '        '//just change the mapping object properties 
    '        If Not (lvItm Is Nothing) Then
    '            objMap = lvItm.Tag

    '            If dirType = modDeclares.enumDirection.DI_SOURCE Then
    '                '//dont allow overwrite 
    '                '//TODO : Remove this logic if want to allow overwrite
    '                If objMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    msgreturn = MsgBox("There is already a source present, do you want to replace it?", MsgBoxStyle.YesNoCancel, MsgTitle)
    '                    If msgreturn = MsgBoxResult.Cancel Then
    '                        Exit Function
    '                    ElseIf msgreturn = MsgBoxResult.No Then
    '                        objMap = New clsMapping
    '                        objMap.SourceDataStore = GetParentDSForThisNode(nd).Text
    '                        objMap.SourceType = GetSourceTypeFromNodeType(obj.Type)
    '                        objMap.MappingSource = obj
    '                        AddMapping(objMap)
    '                        Exit Function
    '                    End If
    '                End If

    '                '//modified by npatel on 9/27/05
    '                objParentDS = GetParentDSForThisNode(nd)
    '                If objParentDS Is Nothing Then
    '                    objMap.SourceDataStore = ""
    '                Else
    '                    objMap.SourceDataStore = objParentDS.Text  'GetDatastoreFromNode(nd) '//8/8 by npatel
    '                End If

    '                objMap.SourceType = GetSourceTypeFromNodeType(obj.Type)
    '                objMap.MappingSource = obj
    '                If objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                    If CType(objMap.MappingSource, clsField).FieldName <> "" Then
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).FieldName 'objMap.SourceDataStore & "." & 
    '                    ElseIf CType(objMap.MappingSource, clsField).OrgName <> "" Then
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).OrgName 'objMap.SourceDataStore & "." & 
    '                    Else
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).CorrectedFieldName 'objMap.SourceDataStore & "." & 
    '                    End If
    '                ElseIf objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Then
    '                    If CType(objMap.MappingSource, clsSQFunction).SQFunctionSyntax.Contains("CDCIN") Then
    '                        CType(objMap.MappingSource, clsSQFunction).Text = CType(objMap.MappingSource, clsSQFunction).SQFunctionSyntax
    '                    Else
    '                        CType(objMap.MappingSource, clsSQFunction).Text = CType(objMap.MappingSource, clsSQFunction).SQFunctionWithInnerText
    '                    End If
    '                End If
    '                lvItm.SubItems(0).Text = objMap.MappingSource.Text
    '                lvItm.ImageIndex = ImgIdxFromName(objMap.MappingSource.Type)
    '                objMap.IsModified = True

    '                '//By default map item when its dropped
    '                If objMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    MapItem(lvItm.Index, , , 1)
    '                Else
    '                    '//Fire change event
    '                    OnChange(Me, New EventArgs)
    '                End If

    '            ElseIf dirType = modDeclares.enumDirection.DI_TARGET Then
    '                '//dont allow overwrite 
    '                '//TODO : Remove this logic if want to allow overwrite
    '                If objMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    msgreturn = MsgBox("There is already a Target present, do you want to replace it?", MsgBoxStyle.YesNoCancel, MsgTitle)
    '                    If msgreturn = MsgBoxResult.Cancel Then
    '                        Exit Function
    '                    ElseIf msgreturn = MsgBoxResult.No Then
    '                        objMap = New clsMapping
    '                        objMap.TargetDataStore = GetParentDSForThisNode(nd).Text
    '                        objMap.TargetType = GetSourceTypeFromNodeType(obj.Type)
    '                        objMap.MappingTarget = obj
    '                        AddMapping(objMap)
    '                        Exit Function
    '                    End If
    '                End If

    '                '//modified by npatel on 9/27/05
    '                objParentDS = GetParentDSForThisNode(nd)
    '                If objParentDS Is Nothing Then
    '                    objMap.TargetDataStore = ""
    '                Else
    '                    objMap.TargetDataStore = objParentDS.Text  'GetDatastoreFromNode(nd) '//8/8 by npatel
    '                End If

    '                objMap.TargetType = GetSourceTypeFromNodeType(obj.Type)
    '                objMap.MappingTarget = obj

    '                If (objMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_VAR) Then
    '                    If CType(objMap.MappingTarget, clsField).FieldName <> "" Then
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).FieldName ' objMap.TargetDataStore & "." & 
    '                    ElseIf CType(objMap.MappingTarget, clsField).OrgName <> "" Then
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).OrgName 'objMap.TargetDataStore & "." & 
    '                    Else
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).CorrectedFieldName 'objMap.TargetDataStore & "." & 
    '                    End If
    '                End If

    '                lvItm.SubItems(1).Text = objMap.MappingTarget.Text
    '                lvItm.ImageIndex = ImgIdxFromName(objMap.MappingTarget.Type)
    '                objMap.IsModified = True

    '                '//By default map item when its dropped
    '                If objMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    MapItem(lvItm.Index, , , 2)
    '                Else
    '                    '//Fire change event
    '                    OnChange(Me, New EventArgs)
    '                End If
    '            End If
    '        Else
    '            '//if item is not dropped on exising item then add new list item and create new mapping object
    '            objMap = New clsMapping

    '            If dirType = modDeclares.enumDirection.DI_SOURCE Then
    '                objMap.SourceDataStore = GetParentDSForThisNode(nd).Text 'GetDatastoreFromNode(nd) '//9/7 by npatel
    '                objMap.SourceType = GetSourceTypeFromNodeType(obj.Type)
    '                objMap.MappingSource = obj
    '                If objMap.SourceType = enumMappingType.MAPPING_TYPE_FUN Then
    '                    If CType(objMap.MappingSource, clsSQFunction).SQFunctionSyntax.Contains("CDCIN") Then
    '                        CType(objMap.MappingSource, clsSQFunction).Text = CType(objMap.MappingSource, clsSQFunction).SQFunctionSyntax
    '                    Else
    '                        CType(objMap.MappingSource, clsSQFunction).Text = CType(objMap.MappingSource, clsSQFunction).SQFunctionWithInnerText
    '                    End If
    '                End If
    '            ElseIf dirType = modDeclares.enumDirection.DI_TARGET Then
    '                objMap.TargetDataStore = GetParentDSForThisNode(nd).Text 'GetDatastoreFromNode(nd) '//9/7 by npatel
    '                objMap.TargetType = GetSourceTypeFromNodeType(obj.Type)
    '                objMap.MappingTarget = obj
    '            End If

    '            AddMapping(objMap)

    '        End If

    '    Catch ex As Exception
    '        LogError(ex)
    '    End Try

    'End Function

    ''//This function will auto map any new field to the matching item based on field name pattern
    ''//sometimes automapping can be incorrect then user can delete mapping and recreate new mapping
    ''// *** Modified for more detailed Mapping by TK 6-7/2007
    'Function AutoFieldMapping(ByVal objFld As clsField, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal DatastoreName As String = "", Optional ByVal Pass As Integer = 1) As Boolean

    '    Dim i As Integer
    '    Dim objMap As clsMapping
    '    Dim exact As Boolean = False
    '    Dim vague As Boolean = False
    '    Dim DLedit As Boolean = False
    '    Dim InList As Boolean = False

    '    Try
    '        '//Now loop through all unmapped items and if pattern if matched for this field for any unmapped field then map it
    '        For i = 0 To lvMappings.Items.Count - 1
    '            objMap = lvMappings.Items(i).Tag

    '            '//only look for unmapped items
    '            If objMap.IsMapped = "0" Then
    '                '//check any unmapped source field if we are dropping on target
    '                If dirType = modDeclares.enumDirection.DI_TARGET And _
    '                objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    '//we dropped on target so we need to check type of source which must be Field
    '                    If objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                        '/// first see if there is an exact match
    '                        Select Case Pass
    '                            Case 1
    '                                InList = CheckMapList(CType(objMap.MappingSource, clsField), objFld, dirType, objMap.SourceParent)
    '                            Case 2
    '                                exact = MatchPatternExact(CType(objMap.MappingSource, clsField), objFld, objMap.SourceParent)
    '                            Case 3
    '                                vague = MatchPattern(CType(objMap.MappingSource, clsField), objFld, objMap.SourceParent)
    '                            Case 4
    '                                DLedit = LastChanceMapping(CType(objMap.MappingSource, clsField), objFld, objMap.SourceParent)
    '                        End Select

    '                        If InList = True Or exact = True Or vague = True Or DLedit = True Then
    '                            objMap.TargetDataStore = DatastoreName
    '                            objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                            objMap.MappingTarget = objFld
    '                            '//Now show field name in target colum
    '                            lvMappings.Items(i).SubItems(1).Text = objFld.Text
    '                            '//also show little icon or text to show mapped item
    '                            AutoFieldMapping = MapItem(i, vague, DLedit, 2)
    '                            Exit Function
    '                        End If
    '                    End If
    '                    '//check any unmapped target field if we are dropping on source
    '                ElseIf dirType = modDeclares.enumDirection.DI_SOURCE And _
    '                objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
    '                    '//we dropped on source so we need to check type of target which must be Field
    '                    If objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                        Select Case Pass
    '                            Case 1
    '                                InList = CheckMapList(CType(objMap.MappingTarget, clsField), objFld, dirType, objMap.TargetParent)
    '                            Case 2
    '                                exact = MatchPatternExact(CType(objMap.MappingTarget, clsField), objFld, objMap.TargetParent)
    '                            Case 3
    '                                vague = MatchPattern(CType(objMap.MappingTarget, clsField), objFld, objMap.TargetParent)
    '                            Case 4
    '                                DLedit = LastChanceMapping(CType(objMap.MappingTarget, clsField), objFld, objMap.TargetParent)
    '                        End Select

    '                        If InList = True Or exact = True Or vague = True Or DLedit = True Then
    '                            '//if we come here means matching field found for auto mapping
    '                            objMap.SourceDataStore = DatastoreName '//new : 8/8
    '                            objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                            objMap.MappingSource = objFld
    '                            '//Now show field name in source column
    '                            lvMappings.Items(i).SubItems(0).Text = objFld.Text
    '                            '//also show little icon or text to show mapped item
    '                            AutoFieldMapping = MapItem(i, vague, DLedit, 1)
    '                            Exit Function
    '                        End If
    '                    End If 'If objMap.SourceType = ....
    '                End If 'If dirType = ....
    '            End If 'If objMap.IsMapped = ....
    '        Next

    '        '//if we come here means no luck, we didnt find any matching field so now add new lsit item at the last
    '        If Pass = 5 Then
    '            Dim objNewMap As New clsMapping
    '            If dirType = modDeclares.enumDirection.DI_SOURCE Then
    '                objNewMap.SourceDataStore = DatastoreName '//new : 8/8
    '                objNewMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                objNewMap.MappingSource = objFld
    '            ElseIf dirType = modDeclares.enumDirection.DI_TARGET Then
    '                objNewMap.TargetDataStore = DatastoreName '//new : 8/8
    '                objNewMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD
    '                objNewMap.MappingTarget = objFld
    '            End If
    '            objNewMap.Task = Me.objThis
    '            AddMapping(objNewMap)
    '            Return True
    '        End If

    '        Return False

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask AutoFieldMapping")
    '        Return False
    '    End Try

    'End Function

    ''/// redone 6/07 by TK
    'Function MapItem(ByVal lstItmIndex As Integer, Optional ByVal Vague As Boolean = False, Optional ByVal DLedit As Boolean = False, Optional ByVal SrcTgt As Integer = 0) As Boolean

    '    Try
    '        If Vague = True Then
    '            lvMappings.Items(lstItmIndex).BackColor = Maybe_mapped_Item_Color
    '            CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).IsMapped = "2"
    '        ElseIf DLedit = True Then
    '            lvMappings.Items(lstItmIndex).BackColor = DL_mapped_Item_Color
    '            CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).IsMapped = "3"
    '        Else
    '            lvMappings.Items(lstItmIndex).BackColor = MAPPED_ITEM_COLOR
    '            CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).IsMapped = "1"
    '        End If

    '        If SrcTgt <> 0 Then
    '            If SrcTgt = 1 Then
    '                If CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                    objThis.LastSrcFld = CType(CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).MappingSource, clsField).FieldName
    '                End If
    '            Else
    '                If CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
    '                    objThis.LastTgtFld = CType(CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).MappingTarget, clsField).FieldName
    '                End If
    '            End If
    '        End If

    '        lvMappings.Items(lstItmIndex).Font = New Font(lvMappings.Font, FontStyle.Bold)

    '        CType(lvMappings.Items(lstItmIndex).Tag, clsMapping).IsModified = True

    '        MapItem = True

    '        '//Fire change event
    '        OnChange(Me, New EventArgs)

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask MapItem")
    '        MapItem = False
    '    End Try

    'End Function

    ''/// New Function added 6/07 by TK
    'Function CheckMapList(ByVal fld1 As clsField, ByVal fld2 As clsField, ByVal dirtype As enumDirection, ByVal Parname As String) As Boolean

    '    Dim Proj As clsProject = fld1.Project
    '    Dim i As Integer
    '    Dim f1 As String = fld1.FieldName
    '    Dim f2 As String = fld2.FieldName

    '    Try
    '        For i = 1 To 4
    '            If i = 2 Then
    '                f1 = LCase(fld1.FieldName.Replace("_", "").Replace("-", ""))
    '                f2 = LCase(fld2.FieldName.Replace("_", "").Replace("-", ""))
    '            End If
    '            If i = 3 Then
    '                If f1.Contains(Parname) = True Then
    '                    If f1.Remove(0, (Strings.InStr(f1, Parname) + (Parname.Length - 1))).Length > 1 Then
    '                        f1 = f1.Remove(0, (Strings.InStr(f1, Parname) + (Parname.Length - 1)))
    '                    End If
    '                End If
    '            End If
    '            If i = 4 Then
    '                If f2.Contains(Parname) = True Then
    '                    If f2.Remove(0, (Strings.InStr(f2, Parname) + (Parname.Length - 1))).Length > 1 Then
    '                        f2 = f2.Remove(0, (Strings.InStr(f2, Parname) + (Parname.Length - 1)))
    '                    End If
    '                End If
    '            End If
    '            If dirtype = enumDirection.DI_SOURCE Then
    '                For Each map As clsMapPattern In Proj.Maplist
    '                    If (map.Source.Contains(f2) = True Or map.Source = f2) And (map.Target.Contains(f1) = True Or map.Target = f1) Then
    '                        Return True
    '                        Exit Function
    '                    End If
    '                Next
    '            Else
    '                For Each map As clsMapPattern In Proj.Maplist
    '                    If (map.Source.Contains(f1) = True Or map.Source = f1) And (map.Target.Contains(f2) = True Or map.Target = f2) Then
    '                        Return True
    '                        Exit Function
    '                    End If
    '                Next
    '            End If
    '        Next

    '        Return False

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask CheckMapList")
    '        Return False
    '    End Try

    'End Function

    ''//We auto map fields based on their name
    ''//This function will decide whether 2 fields are same
    ''/// modified 7/07 by TK
    'Function MatchPatternExact(ByVal fld1 As clsField, ByVal fld2 As clsField, ByVal ParName As String) As Boolean
    '    '/// Checks 1&2 look for exact match.

    '    Try
    '        Dim f1 As String = LCase(fld1.FieldName.Replace("_", "").Replace("-", ""))
    '        Dim f2 As String = LCase(fld2.FieldName.Replace("_", "").Replace("-", ""))
    '        Dim i As Integer

    '        ParName = LCase(ParName)
    '        For i = 1 To 3
    '            If i = 2 Then
    '                '/// Thkis takes out the parent's name (ie: Jstruct-field1 ==> parent is Jstruct => field1)
    '                If f1.Contains(ParName) = True Then
    '                    If f1.Remove(0, (Strings.InStr(f1, ParName) + (ParName.Length - 1))).Length > 1 Then
    '                        f1 = f1.Remove(0, (Strings.InStr(f1, ParName) + (ParName.Length - 1)))
    '                    End If
    '                End If
    '            ElseIf i = 3 Then
    '                If f2.Contains(ParName) = True Then
    '                    If f2.Remove(0, (Strings.InStr(f2, ParName) + (ParName.Length - 1))).Length > 1 Then
    '                        f2 = f2.Remove(0, (Strings.InStr(f2, ParName) + (ParName.Length - 1)))
    '                    End If
    '                End If
    '            End If
    '            '//Check1 : Exact same name (case-insensitive comparision)
    '            If f1 = f2 Then
    '                MatchPatternExact = True
    '                Exit Function
    '            End If
    '        Next

    '        Return False

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask MatchPatternExact")
    '        Return False
    '    End Try

    'End Function

    ''// This decides if two fields are Almost the same
    ''// Added 7/07 by TK
    'Function MatchPattern(ByVal fld1 As clsField, ByVal fld2 As clsField, Optional ByVal parname As String = "") As Boolean

    '    Try
    '        '// Just fieldNames as they are
    '        Dim Fstr1 As String = LCase(fld1.FieldName.Replace("_", "").Replace("-", ""))
    '        Dim Fstr2 As String = LCase(fld2.FieldName.Replace("_", "").Replace("-", ""))

    '        Dim Temp1 As String = Fstr1
    '        Dim Temp2 As String = Fstr2

    '        Dim i As Integer

    '        For i = 1 To 4
    '            If i = 1 Then
    '                If Temp1.Length <> Temp2.Length Then
    '                    If Temp1.Contains(Temp2) Then
    '                        MatchPattern = True
    '                        Exit Function
    '                    ElseIf Temp2.Contains(Temp1) Then
    '                        MatchPattern = True
    '                        Exit Function
    '                    End If
    '                End If
    '            ElseIf i = 2 Then
    '                If Temp1.Contains(parname) = True Then
    '                    If Temp1.Remove(0, (Strings.InStr(Temp1, parname) + (parname.Length - 1))).Length > 1 Then
    '                        Temp1 = Temp1.Remove(0, (Strings.InStr(Temp1, parname) + (parname.Length - 1)))
    '                    End If
    '                End If
    '                If Temp2.Contains(parname) = True Then
    '                    If Temp2.Remove(0, (Strings.InStr(Temp2, parname) + (parname.Length - 1))).Length > 1 Then
    '                        Temp2 = Temp2.Remove(0, (Strings.InStr(Temp2, parname) + (parname.Length - 1)))
    '                    End If
    '                End If
    '            ElseIf i = 3 Then
    '                If Temp1.Length > Temp2.Length Then
    '                    Temp1 = Strings.Right(Temp1, Temp2.Length)
    '                Else
    '                    Temp2 = Strings.Right(Temp2, Temp1.Length)
    '                End If
    '            ElseIf i = 4 Then
    '                Temp1 = Replace(Temp1, "a", "").Replace("e", "").Replace("i", "").Replace("o", "").Replace("u", "")
    '                Temp2 = Replace(Temp2, "a", "").Replace("e", "").Replace("i", "").Replace("o", "").Replace("u", "")
    '            End If

    '            If Temp1 = Temp2 Then
    '                MatchPattern = True
    '                Exit Function
    '            End If
    '        Next

    '        Return False

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask MatchPattern")
    '        Return False
    '    End Try

    'End Function

    ''// This decides if they May be the same
    ''// added 7/07 by TK
    'Function LastChanceMapping(ByVal fld1 As clsField, ByVal fld2 As clsField, Optional ByVal Parname As String = "") As Boolean
    '    '/// *********************
    '    '/// **** Last Chance ****
    '    '/// *********************
    '    '/// Now we try the Damerau-Levenshtein Distance
    '    '/// DL_distance looks for relative similarity
    '    '/// field names truncated with vowels removed
    '    Try

    '        Dim Fstr1 As String = LCase(fld1.FieldName.Replace("_", "").Replace("-", ""))
    '        Dim Fstr2 As String = LCase(fld2.FieldName.Replace("_", "").Replace("-", ""))

    '        Dim Temp1 As String = Fstr1
    '        Dim Temp2 As String = Fstr2

    '        Dim i As Integer
    '        Dim DLret As Integer

    '        For i = 1 To 5
    '            If i = 2 Then
    '                If Temp1.Contains(Parname) = True Then
    '                    If Temp1.Remove(0, (Strings.InStr(Temp1, Parname) + (Parname.Length - 1))).Length > 1 Then
    '                        Temp1 = Temp1.Remove(0, (Strings.InStr(Temp1, Parname) + (Parname.Length - 1)))
    '                        Fstr1 = Temp1
    '                    End If
    '                End If
    '                If Temp2.Contains(Parname) = True Then
    '                    If Temp2.Remove(0, (Strings.InStr(Temp2, Parname) + (Parname.Length - 1))).Length > 1 Then
    '                        Temp2 = Temp2.Remove(0, (Strings.InStr(Temp2, Parname) + (Parname.Length - 1)))
    '                        Fstr2 = Temp2
    '                    End If
    '                End If
    '            ElseIf i = 3 Then
    '                If Temp1.Length > Temp2.Length Then
    '                    Temp1 = Strings.Right(Temp1, Temp2.Length)
    '                Else
    '                    Temp2 = Strings.Right(Temp2, Temp1.Length)
    '                End If
    '            ElseIf i = 4 Then
    '                If Fstr1.Length > Fstr2.Length Then
    '                    Temp1 = Strings.Left(Fstr1, Fstr2.Length)
    '                Else
    '                    Temp2 = Strings.Left(Fstr2, Fstr1.Length)
    '                End If
    '            ElseIf i = 5 Then
    '                Temp1 = Replace(Temp1, "a", "").Replace("e", "").Replace("i", "").Replace("o", "").Replace("u", "")
    '                Temp2 = Replace(Temp2, "a", "").Replace("e", "").Replace("i", "").Replace("o", "").Replace("u", "")
    '            End If

    '            Dim Temp1arr As Char() = Temp1.ToCharArray
    '            Dim Temp2arr As Char() = Temp2.ToCharArray

    '            DLret = DL_EditDistance(Temp1arr, Temp2arr)
    '            If Temp1arr.Length < 5 Then
    '                If DLret < 2 Then
    '                    LastChanceMapping = True
    '                    Exit Function
    '                Else
    '                    LastChanceMapping = False
    '                End If
    '            ElseIf Temp1arr.Length < 8 Then
    '                If DLret < 3 Then
    '                    LastChanceMapping = True
    '                    Exit Function
    '                Else
    '                    LastChanceMapping = False
    '                End If
    '            Else
    '                If DLret < 4 Then
    '                    LastChanceMapping = True
    '                    Exit Function
    '                Else
    '                    LastChanceMapping = False
    '                End If
    '            End If
    '        Next

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask LastChanceMapping")
    '        Return False
    '    End Try

    'End Function

    ''// This function Uses Damerau-Levenshtein Distance to determine
    ''// How close the 2 field names are to matching.
    ''// added 7/07 by TK
    'Function DL_EditDistance(ByVal CharStr1 As Char(), ByVal CharStr2 As Char()) As Integer
    '    ' ... Here is the algorithm  ... converted this to VB.NET below

    '    'int DamerauLevenshteinDistance(char str1[1..lenStr1], char str2[1..lenStr2])
    '    '   // d is a table with lenStr1+1 rows and lenStr2+1 columns
    '    '   declare int d[0..lenStr1, 0..lenStr2]
    '    '   // i and j are used to iterate over str1 and str2
    '    '   declare int i, j, cost

    '    '   for i from 0 to lenStr1
    '    '       d[i, 0] := i
    '    '   LOOP

    '    '   for j from 1 to lenStr2
    '    '       d[0, j] := j
    '    '   LOOP

    '    '   for i from 1 to lenStr1
    '    '       for j from 1 to lenStr2
    '    '           if str1[i] = str2[j] then cost := 0
    '    '                                else cost := 1
    '    '           d[i, j] := minimum(
    '    '                                d[i-1, j  ] + 1,     // deletion
    '    '                                d[i  , j-1] + 1,     // insertion
    '    '                                d[i-1, j-1] + cost   // substitution
    '    '                            )
    '    '           if(i > 1 and j > 1 and str1[i] = str2[j-1] and str1[i-1] = str2[j]) then
    '    '               d[i, j] := minimum(
    '    '                                d[i, j],
    '    '                                d[i-2, j-2] + cost   // transposition
    '    '                             )
    '    '   This is where we LOOP

    '    '   return d[lenStr1, lenStr2]

    '    '// Converted to VB.NET this is the same algorithm
    '    Try
    '        '// two dimensional array .. Matrix to compare 
    '        Dim D(CharStr1.Length + 1, CharStr2.Length + 1) As Integer
    '        '// rows, columns, cost, temp
    '        Dim i, j, cost, min1 As Integer

    '        '// initialize matrix rows
    '        For i = 0 To CharStr1.Length
    '            D(i, 0) = i
    '        Next
    '        '// initialize matrix Columns
    '        For j = 1 To CharStr2.Length
    '            D(0, j) = j
    '        Next

    '        '// go through matrix and quantify differences in strings into number result
    '        For i = 1 To CharStr1.Length
    '            For j = 1 To CharStr2.Length
    '                If CharStr1(i - 1) = CharStr2(j - 1) Then
    '                    cost = 0
    '                Else
    '                    cost = 1
    '                End If
    '                min1 = System.Math.Min(D(i - 1, j) + 1, D(i, j - 1) + 1)  '// compare deletion and insertion
    '                D(i, j) = System.Math.Min(min1, D(i - 1, j - 1) + cost)  '// compare result above w/ substitution
    '                '/// Now compare and try Transposition
    '                '/// this is the Damerau addition to the Levenshtein Distance
    '                If (i > 1 And j > 1) Then
    '                    If (CharStr1(i - 1) = CharStr2(j - 2) And CharStr1(i - 2) = CharStr2(j - 1)) Then
    '                        D(i, j) = System.Math.Min(D(i, j), D(i - 2, j - 2) + cost)
    '                    End If
    '                End If
    '            Next
    '        Next

    '        '/// the less alike the strings are, the higher the number result
    '        DL_EditDistance = D(CharStr1.Length, CharStr2.Length)

    '    Catch ex As Exception
    '        LogError(ex, "DL_editDist ctlTask")
    '        Return 100
    '    End Try

    'End Function

    ''//This function returns item row and column in the listview
    ''//If no item present then it will return row= -1 and col= index of column from x cordinate
    'Function GetListSubItemFromPoint(ByVal lv As ListView, ByVal X As Integer, ByVal Y As Integer, Optional ByRef retRow As Integer = 0, Optional ByRef retCol As Integer = 0) As Boolean

    '    Dim flag As Boolean
    '    Dim cleanupFlag As Boolean
    '    Dim Item As ListViewItem = Nothing
    '    Dim col, row As Integer

    '    Try
    '        flag = lv.FullRowSelect

    '        '//FullRow select must be true in order to use GetItemAt properly
    '        If flag = False Then lv.FullRowSelect = True '//temperory make it true

    '        If lv.Items.Count <= 0 And lv.Columns.Count > 0 Then
    '            '//if no item then to get subitem add dummy listitem/subitems and then very last delete it 
    '            lv.BeginUpdate() '//no update to listview until we are done

    '            With lv.Items.Add(".")
    '                For col = 0 To lv.Columns.Count - 1
    '                    .SubItems.Add("..")
    '                Next
    '                '//Y cordinate can be any where in the list view but shift it on first item so GetItemAt return item
    '                Y = .GetBounds(ItemBoundsPortion.Label).Y
    '            End With
    '            Item = lv.GetItemAt(X, Y)
    '            row = -1
    '            cleanupFlag = True '//this indicates we added one dummy item which should be cleaned after we find coordinates
    '        ElseIf lv.Items.Count > 0 And lv.Columns.Count > 0 Then
    '            Item = lv.GetItemAt(X, Y)
    '            '//if item nothing then may be user after last item 
    '            '//if user drop item after last listitem then we only care about column index so 
    '            '//just get Y for first list item and X from actual dropped cordinate
    '            If Item Is Nothing Then
    '                Y = lv.Items(0).GetBounds(ItemBoundsPortion.Label).Y
    '                row = -1
    '                Item = lv.GetItemAt(X, Y)
    '            End If
    '        End If

    '        lv.FullRowSelect = flag '//switch back to old value

    '        If Not Item Is Nothing Then
    '            Dim I As Integer = 0
    '            Dim R As Rectangle = Item.GetBounds(ItemBoundsPortion.Label)
    '            Do While I < Item.SubItems.Count
    '                If R.Contains(X, Y) Then
    '                    'MsgBox("Column=" & CStr(I) & "; Row=" & CStr(Item.Index) & "  Text: " & Item.SubItems(I).Text)
    '                    retRow = IIf(row = -1, -1, Item.Index)
    '                    retCol = I

    '                    GetListSubItemFromPoint = True
    '                    Exit Do
    '                End If
    '                If I > 0 Then  '// 9/20/2007 fixed bug that sets lv index out of range
    '                    R.X = R.X + lv.Columns(1).Width
    '                    R.Width = lv.Columns(1).Width
    '                Else
    '                    R.X = R.X + lv.Columns(I).Width
    '                    R.Width = lv.Columns(I + 1).Width
    '                End If
    '                I = I + 1
    '            Loop
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask GetSubItemFromPoint")
    '        GetListSubItemFromPoint = False
    '    Finally
    '        If cleanupFlag = True Then
    '            lv.Items.Clear()
    '        End If
    '        lv.EndUpdate()
    '    End Try

    'End Function

    ''//Add new mapping based on position
    'Function AddMapping(Optional ByVal Pos As Integer = -1) As Boolean

    '    Dim lvItm As New ListViewItem
    '    Dim objMap As clsMapping

    '    Try
    '        With lvItm
    '            .Text = " "
    '            .SubItems.Add(" ")
    '            .SubItems.Add(" ")
    '            .ImageIndex = ImgIdxFromName(NODE_STRUCT_FLD)
    '            objMap = New clsMapping

    '            .Tag = objMap
    '        End With

    '        If Pos < 0 Then
    '            '//add new item at the end
    '            lvMappings.Items.Add(lvItm)
    '        Else
    '            '//add new item above the selected item
    '            lvMappings.Items.Insert(Pos, lvItm)
    '            MarkAsModifedBelowThis(Pos)
    '        End If

    '        objMap.SeqNo = lvItm.Index

    '        lvItm.Selected = True
    '        lvItm.EnsureVisible()
    '        AddMapping = True

    '        '//Fire change event
    '        OnChange(Me, New EventArgs)

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask AddMapping")
    '    End Try

    'End Function

    ''//Add new mapping based on position and Mapping object
    'Function AddMapping(ByVal objMap As clsMapping, Optional ByVal Pos As Integer = -1) As Boolean

    '    Dim lvItm As New ListViewItem
    '    Try
    '        With lvItm
    '            '//Make sure that we dont have any SeqNo beyont the list item count
    '            If Pos > lvMappings.Items.Count Then
    '                Pos = lvMappings.Items.Count
    '                objMap.SeqNo = Pos
    '            End If
    '            If objMap.SourceType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And (objMap.MappingSource Is Nothing) = False Then
    '                '//added this logic to handle issue#48 : 9/7/05 by npatel
    '                If objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                    If CType(objMap.MappingSource, clsField).FieldName <> "" Then
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).FieldName 'objMap.SourceDataStore & "." & 
    '                    ElseIf CType(objMap.MappingSource, clsField).OrgName <> "" Then
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).OrgName 'objMap.SourceDataStore & "." & 
    '                    Else
    '                        CType(objMap.MappingSource, clsField).Text = CType(objMap.MappingSource, clsField).ParentStructureName & "." & CType(objMap.MappingSource, clsField).CorrectedFieldName 'objMap.SourceDataStore & "." & 
    '                    End If
    '                    .Text = CType(objMap.MappingSource, clsField).Text
    '                Else
    '                    .Text = CType(objMap.MappingSource, INode).Text
    '                End If
    '            Else
    '                .Text = " "
    '            End If

    '            lvItm.SubItems.Add(" ")
    '            lvItm.SubItems.Add(" ")

    '            If objMap.TargetType <> modDeclares.enumMappingType.MAPPING_TYPE_NONE And (objMap.MappingTarget Is Nothing) = False Then
    '                If objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                    If CType(objMap.MappingTarget, clsField).FieldName <> "" Then
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).FieldName ' objMap.TargetDataStore & "." & 
    '                    ElseIf CType(objMap.MappingTarget, clsField).OrgName <> "" Then
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).OrgName ' objMap.TargetDataStore & "." & 
    '                    Else
    '                        CType(objMap.MappingTarget, clsField).Text = CType(objMap.MappingTarget, clsField).ParentStructureName & "." & CType(objMap.MappingTarget, clsField).CorrectedFieldName ' objMap.TargetDataStore & "." & 
    '                    End If

    '                    lvItm.SubItems(1).Text = CType(objMap.MappingTarget, clsField).Text

    '                Else
    '                    lvItm.SubItems(1).Text = CType(objMap.MappingTarget, INode).Text
    '                End If
    '            End If

    '            .ImageIndex = ImgIdxFromName(GetNodeTypeFromSourceType(objMap.SourceType))

    '            .Tag = objMap

    '        End With

    '        If Pos < 0 Then
    '            '//add new item at the end
    '            lvMappings.Items.Add(lvItm)
    '        Else
    '            '//add new item above the selected item
    '            lvMappings.Items.Insert(Pos, lvItm)
    '            MarkAsModifedBelowThis(Pos)
    '        End If
    '        If objMap.IsMapped = "1" Then
    '            MapItem(lvItm.Index)
    '        ElseIf objMap.IsMapped = "2" Then
    '            MapItem(lvItm.Index, True)
    '        ElseIf objMap.IsMapped = "3" Then
    '            MapItem(lvItm.Index, False, True)
    '        End If

    '        If objMap.MappingDesc <> "" Then
    '            lvMappings.Items(lvItm.Index).ForeColor = HAS_DESC_COLOR
    '        End If

    '        If objMap.HasMissingDependency = True Then
    '            lvMappings.Items(lvItm.Index).BackColor = MISSING_DEP_ITEM_COLOR 'Color.Yellow
    '            'Else
    '            '    lvMappings.Items(lvItm.Index).BackColor = BACK_COLOR
    '        End If

    '        objMap.SeqNo = lvItm.Index

    '        lvItm.EnsureVisible()
    '        AddMapping = True

    '        '//Fire change event
    '        OnChange(Me, New EventArgs)

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask AddMapping")
    '    End Try

    'End Function

    'Private Sub AddNewMapping(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim flagOutside As Boolean

    '    Try
    '        Dim pt As Point = lvMappings.PointToClient(New Point(MousePosition.X, MousePosition.Y))

    '        flagOutside = (CurRow = -1)

    '        If flagOutside = True Then
    '            '//add new item at the end
    '            AddMapping()
    '        ElseIf CurRow = lvMappings.Items.Count - 1 Then
    '            AddMapping()
    '        Else
    '            '//add new item below the selected item
    '            CurRow = CurRow + 1
    '            AddMapping(CurRow)

    '            MarkAsModifedBelowThis(CurRow)

    '            OnChange(Me, New EventArgs)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask AddNewMapping")
    '    End Try

    'End Sub

    'Private Sub MoveUp(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If lvMappings.SelectedItems.Count > 0 Then
    '        SetMappingPosition(lvMappings.SelectedItems(0).Index - 1)
    '    End If

    'End Sub

    'Private Sub MoveDown(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    If lvMappings.SelectedItems.Count > 0 Then
    '        SetMappingPosition(lvMappings.SelectedItems(0).Index + 1)
    '    End If

    'End Sub

    ''//This function will set new position for all selected listview items
    ''//New posiion will be based on first item in the selection list.
    ''//NewPosition : is New Position of the first item in selection group
    'Function SetMappingPosition(ByVal NewPosition As Integer) As Boolean

    '    Try
    '        If lvMappings.SelectedItems.Count > 0 Then
    '            Dim OldPosition As Integer

    '            '//How many items to move (up or down). All items will be shifted by this offset
    '            Dim moveOffset As Integer
    '            Dim newIndex As Integer

    '            OldPosition = lvMappings.SelectedItems(0).Index

    '            '//if offset is +ve then shift all selected item down by moveOffset
    '            '//if offset is -ve then shift all selected item up by moveOffset

    '            moveOffset = NewPosition - OldPosition

    '            If moveOffset = 0 Then Exit Function

    '            '//Any item below the current item must be updated coz SeqNo has been changed
    '            '//since we changed the positions we have to update IsModified flag for some items
    '            Dim startIndex, endIndex As Integer
    '            If moveOffset > 0 Then '// +ve means moved down
    '                startIndex = lvMappings.SelectedItems(0).Index '//old position of first item in selection list
    '                endIndex = lvMappings.SelectedItems(lvMappings.SelectedItems.Count - 1).Index + moveOffset '//new position of last item in selection list
    '            ElseIf moveOffset < 0 Then '// -ve means moved up
    '                startIndex = lvMappings.SelectedItems(0).Index + moveOffset '//new position of first item in selection list
    '                endIndex = lvMappings.SelectedItems(lvMappings.SelectedItems.Count - 1).Index  '//old position of last item in the selection list
    '            End If

    '            '//Check for valid new position
    '            If (endIndex > (lvMappings.Items.Count - 1)) Then Exit Function
    '            If (startIndex < 0) Or (startIndex > (lvMappings.Items.Count - 1)) Then Exit Function

    '            '////////////

    '            Dim objMap As clsMapping
    '            Dim curIndex As Integer
    '            Dim i As Integer

    '            lvMappings.BeginUpdate()

    '            '//If move down then start repositioning with last item in selection list 
    '            '//If move up then start repositioning with first item in selection list 
    '            '//Reposition selected items

    '            If moveOffset > 0 Then '//if move down
    '                For i = lvMappings.SelectedItems.Count - 1 To 0 Step -1
    '                    curIndex = lvMappings.SelectedItems(i).Index
    '                    newIndex = curIndex + moveOffset
    '                    If newIndex >= 0 Then
    '                        objMap = lvMappings.SelectedItems(i).Tag
    '                        lvMappings.Items.RemoveAt(curIndex)
    '                        AddMapping(objMap, newIndex)
    '                        lvMappings.Items(newIndex).Selected = True
    '                        lvMappings.Items(newIndex - 1).Focused = False
    '                    End If
    '                Next
    '            ElseIf moveOffset < 0 Then '//if move up
    '                For i = 0 To lvMappings.SelectedItems.Count - 1
    '                    curIndex = lvMappings.SelectedItems(i).Index
    '                    newIndex = curIndex + moveOffset
    '                    If newIndex >= 0 Then
    '                        objMap = lvMappings.SelectedItems(i).Tag
    '                        lvMappings.Items.RemoveAt(curIndex)
    '                        AddMapping(objMap, newIndex)
    '                        lvMappings.Items(newIndex).Selected = True
    '                        lvMappings.Items(newIndex + 1).Focused = False
    '                        'lvMappings.FullRowSelect = True
    '                        'MsgBox("Checking", MsgBoxStyle.Exclamation)
    '                    End If
    '                Next
    '            End If
    '            '//update IsModified flags for all affected itemas which SeqNo have been changed
    '            For i = startIndex To endIndex
    '                CType(lvMappings.Items(i).Tag, INode).IsModified = True
    '            Next

    '            If Not lvMappings.FocusedItem Is Nothing Then
    '                lvMappings.FocusedItem.Focused = False
    '            End If
    '            ResetMappingSeqNo() '//new by npatel on 8/15/05
    '            SetMappingPosition = True
    '        Else
    '            MsgBox("Please select an item from the list", MsgBoxStyle.Exclamation, MsgTitle)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex)
    '    Finally
    '        lvMappings.EndUpdate()
    '    End Try

    'End Function

    'Sub ResetMappingSeqNo()

    '    Dim i As Integer

    '    Try
    '        For i = 0 To lvMappings.Items.Count - 1
    '            CType(lvMappings.Items(i).Tag, INode).SeqNo = i
    '        Next

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask resetMappingSeqNo")
    '    End Try

    'End Sub

    'Private Sub DelMapping(ByVal sender As System.Object, ByVal e As System.EventArgs)

    '    Dim i As Integer

    '    Try
    '        If lvMappings.SelectedItems.Count > 0 Then
    '            Dim curIndex As Integer
    '            lvMappings.BeginUpdate()

    '            For i = (lvMappings.SelectedItems.Count - 1) To 0 Step -1
    '                curIndex = lvMappings.SelectedItems(i).Index

    '                If curIndex + 1 < lvMappings.Items.Count Then
    '                    lvMappings.Items(curIndex + 1).Selected = True
    '                ElseIf curIndex - 1 >= 0 Then
    '                    lvMappings.Items(curIndex - 1).Selected = True
    '                End If
    '                lvMappings.Items.RemoveAt(curIndex)
    '            Next

    '            ResetMappingSeqNo() '//new on 8/15/05

    '            '//Fire change event
    '            OnChange(Me, New EventArgs)

    '        ElseIf lvMappings.SelectedItems.Count = 0 Then
    '            MsgBox("No item to delete", MsgBoxStyle.Exclamation, MsgTitle)
    '        Else
    '            MsgBox("Please select an item from the list", MsgBoxStyle.Exclamation, MsgTitle)
    '        End If

    '    Catch ex As Exception
    '        LogError(ex)
    '    Finally
    '        lvMappings.EndUpdate()
    '    End Try

    'End Sub

#End Region

#Region "Misc Functions"

    'Function MarkAsModifedBelowThis(ByVal idx As Integer) As Boolean

    '    Dim i As Integer

    '    For i = idx To lvMappings.Items.Count - 1
    '        CType(lvMappings.Items(i).Tag, INode).IsModified = True
    '    Next

    'End Function

    '//if dirType=DI_SOURCE then only look in sources else only look in target
    'Function IsDuplicateItem(ByVal objFld As clsField, Optional ByVal dirType As enumDirection = modDeclares.enumDirection.DI_SOURCE, Optional ByVal CheckonlyMappedItems As Boolean = True) As Boolean

    '    Dim objMap As clsMapping

    '    Try
    '        For Each lvItm As ListViewItem In lvMappings.Items
    '            objMap = lvItm.Tag
    '            If dirType = modDeclares.enumDirection.DI_SOURCE Then
    '                '//Check what type of source in current mapping (we only care about field)
    '                If objMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then
    '                    '//return if we find same field in listview
    '                    If (CheckonlyMappedItems = True) And (Not (objMap.MappingSource Is Nothing)) Then
    '                        If CType(objMap.MappingSource, clsField).FieldName = objFld.FieldName And _
    '                            CType(objMap.MappingSource, clsField).Struct.Text = objFld.Struct.Text And _
    '                            CType(objMap.MappingSource, clsField).Struct.Environment.Text = objFld.Struct.Environment.Text And _
    '                            CType(objMap.MappingSource, clsField).Struct.Environment.Project.Text = objFld.Struct.Environment.Project.Text And (objMap.IsMapped = "1" Or objMap.IsMapped = "2" Or objMap.IsMapped = "3") Then

    '                            IsDuplicateItem = True
    '                            Exit Function
    '                        End If
    '                    Else
    '                        If Not (objMap.MappingSource Is Nothing) Then
    '                            If CType(objMap.MappingSource, clsField).FieldName = objFld.FieldName And _
    '                            CType(objMap.MappingSource, clsField).Struct.Text = objFld.Struct.Text And _
    '                            CType(objMap.MappingSource, clsField).Struct.Environment.Text = objFld.Struct.Environment.Text And _
    '                            CType(objMap.MappingSource, clsField).Struct.Environment.Project.Text = objFld.Struct.Environment.Project.Text Then
    '                                IsDuplicateItem = True
    '                                Exit Function
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            Else
    '                If objMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Then

    '                    '//Check what type of target in current mapping (we only care about field)
    '                    If (CheckonlyMappedItems = True) And (Not (objMap.MappingTarget Is Nothing)) Then
    '                        If CType(objMap.MappingTarget, clsField).FieldName = objFld.FieldName And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Text = objFld.Struct.Text And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Environment.Text = objFld.Struct.Environment.Text And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Environment.Project.Text = objFld.Struct.Environment.Project.Text And (objMap.IsMapped = "1" Or objMap.IsMapped = "2" Or objMap.IsMapped = "3") Then
    '                            IsDuplicateItem = True
    '                            Exit Function
    '                        End If
    '                    Else
    '                        If Not (objMap.MappingTarget Is Nothing) Then
    '                            If CType(objMap.MappingTarget, clsField).Text = objFld.Text And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Text = objFld.Struct.Text And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Environment.Text = objFld.Struct.Environment.Text And _
    '                            CType(objMap.MappingTarget, clsField).Struct.Environment.Project.Text = objFld.Struct.Environment.Project.Text Then
    '                                IsDuplicateItem = True
    '                                Exit Function
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        Next

    '        Return False

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask IsDupItem")
    '    End Try

    'End Function

    Function GetTypeFromText(ByVal txt As String, Optional ByVal dir As enumDirection = modDeclares.enumDirection.DI_SOURCE) As enumMappingType

        If txt = "" Then
            Return modDeclares.enumMappingType.MAPPING_TYPE_NONE
        ElseIf txt.IndexOf("(") >= 0 Or txt.IndexOf(")") >= 0 Or txt.IndexOf(" ") >= 0 Then
            Return modDeclares.enumMappingType.MAPPING_TYPE_FUN
        ElseIf (txt.StartsWith("'") And txt.EndsWith("'")) Or IsNumeric(txt) Then
            Return modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT
            'KS02 4/24/2006
        ElseIf (txt.StartsWith("`") And txt.EndsWith("`")) Or IsNumeric(txt) Then
            Return modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT
            'End KS02 4/24/2006
        Else
            If dir = modDeclares.enumDirection.DI_SOURCE Then
                Return objCurMap.SourceType
            Else
                Return objCurMap.TargetType
            End If
        End If

    End Function

    '//Switch back to mapping listview when user click on back button 
    'Function HideScriptEditor(Optional ByVal SkipSave As Boolean = False) As Boolean

    '    'If SkipSave = False Then
    '    '    If SaveCurrentScript() = False Then
    '    '        Exit Function
    '    '    End If
    '    'End If
    '    pnlScript.Visible = False
    '    'pnlDesc.Visible = False
    '    'pnlSourceTarget.Visible = True
    '    'pnlSourceTarget.BringToFront()

    '    IsCodeEditorOnTop = False
    '    IsDescEditorOnTop = False

    'End Function

    'Function HideDescEditor(Optional ByVal SkipSave As Boolean = False) As Boolean

    '    'If SkipSave = False Then
    '    '    If SaveMapDesc() = False Then
    '    '        Exit Function
    '    '    End If
    '    'End If

    '    'pnlDesc.Visible = False
    '    pnlScript.Visible = False
    '    'pnlSourceTarget.Visible = True
    '    'pnlSourceTarget.BringToFront()

    '    IsCodeEditorOnTop = False
    '    IsDescEditorOnTop = False

    'End Function

    Function TruncatedFieldName(ByVal txt As String) As String

        txt = txt.Substring(txt.LastIndexOf(".") + 1)
        Return txt

    End Function

    Private Sub txtCodeEditor_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        ' show a block cursor of 3x16
        ShowCustomCaret(sender, 3, 16)

    End Sub

    Function GetParentDSForThisNode(ByVal nd As TreeNode) As INode

        Const MAX_LOOP As Integer = 200
        Dim cnt As Integer

        Try
            If CType(nd.Tag, INode).Type = NODE_GEN Or CType(nd.Tag, INode).Type = NODE_LOOKUP Or CType(nd.Tag, INode).Type = NODE_FUN Or CType(nd.Tag, INode).Type = NODE_TEMPLATE Or CType(nd.Tag, INode).Type = NODE_VARIABLE Then
                Return nd.TreeView.Nodes(0).Tag '//return datastore which is always first node in the treeview
            End If

            Do While (CStr(nd.Tag.type) <> NODE_SOURCEDATASTORE And _
            CStr(nd.Tag.type) <> NODE_TARGETDATASTORE) = True And _
            (cnt < MAX_LOOP) = True

                nd = nd.Parent
                '//new by npatel on 9/7/05
                If Not nd Is Nothing Then
                    If nd.Tag.type = NODE_LOOKUP Or nd.Tag.Type = NODE_GEN Or nd.Tag.type = NODE_VARIABLE Then
                        CType(nd.Tag, clsTask).LoadDatastores()
                        If CType(nd.Tag, clsTask).ObjSources.Count > 0 Then
                            Return CType(nd.Tag, clsTask).ObjSources(0)
                        Else
                            Return Nothing
                        End If
                    End If
                End If
                cnt = cnt + 1
            Loop

            Return nd.Tag

        Catch ex As Exception
            LogError(ex, "ctlTask GetParentDSForThisNode")
            Return Nothing
        End Try

    End Function

    'Function GetParentDSForField(ByVal FldNode As TreeNode) As String

    '    Try
    '        Dim tempnode As TreeNode = FldNode.Parent

    '        If tempnode.Tag.type = NODE_SOURCEDSSEL Or tempnode.Tag.type = NODE_TARGETDSSEL Then
    '            Return CType(tempnode.Tag, clsDSSelection).ObjDatastore.DatastoreName
    '            Exit Function
    '        Else
    '            GetParentDSForField(tempnode)
    '        End If

    '        Return Nothing

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask GetParentDSforField")
    '        Return Nothing
    '    End Try

    'End Function

    Function AddToRecentlyUsedFunctionList(ByVal draggedNode As TreeNode) As Boolean

        Dim ndLast As TreeNode, nd As TreeNode
        Dim IsFound As Boolean
        Dim objFun As clsSQFunction
        Dim cnt As Integer

        Try
            cnt = tvFunctions.GetNodeCount(False)
            If cnt <= 0 Then Exit Function

            objFun = draggedNode.Tag

            ndLast = tvFunctions.Nodes(cnt - 1)

            For Each nd In ndLast.Nodes
                If CType(nd.Tag, INode).Text = objFun.Text Then
                    IsFound = True
                    Exit For
                End If
            Next
            '//if this function is not in the list then add it
            If IsFound = False Then
                Dim newNd As TreeNode
                newNd = draggedNode.Clone
                ndLast.Nodes.Add(newNd)
                ndLast.Expand()
                newNd.EnsureVisible()
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "AddToRecentlyUsedFunctList")
            Return False
        End Try

    End Function

    '//This will generate script for template. 
    '//We will generate scripts only for system templates 
    '//some user defined templates will remain as it is
    'Function UpdateTemplateScript(ByVal obj As clsSQFunction) As Boolean

    '    If obj.IsTemplate = False Then Exit Function

    '    Select Case obj.ObjTreeNode.Text
    '        Case "Main" '//Main template
    '            obj.SQFunctionName = GetMainText() 'GetScriptForMainTemplateV3() 'changed 11/07 by TK
    '            obj.SQFunctionWithInnerText = obj.SQFunctionName
    '        Case "Procedure"
    '            obj.SQFunctionName = GetScriptForProcV3()
    '            obj.SQFunctionWithInnerText = obj.SQFunctionName
    '        Case "LOOK"
    '            obj.SQFunctionName = GetScriptForLOOK()
    '            obj.SQFunctionWithInnerText = obj.SQFunctionName
    '    End Select

    'End Function

    Function GetScriptForMainTemplate() As String

        Try
            Dim sb As New System.Text.StringBuilder
            Dim nd As TreeNode
            Dim inDsNode As TreeNode = Nothing

            '//inDs = InputBox("Enter name of your input datastore", , "<Source Datastore>")

            '//TODO : Give user selection for input datastore
            '//Note at this moment just pick first ds but it can be any of the input datastore
            If tvSource.GetNodeCount(False) <= 0 Then
                GetScriptForMainTemplate = ""
                Exit Function
            End If


            For Each nd In tvSource.Nodes
                If nd.GetNodeCount(True) > 0 Then inDsNode = nd : Exit For 'tvSource.Nodes(0)
            Next

            If inDsNode Is Nothing Then
                GetScriptForMainTemplate = ""
                Exit Function
            End If


            '//REPLACE INTO Target1, Target2 .... is pending
            'sb.Append("SELECT" & vbCrLf)
            'sb.Append("IF" & vbCrLf)
            sb.Append("CASE" & vbCrLf)
            sb.Append("(" & vbCrLf)
            For Each nd In inDsNode.Nodes
                sb.Append(vbTab & IIf(nd.Index = 0, "", ",") & "EQ(RECNAME(" & inDsNode.Text & "),'" & nd.Text & "')" & vbCrLf)
                sb.Append(vbTab & ",CALLPROC( )" & vbCrLf)
                sb.Append(vbCrLf)
            Next
            sb.Append(")" & vbCrLf)
            'sb.Append("FROM " & inDsNode.Text)
            GetScriptForMainTemplate = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlTask GetScriptForMainTemplate")
            Return ""
        End Try

    End Function

    '/// version 3 Main Template
    Function GetScriptForMainTemplateV3() As String

        Try
            Dim sb As New System.Text.StringBuilder
            Dim nd As TreeNode
            Dim inDsNode As TreeNode = Nothing

            '//inDs = InputBox("Enter name of your input datastore", , "<Source Datastore>")

            '//TODO : Give user selection for input datastore
            '//Note at this moment just pick first ds but it can be any of the input datastore
            If tvSource.GetNodeCount(False) <= 0 Then
                GetScriptForMainTemplateV3 = ""
                Exit Function
            End If

            For Each nd In tvSource.Nodes
                If nd.GetNodeCount(True) > 0 Then inDsNode = nd : Exit For 'tvSource.Nodes(0)
            Next

            If inDsNode Is Nothing Then
                GetScriptForMainTemplateV3 = ""
                Exit Function
            End If

            '//REPLACE INTO Target1, Target2 .... is pending
            'sb.Append("SELECT" & vbCrLf)
            'sb.Append("IF" & vbCrLf)
            sb.Append("{" & vbCrLf)
            sb.Append("CASE RECNAME(" & inDsNode.Text & ")" & vbCrLf)
            For Each nd In inDsNode.Nodes
                sb.Append(TAB & "WHEN '" & nd.Text & "'" & TAB & TAB & "CALLPROC( )" & vbCrLf) '& IIf(nd.Index = 0, "", ",")
                'sb.Append(TAB & "DO" & vbCrLf)
                'sb.Append(TAB & TAB & "CALLPROC( )" & vbCrLf)
                'sb.Append(TAB & "END" & vbCrLf)
                'sb.Append(vbCrLf)
            Next
            'sb.Append(")" & vbCrLf)
            sb.Append("}")   '& inDsNode.Text
            GetScriptForMainTemplateV3 = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlTask GetScriptForMainTemplate")
            Return ""
        End Try

    End Function

    Private Function GetMainText() As String

        Dim sb As New System.Text.StringBuilder
        Dim sel As clsDSSelection
        'Dim i As Integer
        Dim taskName As String
        Dim sDs As clsDatastore = Nothing
        Dim nDs As clsDatastore = Nothing
        Dim count As Integer = 0
        Dim SrcCount As Integer = 0

        Try
            While SrcCount < objThis.Engine.Sources.Count

                sDs = objThis.Engine.Sources(SrcCount + 1)


                If SrcCount > 0 Then
                    nDs = objThis.Engine.Sources(SrcCount)
                    If sDs.IsLookUp = False Then
                        sb.AppendLine("}")   '& inDsNode.Text
                        sb.AppendLine("FROM " & nDs.DatastoreName)
                        sb.AppendLine("UNION")
                        sb.AppendLine("{")
                    End If
                End If
                SrcCount += 1

                'sb.Append("CASE" & vbCrLf)
                'sb.Append("(" & vbCrLf)

                'For i = 0 To sDs.ObjSelections.Count - 1
                '    sel = sDs.ObjSelections(i)
                '    taskName = "P_" & sel.Text
                '    sb.Append(vbTab & IIf(i = 0, "", ",") & "EQ(RECNAME(" & sDs.Text & "),'" & sel.Text & "')" & vbCrLf)
                '    sb.Append(vbTab & ",CALLPROC(" & taskName & ")" & vbCrLf)
                '    sb.Append(vbCrLf)
                'Next

                'sb.Append(")" & vbCrLf)

                'sb.AppendLine("{")
                If sDs.IsLookUp = False Then
                    sb.AppendLine("CASE RECNAME(" & sDs.Text & ")")
                    For Each sel In sDs.ObjSelections
                        'count += 1
                        'If count <= objThis.Procs.Count Then
                        '    taskName = CType(objThis.Procs(count), clsTask).TaskName
                        'Else
                        taskName = ""
                        'End If
                        For Each proc As clsTask In objThis.Engine.Procs
                            If proc.TaskName.Contains(sel.Text) = True Then
                                taskName = proc.TaskName
                                Exit For
                            End If
                        Next

                        sb.AppendLine(TAB & "WHEN '" & sel.Text & "'") '& IIf(nd.Index = 0, "", ",")
                        sb.AppendLine(TAB & "DO")
                        sb.AppendLine(TAB & TAB & "CALLPROC(" & taskName & ")")
                        sb.AppendLine(TAB & "END") '& vbCrLf)
                        'sb.Append(vbCrLf)
                    Next
                End If

                'sb.Append(")" & vbCrLf)
                'sb.AppendLine("}")   '& inDsNode.Text
                'sb.AppendLine("FROM " & sDs.DatastoreName)
            End While

            GetMainText = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlEngine GetMainText")
            GetMainText = ""
        End Try

    End Function

    Function GetScriptForProcV3() As String

        Try
            Dim sb As New System.Text.StringBuilder

            'sb.Append("{" & vbCrLf & vbCrLf)
            sb.Append("CASE" & vbCrLf)
            sb.Append(TAB & "WHEN(  )" & vbCrLf)
            sb.Append(TAB & "DO" & vbCrLf & vbCrLf)
            sb.Append(TAB & "END" & vbCrLf)
            'sb.Append(vbCrLf)
            'sb.Append("}")

            GetScriptForProcV3 = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlTask GetScriptForProc")
            Return ""
        End Try

    End Function

    'CREATE PROC P_Lookup AS SELECT
    '{
    '  LOOK(LOOKUP.FILE,'AAA')
    '  IF LOOKFOUND(LOOKUP.TABLE) = TRUE
    '   DO
    '      DESIRE_FIELD =  LOOKFLD(LOOKUP.TABLE,FIELD_NAME)
    '      DESIRE_FIELD =  LOOKFLD(LOOKUP.TABLE,FIELD_NAME)
    '      DESIRE_FIELD =  LOOKFLD(LOOKUP.TABLE,FIELD_NAME)
    '      END
    '}
    'FROM ;
    Function GetScriptForLOOK() As String

        Try
            Dim sb As New System.Text.StringBuilder
            Dim DSlu As clsDatastore = Nothing
            Dim LUsel As clsDSSelection = Nothing

            GetScriptForLOOK = ""

            If objThis.ObjSources.Count <> 1 Then
                MsgBox("You must choose ONE Source Lookup Datastore to use this template", MsgBoxStyle.OkOnly, "Lookup Template")
                Exit Try
            End If

            For Each ds As clsDatastore In objThis.ObjSources
                If ds.IsLookUp = True Then
                    DSlu = ds
                End If
            Next
            If DSlu IsNot Nothing Then
                If DSlu.ObjSelections.Count <> 1 Then
                    MsgBox("You must choose ONE Datastore Selection for a Lookup", MsgBoxStyle.OkOnly, "Lookup Template")
                    Exit Try
                Else
                    For Each sel As clsDSSelection In DSlu.ObjSelections
                        LUsel = sel
                    Next
                End If
            Else
                Exit Try
            End If

            sb.AppendLine("LOOK(" & DSlu.DsPhysicalSource & ",'   ')")
            sb.AppendLine("IF LOOKFOUND(" & DSlu.DatastoreName & ") = TRUE")
            sb.AppendLine("DO")
            For Each fld As clsField In LUsel.DSSelectionFields
                sb.AppendLine(TAB & "    = LOOKFLD(" & DSlu.DatastoreName & "," & fld.FieldName & ")")
            Next
            sb.AppendLine("END")

            GetScriptForLOOK = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlTask GetScriptForLOOK")
            Return ""
        End Try

    End Function

    Function GetSetImageTemplate() As String

        Try
            Dim sb As New System.Text.StringBuilder

            'sb.Append("{" & vbCrLf & vbCrLf)
            sb.AppendLine("--You will need to define a variable field V_IMAGE")
            sb.AppendLine("--You will need to create a Procedure Called P_ROUTE")
            sb.AppendLine("")
            sb.AppendLine("IF CDCOP(CDCIN) = 'R' AND SETIMAGE('BEFORE') = TRUE")
            sb.AppendLine("DO")
            sb.AppendLine(TAB & "V_IMAGE = 'B'")
            sb.AppendLine(TAB & "CALLPROC(P_ROUTE)")
            sb.AppendLine("END")
            sb.AppendLine("")
            sb.AppendLine("IF SETIMAGE('AFTER') = TRUE")
            sb.AppendLine("DO")
            sb.AppendLine(TAB & "V_IMAGE = 'A'")
            sb.AppendLine(TAB & "CALLPROC(P_ROUTE)")
            sb.AppendLine("END")

            GetSetImageTemplate = sb.ToString

            '--You will need to define a variable field V_IMAGE
            '--You will need to create a Procedure Called P_ROUTE

            '  IF  SETIMAGE('BEFORE') = TRUE       
            '  Do
            '      V_IMAGE = 'B'                
            '  CALLPROC(P_ROUTE)
            '  End

            '  IF SETIMAGE('AFTER') = TRUE                    
            '  Do
            '      V_IMAGE = 'A'                            
            '      CALLPROC(P_ROUTE)
            '  End

            'Return True

        Catch ex As Exception
            LogError(ex, "ctlMain GetSetImageTemplate")
            GetSetImageTemplate = ""
        End Try

    End Function

    Function GetScriptForCASE() As String

        Try
            Dim sb As New System.Text.StringBuilder

            sb.AppendLine("--You will need to define a 'When condition'")
            sb.AppendLine("--You will need to define true and false actions")
            sb.AppendLine("")
            sb.AppendLine("CASE")
            sb.AppendLine("WHEN(condition)")
            sb.AppendLine("DO")
            sb.AppendLine("-- true_action")
            sb.AppendLine("END")
            sb.AppendLine("")
            sb.AppendLine("OTHERWISE")
            sb.AppendLine("DO")
            sb.AppendLine("-- false_action")
            sb.AppendLine("END")

            GetScriptForCASE = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlMain GetScriptForCASE")
            Return ""
        End Try

    End Function

    Function GetScriptForCurrentDate() As String

        Try
            GetScriptForCurrentDate = "LEFT(DATETIME(),10)"

        Catch ex As Exception
            LogError(ex, "ctlTask GetScriptForCURRENTDATE")
            Return ""
        End Try

    End Function

    'IF CDCOP(CDCIN) = 'R' DO                                              
    '  MAP_BEFORE(CDCBEFORE(CDCIN.S2PACITP.AP_ID), 'T_S2PACITP.AP_ID')     
    '  MAP_BEFORE(CDCBEFORE(CDCIN.S2PACITP.INIT_DEP_TP_CD),                
    '  'T_S2PACITP.INIT_DEP_TP_CD')                                  
    'END
    Function GetScriptForKeyChange() As String

        Try
            Dim sb As New System.Text.StringBuilder

            sb.AppendLine("--CDCIN is Source Datastore")
            sb.AppendLine("--INDESC is the Source Description containing Key field")
            sb.AppendLine("--OUTDESC is the Target Description containing Key field")
            sb.AppendLine("--KFLDx are the Key fields")
            sb.AppendLine("--'R' represents the 'Replace' Operation")
            sb.AppendLine("")
            sb.AppendLine("IF CDCOP(CDCIN) = 'R' DO")
            sb.AppendLine("      MAP_BEFORE(CDCBEFORE(CDCIN.INDESC.KFLD1), 'OUTDESC.KFLD1')")
            sb.AppendLine("      MAP_BEFORE(CDCBEFORE(CDCIN.INDESC.KFLD2), 'OUTDESC.KFLD2')")
            sb.AppendLine("--           .   .      .")
            sb.AppendLine("--           .   .      .")
            sb.AppendLine("--        Repeat as necessary")
            sb.AppendLine("END")


            GetScriptForKeyChange = sb.ToString

        Catch ex As Exception
            LogError(ex, "ctlMain GetScriptForKeyChange")
            Return ""
        End Try

    End Function

    '//Switch to code edit view when user double click or select edit menu item
    Function ShowScriptEditor(ByVal objMap As clsMapping) As Boolean

        Try
            objCurMap = objMap
            IsEventFromCode = True
            txtCodeEditor.Text = ""

            If curEditType = modDeclares.enumDirection.DI_SOURCE Then
                If objCurMap.MappingSource IsNot Nothing Then
                    If objCurMap.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                        ObjCurFld = CType(objCurMap.MappingSource, clsField)
                    End If
                    If objCurMap.SourceType = modDeclares.enumMappingType.MAPPING_TYPE_FUN Then
                        txtCodeEditor.Text = CType(objCurMap.MappingSource, clsSQFunction).SQFunctionWithInnerText
                    Else
                        txtCodeEditor.Text = objCurMap.MappingSource.Text
                    End If
                Else
                    txtCodeEditor.Text = ""
                End If
                'If objCurMap.MappingTarget IsNot Nothing Then
                '    If objCurMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
                '        lblTgtFldName.Text = CType(objCurMap.MappingTarget, clsField).FieldName
                '    Else
                '        lblTgtFldName.Text = "No Target Field Mapped"
                '    End If
                'End If
            Else
                If objCurMap.MappingTarget IsNot Nothing Then
                    If objCurMap.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
                        ObjCurFld = CType(objCurMap.MappingTarget, clsField)
                    End If
                    If objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_FIELD Or _
                        objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_NONE Then
                        txtCodeEditor.Text = objCurMap.MappingTarget.Text
                    ElseIf objCurMap.TargetType = modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR Or _
                    objCurMap.TargetType = enumMappingType.MAPPING_TYPE_VAR Then
                        If CType(objCurMap.MappingTarget, clsVariable).CorrectedVariableName = "" Then
                            txtCodeEditor.Text = objCurMap.MappingTarget.Text
                        Else
                            txtCodeEditor.Text = CType(objCurMap.MappingTarget, clsVariable).CorrectedVariableName
                        End If
                    End If
                Else
                    txtCodeEditor.Text = ""
                End If
            End If

            IsEventFromCode = False
            pnlScript.Visible = True
            'pnlSourceTarget.Visible = False
            'pnlDesc.Visible = False
            pnlScript.BringToFront()

            IsCodeEditorOnTop = True
            IsDescEditorOnTop = False
            'txtCodeEditor.Focus()

            Return True

        Catch ex As Exception
            LogError(ex, "ctlTask ShowScriptEditor")
            Return False
        End Try

    End Function

    'Function ShowDescEditor(ByVal objMap As clsMapping) As Boolean

    '    Try
    '        objCurMap = objMap
    '        IsEventFromCode = True

    '        'txtMapDesc.Text = Strings.Replace(objCurMap.MappingDesc, "''", "'")

    '        IsEventFromCode = False
    '        pnlScript.Visible = False
    '        'pnlSourceTarget.Visible = False
    '        'pnlDesc.Visible = True
    '        'pnlDesc.BringToFront()
    '        IsCodeEditorOnTop = False
    '        IsDescEditorOnTop = True
    '        'txtMapDesc.Focus()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask ShowDescEditor")
    '        Return False
    '    End Try

    'End Function

    'Function SetLastMapped() As Boolean

    '    Dim i As Integer
    '    Dim map As clsMapping
    '    Dim srcfld As clsField = Nothing
    '    Dim tgtfld As clsField = Nothing

    '    Try
    '        For i = 0 To objThis.ObjMappings.Count - 1
    '            map = CType(objThis.ObjMappings(i), INode)
    '            If map.MappingSource IsNot Nothing Then
    '                If CType(map.MappingSource, INode).Type = NODE_STRUCT_FLD Then
    '                    If CType(map.MappingSource, clsField).Parent IsNot Nothing Then
    '                        srcfld = CType(map.MappingSource, clsField)
    '                    End If
    '                End If
    '            End If
    '            If map.MappingTarget IsNot Nothing Then
    '                If CType(map.MappingTarget, INode).Type = NODE_STRUCT_FLD Then
    '                    If CType(map.MappingTarget, clsField).Parent IsNot Nothing Then
    '                        tgtfld = CType(map.MappingTarget, clsField)
    '                    End If
    '                End If
    '            End If
    '        Next
    '        If srcfld IsNot Nothing Then
    '            objThis.LastSrcFld = srcfld.ParentName & "." & srcfld.FieldName
    '        End If
    '        If tgtfld IsNot Nothing Then
    '            objThis.LastTgtFld = tgtfld.ParentName & "." & tgtfld.FieldName
    '        End If

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "clsTask SetLastMapped")
    '        Return False
    '    End Try

    'End Function

#End Region

    'Private Sub mnuOpenStrFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuOpenStrFile.Click

    '    Try
    '        If CurRow >= 0 Then
    '            Dim objMap As clsMapping = lvMappings.Items(CurRow).Tag
    '            Dim fld As clsField

    '            If CurCol = 0 Then
    '                If CType(objMap.MappingSource, INode).Type = NODE_STRUCT_FLD Then
    '                    fld = CType(objMap.MappingSource, clsField)
    '                    fld.Struct.showStrFile()
    '                End If
    '            Else
    '                If CType(objMap.MappingTarget, INode).Type = NODE_STRUCT_FLD Then
    '                    fld = CType(objMap.MappingTarget, clsField)
    '                    fld.Struct.showStrFile()
    '                End If
    '            End If
    '        End If

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask mnuOpenStrFile_Click")
    '    End Try

    'End Sub

    'Private Sub OpenStructureFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenStructureFileToolStripMenuItem.Click

    '    Try
    '        Dim str As clsStructure = Nothing
    '        Dim Node As TreeNode = tvSource.GetNodeAt(MousePos)
    '        Dim NodeType As String = CType(Node.Tag, INode).Type

    '        Select Case NodeType
    '            Case NODE_SOURCEDSSEL
    '                str = CType(Node.Tag, clsDSSelection).ObjStructure
    '            Case NODE_STRUCT_FLD
    '                str = CType(Node.Tag, clsField).Struct
    '            Case NODE_STRUCT_SEL
    '                str = CType(Node.Tag, clsStructureSelection).ObjStructure
    '            Case Else
    '                Exit Try
    '        End Select
    '        str.showStrFile()

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask OpenStrFileSource")
    '    End Try

    'End Sub

    'Private Sub OpenStructureFileToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenStructureFileToolStripMenuItem1.Click

    '    Try
    '        Dim str As clsStructure = Nothing
    '        Dim Node As TreeNode = tvTarget.GetNodeAt(MousePos)
    '        Dim NodeType As String = CType(Node.Tag, INode).Type

    '        Select Case NodeType
    '            Case NODE_TARGETDSSEL
    '                str = CType(Node.Tag, clsDSSelection).ObjStructure
    '            Case NODE_STRUCT_FLD
    '                str = CType(Node.Tag, clsField).Struct
    '            Case NODE_STRUCT_SEL
    '                str = CType(Node.Tag, clsStructureSelection).ObjStructure
    '            Case Else
    '                Exit Try
    '        End Select
    '        str.showStrFile()

    '    Catch ex As Exception
    '        LogError(ex, "ctlTask OpenStrFileTarget")
    '    End Try

    'End Sub

    Private Sub txtCodeEditor_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCodeEditor.TextChanged
        Try
            Dim ColNum As Integer
            'Dim idx As Integer
            ''Dim StartIdx As Integer
            'Dim SubStr As String

            objCurFunct.SQFunctionWithInnerText = txtCodeEditor.Text

            ''StartIdx = txtCodeEditor.GetFirstCharIndexOfCurrentLine()
            'idx = txtCodeEditor.GetCharIndexFromPosition(Windows.Forms.Cursor.Position)


            'SubStr = txtCodeEditor.Text.Substring(0, idx)
            'ColNum = SubStr.Length - SubStr.LastIndexOf(Chr(13))

            Dim lineindex As Integer
            lineindex = Me.txtCodeEditor.GetLineFromCharIndex(Me.txtCodeEditor.SelectionStart)

            Dim columnindex As Integer
            columnindex = Me.txtCodeEditor.SelectionStart - Me.txtCodeEditor.GetFirstCharIndexFromLine(lineindex)

            If Me.txtCodeEditor.GetFirstCharIndexFromLine(lineindex + 1) > 1 Then
                ColNum = Me.txtCodeEditor.GetFirstCharIndexFromLine(lineindex + 1) - Me.txtCodeEditor.GetFirstCharIndexFromLine(lineindex)
            Else
                ColNum = Me.txtCodeEditor.Text.Length - Me.txtCodeEditor.GetFirstCharIndexFromLine(lineindex)
            End If

            txtColNum.Text = ColNum.ToString

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "ctlTask txtCodeEditorTextChng")
        End Try

    End Sub

    Private Sub cbIsMain_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbIsMain.CheckedChanged

        Try
            If IsEventFromCode = True Then Exit Try

            If cbIsMain.Checked = True Then
                'objCurMap.IsMapped = "1"
                objThis.TaskType = enumTaskType.TASK_MAIN
                'RemoveFromCollection(objThis.Engine.Procs, objThis.GUID)
                'AddToCollection(objThis.Engine.Mains, objThis, objThis.GUID)
            Else
                'objCurMap.IsMapped = "0"
                objThis.TaskType = enumTaskType.TASK_GEN
                'RemoveFromCollection(objThis.Engine.Mains, objThis.GUID)
                'AddToCollection(objThis.Engine.Procs, objThis, objThis.GUID)
            End If

            DoSetTaskType = True

            OnChange(Me, New EventArgs)

        Catch ex As Exception
            LogError(ex, "cbIsMain_CheckedChanged")
        End Try

    End Sub

End Class
