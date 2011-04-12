Public Class ctlInclude
    Inherits System.Windows.Forms.UserControl

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)
    Public Event Saved(ByVal sender As System.Object, ByVal e As INode)
    Public Event Renamed(ByVal sender As System.Object, ByVal e As INode)
    Public Event Closed(ByVal sender As System.Object, ByVal e As INode)

    Dim objThis As INode
    Dim IsNewObj As Boolean

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFile.TextChanged, txtName.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        cmdSave.Enabled = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Save()

    End Sub

    Private Sub cmdClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClose.Click

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

    End Sub

    Private Sub txtFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFile.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        objThis.IsLoaded = False
        RaiseEvent Modified(Me, objThis)

        If txtName.Text.Trim = "" Then
            txtName.Text = GetFileNameWithoutExtenstionFromPath(txtFile.Text)
        End If

    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click

        Dim strFilter As String = "SQD Script Files (*.sqd)|*.sqd|All files (*.*)|*.*"

        dlgOpen.Filter = strFilter

        If dlgOpen.ShowDialog() = DialogResult.OK Then
            txtFile.Text = dlgOpen.FileName
        End If

    End Sub

    Public Function Save() As Boolean
        Try
            Me.Cursor = Cursors.WaitCursor
            '// First Check Validity before Saving
            If ValidateNewName(txtName.Text) = False Then
                Me.Cursor = Cursors.Default
                Exit Function
            End If
            Select Case objThis.Type
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    If CType(objThis, clsDatastore).DatastoreName <> txtName.Text Then
                        objThis.IsRenamed = RenameDatastore(objThis, txtName.Text)
                    End If
                    If objThis.IsRenamed = False Then
                        txtName.Text = CType(objThis, clsDatastore).DatastoreName
                    Else
                        CType(objThis, clsDatastore).DatastoreName = txtName.Text
                    End If

                    CType(objThis, clsDatastore).DsPhysicalSource = Me.txtFile.Text
                    CType(objThis, clsDatastore).DatastoreName = Me.txtName.Text

                Case NODE_PROC
                    If CType(objThis, clsTask).TaskName <> txtName.Text Then
                        objThis.IsRenamed = RenameTask(objThis, txtName.Text)
                    End If
                    If objThis.IsRenamed = False Then
                        txtName.Text = CType(objThis, clsTask).TaskName
                    Else
                        CType(objThis, clsTask).TaskName = txtName.Text
                    End If

                    CType(objThis, clsTask).TaskDescription = Me.txtFile.Text
                    CType(objThis, clsTask).TaskName = Me.txtName.Text

            End Select

            If IsNewObj = True Then
                If objThis.AddNew = False Then
                    Me.Cursor = Cursors.Default
                    Exit Function
                End If
            Else
                If objThis.Save() = False Then
                    Me.Cursor = Cursors.Default
                    Exit Function
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
            LogError(ex, "ctlInclude Save")
            Me.Cursor = Cursors.Default
        End Try

    End Function

    Public Function EditObj(ByVal obj As INode) As INode

        Try
            Me.Cursor = Cursors.WaitCursor
            IsNewObj = False

            objThis = obj '//Load the form object

            StartLoad()

            Select Case objThis.Type
                Case NODE_SOURCEDATASTORE
                    objThis = CType(objThis, clsDatastore)
                    Me.txtFile.Text = CType(objThis, clsDatastore).DsPhysicalSource
                    Me.txtName.Text = CType(objThis, clsDatastore).DatastoreName
                    Me.gbInclude.Text = "Include Source Datastore"

                Case NODE_TARGETDATASTORE
                    objThis = CType(objThis, clsDatastore)
                    Me.txtFile.Text = CType(objThis, clsDatastore).DsPhysicalSource
                    Me.txtName.Text = CType(objThis, clsDatastore).DatastoreName
                    Me.gbInclude.Text = "Include Target Datastore"

                Case NODE_PROC
                    objThis = CType(objThis, clsTask)
                    Me.txtFile.Text = CType(objThis, clsTask).TaskDescription
                    Me.txtName.Text = CType(objThis, clsTask).TaskName
                    Me.gbInclude.Text = "Include Procedure"

            End Select

            EndLoad()
            Me.Cursor = Cursors.Default
            EditObj = objThis

        Catch ex As Exception
            LogError(ex, "ctlInclude EditObj")
            Me.Cursor = Cursors.Default
            EditObj = Nothing
        End Try

    End Function

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False
        cmdSave.Enabled = False
       
        'ClearControls(Me.Controls)
        
    End Sub

    Private Sub EndLoad()

        Me.BringToFront()
        Me.Visible = True
        IsEventFromCode = False

    End Sub

End Class