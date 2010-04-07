Public Class frmInclude

    Public objThis As INode
    Dim IsNewObj As Boolean
    Dim type As String = ""

    Public Event Modified(ByVal sender As System.Object, ByVal e As INode)

    Private Sub OnChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
        RaiseEvent Modified(Me, objThis)

    End Sub

    '//Returns new Include object
    Public Function NewObj(ByVal InodeType As String, ByVal parent As INode) As INode

        IsNewObj = True
        type = InodeType

        Select Case type
            Case NODE_SOURCEDATASTORE
                objThis = New clsDatastore
                objThis.Parent = parent
                gbInclude.Text = "Include Source Datastore"

            Case NODE_TARGETDATASTORE
                objThis = New clsDatastore
                objThis.Parent = parent
                gbInclude.Text = "Include Target Datastore"

            Case NODE_PROC
                objThis = New clsTask
                objThis.Parent = parent
                gbInclude.Text = "Include Procedure"

            Case Else
                objThis = New clsDatastore
                objThis.Parent = parent
                gbInclude.Text = "Include Datastore"

        End Select

        StartLoad()

        txtName.SelectAll()

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

    Private Sub StartLoad()

        IsEventFromCode = True
        objThis.IsModified = False

    End Sub

    Private Sub EndLoad()

        IsEventFromCode = False

    End Sub

    Private Sub txtFile_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFile.TextChanged

        If IsEventFromCode = True Then Exit Sub
        objThis.IsModified = True
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

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Try
            If objThis.ValidateNewObject(txtName.Text) = False Or ValidateNewName(txtName.Text) = False Then
                DialogResult = Windows.Forms.DialogResult.Retry
                Exit Sub
            End If

            Select Case type
                Case NODE_SOURCEDATASTORE
                    CType(objThis, clsDatastore).DsDirection = DS_DIRECTION_SOURCE
                    CType(objThis, clsDatastore).DatastoreType = enumDatastore.DS_INCLUDE
                    CType(objThis, clsDatastore).DatastoreName = txtName.Text
                    CType(objThis, clsDatastore).DsPhysicalSource = txtFile.Text

                Case NODE_TARGETDATASTORE
                    CType(objThis, clsDatastore).DsDirection = DS_DIRECTION_SOURCE
                    CType(objThis, clsDatastore).DatastoreType = enumDatastore.DS_INCLUDE
                    CType(objThis, clsDatastore).DatastoreName = txtName.Text
                    CType(objThis, clsDatastore).DsPhysicalSource = txtFile.Text

                Case NODE_PROC
                    CType(objThis, clsTask).TaskType = enumTaskType.TASK_IncProc
                    CType(objThis, clsTask).TaskName = txtName.Text
                    CType(objThis, clsTask).TaskDescription = txtFile.Text

            End Select

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
            LogError(ex)
            DialogResult = Windows.Forms.DialogResult.Retry
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

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

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

    End Sub

End Class
