Public Class frmRplcDescRet

    Private NewStr As clsStructure
    Private OldStr As clsStructure
    Private AddList As ArrayList
    Private DelList As ArrayList
    Dim Procs As Collection
    'Dim Engs As Collection

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Change_Desc)

    End Sub

    Private Sub me_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.F1 Then
            ShowHelp(modHelp.HHId.H_Change_Desc)
        End If

    End Sub

    Sub StartLoad()

        IsEventFromCode = True

        txtName.Text = OldStr.StructureName
        txtOldPath.Text = OldStr.fPath1
        txtNewPath.Text = NewStr.fPath1

        LoadFields()
        LoadProcs()

    End Sub

    Sub LoadFields()

        Try
            Dim i, j As Integer
            Dim addtxt As New System.Text.StringBuilder
            Dim deltxt As New System.Text.StringBuilder


            For i = 0 To AddList.Count - 1
                addtxt.AppendLine(CType(AddList(i), clsField).FieldName)
            Next
            txtNewFlds.Text = addtxt.ToString

            For j = 0 To DelList.Count - 1
                deltxt.AppendLine(CType(DelList(j), clsField).FieldName)
            Next
            txtOldFlds.Text = deltxt.ToString

        Catch ex As Exception
            LogError(ex, "frmRplcDescRet LoadFields")
        End Try

    End Sub

    Sub LoadProcs()

        Try
            Dim Obj As Object
            Dim EnvObj As clsEnvironment = OldStr.Environment

            For Each sys As clsSystem In EnvObj.Systems
                For Each eng As clsEngine In sys.Engines
                    For Each proc As clsTask In eng.Procs
                        proc.LoadItems()

                        For Each map As clsMapping In proc.ObjMappings
                            Obj = map.MappingSource
                            If Obj IsNot Nothing Then
                                If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                    Dim fld As clsField = CType(Obj, clsField)
                                    If fld.ParentStructureName = OldStr.StructureName Then
                                        lvProc.Items.Add(eng.EngineName).SubItems.Add(proc.TaskName)
                                        'If Engs.Contains(eng.Key) = False Then
                                        '    Engs.Add(eng, eng.Key)
                                        'End If
                                        If Procs.Contains(proc.TaskName) = False Then
                                            Procs.Add(proc, proc.TaskName)
                                        End If
                                        GoTo NextProc
                                    End If
                                End If
                            End If
                            Obj = map.MappingTarget
                            If Obj IsNot Nothing Then
                                If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                    Dim fld As clsField = CType(Obj, clsField)
                                    If fld.ParentStructureName = OldStr.StructureName Then
                                        lvProc.Items.Add(eng.EngineName).SubItems.Add(proc.TaskName)
                                        'If Engs.Contains(eng.Key) = False Then
                                        '    Engs.Add(eng, eng.Key)
                                        'End If
                                        If Procs.Contains(proc.TaskName) = False Then
                                            Procs.Add(proc, proc.TaskName)
                                        End If
                                        GoTo NextProc
                                    End If
                                End If
                            End If
                        Next
NextProc:           Next
                Next
            Next

        Catch ex As Exception
            LogError(ex, "frmRplcDescRet LoadProcs")
        End Try

    End Sub

    Sub EndLoad()

        IsEventFromCode = False

    End Sub

    Public Function ShowFldDiff(ByVal NewObj As clsStructure, ByVal OldObj As clsStructure, ByVal newList As ArrayList, ByVal oldList As ArrayList, ByRef procedures As Collection) As Boolean

        Try
            NewStr = NewObj
            OldStr = OldObj
            AddList = newList
            DelList = oldList
            Procs = procedures
            StartLoad()
            EndLoad()

doAgain:
            Select Case Me.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    ShowFldDiff = True
                Case Windows.Forms.DialogResult.Retry
                    GoTo doAgain
                Case Else
                    Return False
            End Select

        Catch ex As Exception
            LogError(ex, "frmRplcDescRet ShowFldDiff")
        End Try

    End Function

End Class
