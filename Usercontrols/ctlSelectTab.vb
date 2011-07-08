Public Class ctlSelectTab

    Dim dt2 As New System.Data.DataTable("temp2") '/// Table of columns for chosen sys table
    Public objthis As New clsDMLinfo

#Region "Toolbar and Context Menu Buttons"

    '/// these call mnu MoveUp, Movedown, and mnuRemove Subs
    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click

        mnuMoveUp_Click(sender, New EventArgs)

    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click

        mnuRemove_Click(sender, New EventArgs)

    End Sub

    Private Sub btnDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDown.Click

        mnuMoveDown_Click(sender, New EventArgs)

    End Sub
    '// remove column from "selected" list
    Private Sub mnuRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuRemove.Click

        If ListView1.FocusedItem IsNot Nothing Then
            ListView1.FocusedItem.Remove()
        End If
        'If ListView1.Items.Count > 0 Then
        '    cmdOk.Enabled = True
        'Else
        '    cmdOk.Enabled = False
        'End If

    End Sub

    '/// next two functions are to change the order of chosen columns
    Private Sub mnuMoveUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMoveUp.Click

        If ListView1.FocusedItem IsNot Nothing Then
            SetNewPosition(ListView1.SelectedItems(0).Index - 1)
        End If

    End Sub

    Private Sub mnuMoveDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuMoveDown.Click

        If ListView1.FocusedItem IsNot Nothing Then
            SetNewPosition(ListView1.SelectedItems(0).Index + 1)
        End If

    End Sub

    '// Removes all Columns
    Private Sub SelectAllColumnsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllColumnsToolStripMenuItem.Click

        Try
            ListView1.Items.Clear()

        Catch ex As Exception
            LogError(ex, "frmChooseCol RemoveAllColumnsInListView")
        End Try

    End Sub

    '/// Adds all Columns
    Private Sub SelectAllColumnsToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllColumnsToolStripMenuItem1.Click

        Dim rownum As Integer = 0
        Dim i As Integer
        Dim column As String
        Dim colType As String
        Dim colLen As String
        Dim colScale As String

        Try
            ListView1.SmallImageList = ImageList1

            For i = 0 To dgvColumns.RowCount - 1
                column = dgvColumns.Item(0, i).Value.ToString
                colType = dgvColumns.Item(1, i).Value.ToString
                colLen = dgvColumns.Item(2, i).Value.ToString
                colScale = dgvColumns.Item(3, i).Value.ToString

                If ListView1.Items.ContainsKey(column) = False Then
                    ListView1.Items.Add(column, column, 0).SubItems.Add(colType)
                    ListView1.Items(column).SubItems.Add(colLen)
                    ListView1.Items(column).SubItems.Add(colScale)
                End If
            Next
            
        Catch ex As Exception
            LogError(ex, "frmChooseCol DGV1_SelectionChg")
        End Try

    End Sub

    '/// Removes all Columns
    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click

        SelectAllColumnsToolStripMenuItem_Click(sender, New EventArgs)

    End Sub

    '// Adds All Columns
    Private Sub btnAddAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddAll.Click

        SelectAllColumnsToolStripMenuItem1_Click(sender, New EventArgs)

    End Sub

#End Region

    Function SaveColumns() As Boolean

        Try
            For Each lvItem As ListViewItem In ListView1.Items
                objthis.ColumnArray.Add(lvItem.Name)
            Next
            Return True

        Catch ex As Exception
            LogError(ex, "frmDMLinfo SaveColumns")
            Return False
        End Try

    End Function

    '///Clears the "Select Columns" tab
    Sub clearTabColumns()

        dt2.Clear()
        ListView1.Items.Clear()

    End Sub

    '///Function copied from ctlTask and modified by TKarasch 6/14/2007
    Private Function SetNewPosition(ByVal NewPosition As Integer) As Boolean

        Try
            If ListView1.SelectedItems.Count > 0 Then
                Dim OldPosition As Integer

                '//How many items to move (up or down). All items will be shifted by this offset
                Dim moveOffset As Integer
                Dim newIndex As Integer

                OldPosition = ListView1.SelectedItems(0).Index

                '//if offset is +ve then shift all selected item down by moveOffset
                '//if offset is -ve then shift all selected item up by moveOffset

                moveOffset = NewPosition - OldPosition

                If moveOffset = 0 Then Exit Function

                '//Any item below the current item must be updated coz SeqNo has been changed
                '//since we changed the positions we have to update IsModified flag for some items
                Dim startIndex, endIndex As Integer
                If moveOffset > 0 Then '// +ve means moved down
                    '//old position of first item in selection list
                    startIndex = ListView1.SelectedItems(0).Index
                    '//new position of last item in selection list
                    endIndex = ListView1.SelectedItems(ListView1.SelectedItems.Count - 1).Index + moveOffset
                ElseIf moveOffset < 0 Then '// -ve means moved up
                    '//new position of first item in selection list
                    startIndex = ListView1.SelectedItems(0).Index + moveOffset
                    '//old position of last item in the selection list
                    endIndex = ListView1.SelectedItems(ListView1.SelectedItems.Count - 1).Index
                End If

                '//Check for valid new position
                If (endIndex > (ListView1.Items.Count - 1)) Then Exit Function
                If (startIndex < 0) Or (startIndex > (ListView1.Items.Count - 1)) Then Exit Function


                Dim tempItem As ListViewItem
                Dim curIndex As Integer
                Dim i As Integer

                ListView1.BeginUpdate()

                '//If move down then start repositioning with last item in selection list 
                '//If move up then start repositioning with first item in selection list 
                '//Reposition selected items

                If moveOffset > 0 Then '//if move down
                    For i = ListView1.SelectedItems.Count - 1 To 0 Step -1
                        curIndex = ListView1.SelectedItems(i).Index
                        newIndex = curIndex + moveOffset
                        If newIndex >= 0 Then
                            tempItem = ListView1.SelectedItems(i)
                            ListView1.Items.RemoveAt(curIndex)
                            AddlvItem(tempItem, newIndex)
                            ListView1.Items(newIndex).Selected = True
                            ListView1.Items(newIndex).Focused = True
                        End If
                    Next
                ElseIf moveOffset < 0 Then '//if move up
                    For i = 0 To ListView1.SelectedItems.Count - 1
                        curIndex = ListView1.SelectedItems(i).Index
                        newIndex = curIndex + moveOffset
                        If newIndex >= 0 Then
                            tempItem = ListView1.SelectedItems(i)
                            ListView1.Items.RemoveAt(curIndex)
                            AddlvItem(tempItem, newIndex)
                            ListView1.Items(newIndex).Selected = True
                            ListView1.Items(newIndex).Focused = True
                        End If
                    Next
                End If

                SetNewPosition = True
            Else
                MsgBox("Please select an item from the list", MsgBoxStyle.Exclamation, MsgTitle)
            End If

        Catch ex As Exception
            LogError(ex, "frmDMLinfo SetMappingPosition")
            SetNewPosition = False
        Finally
            ListView1.EndUpdate()
        End Try

    End Function

    '///Function copied from ctlTask and modified by TKarasch 6/14/2007
    Function AddlvItem(ByVal tempItem As ListViewItem, Optional ByVal Pos As Integer = -1) As Boolean

        Try
            Dim lvItm As New ListViewItem
            lvItm = tempItem

            If Pos > ListView1.Items.Count Then
                Pos = ListView1.Items.Count
            End If

            If Pos < 0 Then
                '//add new item at the end
                ListView1.Items.Add(lvItm)
            Else
                '//add new item above the selected item
                ListView1.Items.Insert(Pos, lvItm)
            End If

            lvItm.EnsureVisible()

            AddlvItem = True

        Catch ex As Exception
            LogError(ex, "frmDMLinfo AddlvItem")
            AddlvItem = False
        End Try

    End Function

    Private Sub dgvColumns_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvColumns.CellClick

        Dim rownum As Integer = 0
        Dim column As String
        Dim colType As String
        Dim colLen As String
        Dim colScale As String

        Try


            rownum = dgvColumns.CurrentCell.RowIndex
            column = dgvColumns.Item(0, rownum).Value.ToString
            colType = dgvColumns.Item(1, rownum).Value.ToString
            colLen = dgvColumns.Item(2, rownum).Value.ToString
            colScale = dgvColumns.Item(3, rownum).Value.ToString

            If ListView1.Items.ContainsKey(column) = False Then
                ListView1.Items.Add(column, column, 0).SubItems.Add(colType)
                ListView1.Items(column).SubItems.Add(colLen)
                ListView1.Items(column).SubItems.Add(colScale)
            End If
            'If ListView1.Items.Count > 0 Then
            '    cmdOk.Enabled = True
            'Else
            '    cmdOk.Enabled = False
            'End If


        Catch ex As Exception
            LogError(ex, "frmChooseCol DGV1_SelectionChg")
        End Try

    End Sub

    Private Sub dgvColumns_CellContentDblClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvColumns.CellDoubleClick

        dgvColumns_CellContentClick(sender, e)

    End Sub

    '/// Loads table of Columns from table chosen from sys catalog
    Function Load_ColumnTable() As Boolean

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim conn As New System.Data.Odbc.OdbcConnection(objthis.MetaConnectionString)
        Dim sql As String = ""

        Try
            Me.Cursor = Cursors.WaitCursor
            dt2.Clear()
            dt2.Columns.Clear()

            conn.Open()

            Select Case objthis.DSNtype

                Case enumODBCtype.DB2
                    sql = "SELECT colname, typename, length, scale FROM syscat.columns WHERE tabschema = '" & objthis.Schema & "' AND tabname = '" & objthis.TableName & "' order by colno"

                Case enumODBCtype.ORACLE
                    sql = "select column_name, data_type, data_length, data_scale FROM user_tab_columns WHERE table_name = '" & objthis.TableName & "' order by column_id"

                Case enumODBCtype.SQL_SERVER
                    sql = "select column_name, Data_Type, character_maximum_length, numeric_scale from INFORMATION_SCHEMA.COLUMNS where table_name = '" & objthis.TableName & "' order by ordinal_position"

            End Select

            da = New System.Data.Odbc.OdbcDataAdapter(sql, conn)

            If Not da.Fill(dt2) > 0 Then
                dgvColumns.DataSource = dt2
                dgvColumns.Visible = True
                dgvColumns.Enabled = True
                dgvColumns.Show()
                dgvColumns.ClearSelection()
                MsgBox("The table you have chosen contains no columns" & (Chr(13)) & "Please cancel and select a different Table", MsgBoxStyle.Information, MsgTitle)
                dgvColumns.Enabled = False
            Else
                dgvColumns.DataSource = dt2
                dgvColumns.Visible = True
                dgvColumns.Enabled = True
                dgvColumns.Show()
                dgvColumns.ClearSelection()
            End If

            Try
                If ListView1.Items.Count = 0 Then
                    '/// Now Select all columns
                    Dim rownum As Integer = 0
                    Dim i As Integer
                    Dim column As String
                    Dim colType As String
                    Dim colLen As String
                    Dim colScale As String


                    ListView1.SmallImageList = ImageList1

                    For i = 0 To dgvColumns.RowCount - 1
                        column = dgvColumns.Item(0, i).Value.ToString
                        colType = dgvColumns.Item(1, i).Value.ToString
                        colLen = dgvColumns.Item(2, i).Value.ToString
                        colScale = dgvColumns.Item(3, i).Value.ToString

                        If ListView1.Items.ContainsKey(column) = False Then
                            ListView1.Items.Add(column, column, 0).SubItems.Add(colType)
                            ListView1.Items(column).SubItems.Add(colLen)
                            ListView1.Items(column).SubItems.Add(colScale)
                        End If
                    Next
                End If
                Me.Cursor = Cursors.Default

                Return True

            Catch ex As Exception
                LogError(ex, "frmChooseCol DGV1_SelectionChg")
                Me.Cursor = Cursors.Default
                Return False
            End Try

        Catch OE As Odbc.OdbcException
            Me.Cursor = Cursors.Default
            LogODBCError(OE, "frmNewProj ValidateDatabase")
            MsgBox("An ODBC exception error occured: " & Chr(13) & _
                   OE.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the ODBC Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            Return False

        Catch ex As Exception
            Me.Cursor = Cursors.Default
            If ex.ToString.Contains("Oracle") = True Then
                If ex.ToString.Contains("null password") = True Then
                    MsgBox("Password cannot be left blank", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("password") = True Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.", MsgBoxStyle.Information, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.Information, MsgTitle)
                End If
            ElseIf ex.ToString.Contains("IBM") = True Then
                If ex.ToString.Contains("PASSWORD MISSING") = True Then
                    MsgBox("You have entered an incorrect Password.", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("24") = True Then
                    MsgBox("You have entered an incorrect Username and/or Password.", MsgBoxStyle.Information, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.Information, MsgTitle)
                End If
            ElseIf ex.ToString.Contains("MS") = True Then
                If ex.ToString.Contains("PASSWORD") Then
                    MsgBox("You have entered an incorrect Password.", MsgBoxStyle.Information, MsgTitle)
                ElseIf ex.ToString.Contains("connection") Then
                    MsgBox("You have entered an incorrect Table Schema or an incorrect Username and/or Password.", MsgBoxStyle.Information, MsgTitle)
                Else
                    MsgBox("You have chosen an invalid ODBC source," & Chr(13) & "entered an incorrect Table Schema or an incorrect User name and Password.", MsgBoxStyle.Information, MsgTitle)
                End If
            Else
                MsgBox("A Windows exception error occured: " & Chr(13) & _
                   ex.Message.ToString & Chr(13) & Chr(13) & _
                   "For more information, see the Error Log" & Chr(13) & _
                   "in Main Program Window", MsgBoxStyle.OkOnly, MsgTitle)
            End If
            Me.Cursor = Cursors.Default

            Return False

        Finally
            conn.Close()
            Me.Cursor = Cursors.Default
        End Try

    End Function

    
End Class
