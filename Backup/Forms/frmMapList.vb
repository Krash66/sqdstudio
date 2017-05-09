Public Class frmMapList

    Dim IsEventFromCode As Boolean

    Dim CurMapObj As clsMapPattern
    Dim TempMapObj As clsMapPattern
    Dim ProjObj As clsProject

    Dim SrcLoaded As Boolean = False
    Dim TgtLoaded As Boolean = False

    Dim Filepath As String
    Dim PrevRowIdx As Integer

#Region "Buttons"

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        If SaveMapList() = True Then
            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK
        Else
            DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    Private Sub cmdSaveAs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveAs.Click

        If SaveMapList(True) = True Then
            Me.Close()
            DialogResult = Windows.Forms.DialogResult.OK
        Else
            DialogResult = Windows.Forms.DialogResult.Retry
        End If

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()
        DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(HHId.H_MapListTool)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click_1(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click_1(sender, New EventArgs)
        End Select

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Try
            IsEventFromCode = True

            Dim NewMap As New clsMapPattern
            Dim NewRow As New DataGridViewRow
            Dim NewArr As String()
            Dim idx As Integer

            NewMap.Source = txtAddSource.Text
            NewMap.Target = txtAddTarget.Text

            For Each map As clsMapPattern In ProjObj.Maplist
                If (NewMap.Source = map.Source And NewMap.Source <> "") Or (NewMap.Target = map.Target And NewMap.Target <> "") Then
                    Exit Try
                End If
            Next

            NewMap.Project = ProjObj

            NewArr = Strings.Split(NewMap.ObjString, ",")

            ProjObj.Maplist.Add(NewMap)

            idx = DataGridView1.Rows.Add(NewArr)
            DataGridView1.Rows(idx).Tag = NewMap
            DataGridView1.CurrentCell = DataGridView1.Item(0, idx)

            txtAddSource.Text = ""
            txtAddTarget.Text = ""

            SrcLoaded = False
            TgtLoaded = False

            IsEventFromCode = False

        Catch ex As Exception
            LogError(ex, "frmMapList btnAddClick")
        End Try

    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click

        Try
            Dim idx As Integer = DataGridView1.CurrentRow.Index
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
            ProjObj.Maplist.RemoveAt(idx)

            txtAddSource.Text = ""
            txtAddTarget.Text = ""

            SrcLoaded = False
            TgtLoaded = False

        Catch ex As Exception
            LogError(ex, "btnRemove_click frmMapList")
        End Try

    End Sub

    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click

        Try
            Dim NewMap As New clsMapPattern
            Dim curRow As Integer = DataGridView1.CurrentRow.Index
            Dim NewArr As String()

            NewMap.Source = CurMapObj.Source
            NewMap.Target = CurMapObj.Target
            NewMap.Project = ProjObj

            NewArr = Strings.Split(NewMap.ObjString, ",")

            ProjObj.Maplist.Item(curRow) = NewMap

            DataGridView1.CurrentRow.SetValues(NewArr)
            DataGridView1.CurrentRow.Tag = NewMap

            txtAddSource.Text = ""
            txtAddTarget.Text = ""

            SrcLoaded = False
            TgtLoaded = False

        Catch ex As Exception
            LogError(ex, "frmMapList btnUpdateClick")
        End Try

    End Sub

    Private Sub btnAddFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFile.Click

        Dim strArr As String()
        Dim i As Integer

        Try
            If OpenFD1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                If OpenFD1.FileName <> "" Then
                    Filepath = OpenFD1.FileName
                Else
                    Exit Try
                End If
            Else
                Exit Try
            End If

            txtAddFile.Text = Filepath

            Dim OFS As System.IO.FileStream = CType(OpenFD1.OpenFile, System.IO.FileStream)
            Dim ObjReader As New System.IO.StreamReader(OFS)

            While (ObjReader.Peek() > -1)
                strArr = Strings.Split(ObjReader.ReadLine(), ",")
                If strArr IsNot Nothing Then
                    Dim TempMap As New clsMapPattern

                    TempMap.Source = strArr(0)
                    TempMap.Target = strArr(1)
                    TempMap.Project = ProjObj
                    ProjObj.Maplist.Add(TempMap)

                    i = DataGridView1.Rows.Add(strArr)
                    DataGridView1.Rows(i).Tag = TempMap
                End If
            End While

            ObjReader.Close()
            OFS.Close()

        Catch ex As Exception
            LogError(ex, "btnAddFile_Click frmMapList")
        End Try

    End Sub

#End Region

#Region "DataGridView"

    Private Sub DataGridView1_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick

        Try
            If IsEventFromCode = False Then
                TempMapObj = DataGridView1.CurrentRow.Tag

                If CurMapObj IsNot TempMapObj Then
                    If SrcLoaded = True And TgtLoaded = False Then
                        If TempMapObj.HasSource = True Then
                            CurMapObj.Source = TempMapObj.Source
                            CurMapObj.Target = TempMapObj.Target
                        Else
                            CurMapObj.Target = TempMapObj.Target
                        End If
                    End If
                    If TgtLoaded = True And SrcLoaded = False Then
                        If TempMapObj.HasTarget = True Then
                            CurMapObj.Source = TempMapObj.Source
                            CurMapObj.Target = TempMapObj.Target
                        Else
                            CurMapObj.Source = TempMapObj.Source
                        End If
                    End If
                    If TgtLoaded = False And SrcLoaded = False Then
                        CurMapObj.Source = TempMapObj.Source
                        CurMapObj.Target = TempMapObj.Target
                    End If
                    If TgtLoaded = True And SrcLoaded = True Then
                        CurMapObj.Source = TempMapObj.Source
                        CurMapObj.Target = TempMapObj.Target
                    End If
                    If CurMapObj.Source.Trim <> "" Then
                        SrcLoaded = True
                    Else
                        SrcLoaded = False
                    End If
                    If CurMapObj.Target.Trim <> "" Then
                        TgtLoaded = True
                    Else
                        TgtLoaded = False
                    End If

                    IsEventFromCode = True
                    txtAddSource.Text = CurMapObj.Source
                    txtAddTarget.Text = CurMapObj.Target
                    IsEventFromCode = False
                End If
            End If

        Catch ex As Exception
            LogError(ex, "frmMapList DGV1_CellMouseClick")
        End Try

    End Sub

    Private Sub txtAddSourceTarget_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddSource.TextChanged, txtAddTarget.TextChanged

        Try
            If IsEventFromCode = False Then
                If txtAddSource.Text = "" And txtAddTarget.Text = "" Then
                    CurMapObj = New clsMapPattern
                    SrcLoaded = False
                    TgtLoaded = False
                Else
                    CurMapObj.Source = txtAddSource.Text
                    CurMapObj.Target = txtAddTarget.Text
                End If
                btnAdd.Enabled = True
            End If

        Catch ex As Exception
            LogError(ex, "frmMapList txtchg")
        End Try

    End Sub

#End Region
   
    Function GetMapList(Optional ByVal MapListPath As String = "") As Boolean

        Dim strArr As String()
        Dim i As Integer = 0

        Try
            '/// this first set of logic is to make sure that if a file exists, it is opened
            '/// and if it doesn't exist, it is created. 
            '/// even if the registry entry is wrong or not created yet.

            'First see if the Project already has a maplist file that matches the filepath loaded from the registry
            If MapListPath <> "" Then
                If System.IO.File.Exists(MapListPath) = True Then
                    Filepath = MapListPath
                Else
                    ' this means that that the filename from the registry doesn't exist
                    ' someone deleted or moved the maplist file, so create a new one
                    GoTo create
                End If
            Else
                '/// see if a file exists for this project 
                '/// if it doesn't, create a new one
create:         If System.IO.File.Exists(GetAppProj() & ProjObj.ProjectName & "MapList.csv") = False Then
                    '/// Create a new Maplist file in the Temp directory
                    Dim newfs As System.IO.FileStream
                    newfs = System.IO.File.Create(GetAppProj() & ProjObj.ProjectName & "MapList.csv")
                    newfs.Close()
                End If
                ProjObj.MapListPath = GetAppProj() & ProjObj.ProjectName & "MapList.csv"
                Filepath = ProjObj.MapListPath
            End If

            OpenFD1.FileName = Filepath

            '/// open the maplist file for reading, clear the existing maplist array
            Dim OFS As System.IO.FileStream = CType(OpenFD1.OpenFile, System.IO.FileStream)
            Dim ObjReader As New System.IO.StreamReader(OFS)

            ProjObj.Maplist.Clear()

            '// now read in the maplist array from the file
            While (ObjReader.Peek() > -1)
                strArr = Strings.Split(ObjReader.ReadLine(), ",")
                If strArr IsNot Nothing Then
                    Dim TempMap As New clsMapPattern

                    TempMap.Source = strArr(0)
                    TempMap.Target = strArr(1)
                    TempMap.Project = ProjObj
                    i = ProjObj.Maplist.Add(TempMap)

                    DataGridView1.Rows.Add(strArr)
                    DataGridView1.Rows(i).Tag = TempMap
                End If
            End While

            ObjReader.Close()
            OFS.Close()

            '/// set the maplist path to the file just read in
            ProjObj.MapListPath = Filepath

            Return True

        Catch ex As Exception
            LogError(ex, "frmMapList GetMapList")
            Return False
        End Try

    End Function

    Function SaveMapList(Optional ByVal saveAs As Boolean = False) As Boolean

        Dim i As Integer

        Try
            SaveFD1.FileName = ProjObj.MapListPath

            If saveAs = True Then
                If System.IO.File.Exists(SaveFD1.FileName) Then
                    SaveFD1.ShowDialog(Me)
                End If
            End If

            If SaveFD1.FileName <> "" Then
                Dim fs As System.IO.FileStream = CType(SaveFD1.OpenFile(), System.IO.FileStream)
                Dim sw As New System.IO.StreamWriter(fs)

                For i = 0 To ProjObj.Maplist.Count - 1
                    If CType(ProjObj.Maplist.Item(i), clsMapPattern) IsNot Nothing Then
                        sw.WriteLine(CType(ProjObj.Maplist.Item(i), clsMapPattern).ObjString)
                    End If
                Next
                sw.Close()
                fs.Close()
            End If

            ProjObj.MapListPath = SaveFD1.FileName

            Return True

        Catch ex As Exception
            LogError(ex, "frmMapList SaveMapList")
            Return False
        End Try

    End Function

    Function NewOrOpen(ByRef CurProject As clsProject, Optional ByVal SrcNames As ArrayList = Nothing, Optional ByVal TgtNames As ArrayList = Nothing) As Boolean

        IsEventFromCode = True

        ProjObj = CurProject
        CurMapObj = New clsMapPattern

        If ProjObj.MapListPath <> "" Then
            If GetMapList(ProjObj.MapListPath) = False Then
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.Cancel
                GoTo doAgain
            End If
        Else
            If GetMapList() = False Then
                Me.Close()
                DialogResult = Windows.Forms.DialogResult.Cancel
                GoTo doAgain
            End If
        End If

        '/// import new source and/or target entry and give focus to other side of entry
        If SrcNames IsNot Nothing And TgtNames IsNot Nothing Then
            If SrcNames.Count = TgtNames.Count Then
                Dim i As Integer
                For i = 0 To SrcNames.Count - 1
                    IsEventFromCode = True
                    CurMapObj.Source = SrcNames(i)
                    CurMapObj.Target = TgtNames(i)
                    txtAddSource.Text = CurMapObj.Source
                    txtAddTarget.Text = CurMapObj.Target
                    SrcLoaded = True
                    TgtLoaded = True
                    Call btnAdd_Click(Me, New EventArgs)
                Next
                txtAddSource.Focus()
                txtAddSource.Select()
            End If
        Else
            If SrcNames IsNot Nothing Then
                If SrcNames.Count > 0 Then
                    For Each src As String In SrcNames
                        IsEventFromCode = True
                        CurMapObj.Source = src
                        txtAddSource.Text = CurMapObj.Source
                        SrcLoaded = True
                        Call btnAdd_Click(Me, New EventArgs)
                    Next
                    txtAddTarget.Focus()
                    txtAddTarget.Select()
                End If
            End If
            If TgtNames IsNot Nothing Then
                If TgtNames.Count > 0 Then
                    For Each Tgt As String In TgtNames
                        IsEventFromCode = True
                        CurMapObj.Target = Tgt
                        txtAddTarget.Text = CurMapObj.Target
                        TgtLoaded = True
                        Call btnAdd_Click(Me, New EventArgs)
                    Next
                    txtAddSource.Focus()
                    txtAddSource.Select()
                End If
            End If
            If TgtNames Is Nothing And SrcNames Is Nothing Then
                txtAddSource.Focus()
                txtAddSource.Select()
            End If
        End If

        CurMapObj.Project = ProjObj
        Windows.Forms.Cursor.Show()
        IsEventFromCode = False

doAgain:
        Select Case Me.ShowDialog
            Case Windows.Forms.DialogResult.OK
                CurProject.MapListPath = ProjObj.MapListPath
                NewOrOpen = True
            Case Windows.Forms.DialogResult.Retry
                GoTo doAgain
            Case Else
                NewOrOpen = False
        End Select

    End Function

End Class
