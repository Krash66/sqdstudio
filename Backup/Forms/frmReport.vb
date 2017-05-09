Public Class frmReport

    Private Struct As clsStructure = Nothing

#Region "Form Events"

    Public Function Report(ByVal InVal As clsStructure) As frmReport

        Try
            Struct = InVal
            txtName.Text = Struct.StructureName
            DGV1.Name = Struct.StructureName

        Catch ex As Exception
            Report = Nothing
            Exit Function
        End Try

        Me.WindowState = FormWindowState.Normal
        Me.Show()
        Report = Me

    End Function

#End Region

#Region "DataGrid Events"

    Private Sub mnuClrCell_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuClrCell.Click

    End Sub

#End Region

#Region "Click Events"

    Private Sub btnGenRep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenRep.Click

        If GenerateReport() = False Then

        End If

    End Sub

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Dim i, j As Integer
        Dim ColCount As Integer = DGV1.ColumnCount
        Dim RowCount As Integer = DGV1.RowCount
        Dim currentfield As String = ""
        Dim strFilter As String = "Comma Separated Variable Files (*.csv)|*.csv|All files (*.*)|*.*"

        Try
            SaveFD.InitialDirectory = ReportFolderPath()
            SaveFD.FileName = Struct.StructureName & "-Report"
            SaveFD.Filter = strFilter

            SaveFD.ShowDialog()
            ' If the file name is not an empty string open it for saving.
            If SaveFD.FileName <> "" Then
                ' Saves the Query File via a FileStream created by the OpenFile method.
                Dim fs As System.IO.FileStream = CType(SaveFD.OpenFile(), System.IO.FileStream)
                Dim objWriter As New System.IO.StreamWriter(fs)
                '/// first write the Table Name
                objWriter.WriteLine(DGV1.Name)
                objWriter.WriteLine()
                objWriter.WriteLine()

                '/// now write the data
                For i = 0 To RowCount - 1
                    For j = 0 To ColCount - 1
                        If Not DGV1.Item(j, i) Is System.DBNull.Value Then
                            currentfield = DGV1.Item(j, i).Value.ToString
                            objWriter.Write(currentfield & ",")
                        Else
                            objWriter.Write("Null" & ",")
                        End If
                    Next
                    If Not DGV1.Item(ColCount - 1, i) Is System.DBNull.Value Then
                        currentfield = DGV1.Item(ColCount - 1, i).Value.ToString
                        objWriter.Write(currentfield)
                    Else
                        objWriter.Write("Null")
                    End If
                    objWriter.WriteLine()
                Next
                'fs.Close()
                objWriter.Close()
            End If


        Catch ex As Exception
            LogError(ex, "frmReport cmdOk_click")
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(HHId.H_Generate_Reports)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown

        Select Case e.KeyCode
            Case Keys.Escape
                cmdCancel_Click_1(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click_1(sender, New EventArgs)
        End Select

    End Sub

#End Region

#Region "Generate Functions"

    Function GenerateReport() As Boolean

        Dim KeyCount As Integer
        Dim RowCnt As Integer
        Dim i As Integer = 0
        Dim j, k As Integer

        Try
            KeyCount = GetKeyNum()
            RowCnt = 24 + KeyCount
            DGV1.RowCount() = RowCnt
            DGV1.ColumnCount() = 5

            For j = 0 To RowCnt - 1
                For k = 0 To 4
                    DGV1.Item(k, j).Value = New Mylist("", "")
                Next
            Next
            FillText(RowCnt, KeyCount)

            DGV1.Item(0, 9).Value = New Mylist(Struct.IMSDBName, Struct.IMSDBName)
            DGV1.Item(3, 9).Value = New Mylist(GetFileNameWithoutExtenstionFromPath(Struct.fPath1), GetFileNameWithoutExtenstionFromPath(Struct.fPath1))
            DGV1.Item(1, 10).Value = New Mylist(Struct.StructureName, Struct.StructureName)

            For Each fld As clsField In Struct.ObjFields
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                    i += 1
                    DGV1.Item(1, 16 + i).Value = New Mylist(fld.FieldName, fld.FieldName)
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmReport GenerateReport")
            Return False
        End Try

    End Function

    Function GetKeyNum() As Integer

        Dim Count As Integer = 0
        Try
            For Each fld As clsField In Struct.ObjFields
                If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                    Count += 1
                End If
            Next
            Return Count

        Catch ex As Exception
            LogError(ex, "frmReport GetKeyNum")
            Return 0
        End Try

    End Function

    Sub FillText(ByVal RowCount As Integer, ByVal Kcnt As Integer)

        Try
            DGV1.Item(0, 0).Value = New Mylist("DATE SUBMITTED:", "DATE SUBMITTED:")
            DGV1.Item(2, 0).Value = New Mylist("BY:", "BY:")
            DGV1.Item(0, 1).Value = New Mylist("Review Meeting:", "Review Meeting:")
            DGV1.Item(3, 1).Value = New Mylist("Attendees:", "Attendees:")
            DGV1.Item(0, 2).Value = New Mylist("Date Revised:", "Date Revised:")
            DGV1.Item(0, 3).Value = New Mylist("Revised for Data Types:", "Revised for Data Types:")
            DGV1.Item(0, 4).Value = New Mylist("Customer Review:", "Customer Review:")
            DGV1.Item(0, 5).Value = New Mylist("Final Review Date:", "Final Review Date:")
            DGV1.Item(2, 6).Value = New Mylist("APPROVED BY:", "APPROVED BY:")
            DGV1.Item(4, 6).Value = New Mylist("APPROVAL DATE:", "APPROVAL DATE:")
            DGV1.Item(0, 8).Value = New Mylist("PHYSICAL NAME:", "PHYSICAL NAME:")
            DGV1.Item(3, 8).Value = New Mylist("FILE NAME (DD NAME):", "FILE NAME (DD NAME):")
            DGV1.Item(0, 10).Value = New Mylist("Description Name:", "Description Name:")
            DGV1.Item(3, 10).Value = New Mylist("Assigned Table Name:", "Assigned Table Name:")
            DGV1.Item(4, 10).Value = New Mylist("New Table Name:", "New Table Name:")
            DGV1.Item(0, 11).Value = New Mylist("QUEUED ORDER ERROR REC:", "QUEUED ORDER ERROR REC:")
            DGV1.Item(0, 12).Value = New Mylist("QUEUED ORDER ERROR TABLE:", "QUEUED ORDER ERROR TABLE:")
            DGV1.Item(2, 14).Value = New Mylist("Change Needed? (y/n)", "Change Needed? (y/n)")
            DGV1.Item(3, 14).Value = New Mylist("Delete Item?", "Delete Item?")
            DGV1.Item(4, 14).Value = New Mylist("Additional Comments", "Additional Comments")
            DGV1.Item(0, 16).Value = New Mylist("Primary Key Columns/Order:", "Primary Key Columns/Order:")
            DGV1.Item(0, 17 + Kcnt).Value = New Mylist("Character fields reviewed for size/datatype?", _
            "Character fields reviewed for size/datatype?")
            DGV1.Item(0, 18 + Kcnt).Value = New Mylist("Numeric fields reviewed for size/datatype?", _
            "Numeric fields reviewed for size/datatype?")
            DGV1.Item(0, 19 + Kcnt).Value = New Mylist("Date fields reviewed/format defined?", _
            "Date fields reviewed/format defined?")
            DGV1.Item(0, 20 + Kcnt).Value = New Mylist("Field names reviewed/modified?", "Field names reviewed/modified?")
            DGV1.Item(0, 21 + Kcnt).Value = New Mylist("Additional Indexes needed?", "Additional Indexes needed?")
            DGV1.Item(0, 22 + Kcnt).Value = New Mylist("OTHER:", "OTHER:")

        Catch ex As Exception
            LogError(ex, "frmReport FillText")
        End Try

    End Sub

#End Region

End Class
