Public Class frmScriptImport
    Inherits frmBlank

    Dim InputDir As String = ""
    Dim InFilePath As String = ""
    Dim objThis As INode
    Dim INLstream As System.IO.FileStream
    Dim InLineList As New Collection
    Dim PreParsed As Boolean
    Dim InScriptType As enumInType

    Enum enumInType
        None = 0
        INL = 1
        RPT = 2
    End Enum

#Region "Form Button Click Events"

    '/// Import Script to Studio Event
    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        Try

        Catch ex As Exception
            LogError(ex, "frmScriptImport cmdOk_Click_1")
        End Try

    End Sub

    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Try
            Me.Close()
            Me.DialogResult = Windows.Forms.DialogResult.Abort

        Catch ex As Exception
            LogError(ex, "frmScriptImport cmdCancel_Click_1")
        End Try

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        Try

        Catch ex As Exception
            LogError(ex, "frmScriptImport cmdHelp_Click_1")
        End Try

    End Sub

#End Region

#Region "Load Script"

    Public Function GetObject(Optional ByVal filePath As String = "", Optional ByVal openDir As String = "") As clsImpRetCode

        Try
            '/// set and/or create filebrowser dialog paths
            InputDir = openDir
            InFilePath = filePath
            txtScriptPath.Text = InFilePath
            If filePath <> "" Then
                openDir = GetDirFromPath(filePath)
            End If

            '/// set group boxes inactive
            gbAnalyzedScript.Enabled = False
            gbOption.Enabled = False
            gbProperties.Enabled = False
            gbScriptTree.Enabled = False

            cmdOk.Enabled = False

doAgain:
            Select Case Me.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    GetObject = objThis
                Case Windows.Forms.DialogResult.Retry
                    GoTo doAgain
                Case Else
                    Return Nothing
            End Select

        Catch ex As Exception
            LogError(ex, "frmXMLconv OpenForm")
            Return Nothing
        End Try

    End Function

    '/toolbar// Open Script File Button Event
    Private Sub OpenToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripButton.Click

        Try
            If InputDir <> "" Then
                OFD1.InitialDirectory = InputDir
            End If
            OFD1.ShowDialog()

        Catch ex As Exception
            LogError(ex, "frmScriptImport OpenToolStripButton_Click")
        End Try

    End Sub

    '/toolbar// Print Script File Button Event
    Private Sub PrintToolStripButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripButton.Click

        Try

        Catch ex As Exception
            LogError(ex, "frmScriptImport PrintToolStripButton_Click")
        End Try

    End Sub

    Private Sub OFD1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OFD1.FileOk

        Try
            InFilePath = OFD1.FileName
            If InFilePath <> "" Then
                txtScriptPath.Text = InFilePath
                If LoadLVtext() = True Then
                    gbOption.Enabled = True
                Else
                    MsgBox("Please choose a valid Script file", MsgBoxStyle.OkOnly, "Invalid Script File")
                End If
            Else

            End If

        Catch ex As Exception
            LogError(ex, "frmScriptImport OFD1_FileOk")
        End Try

    End Sub

    Function LoadLVtext() As Boolean

        Try
            If InFilePath <> "" Then
                Dim Ofs As System.IO.FileStream = CType(OFD1.OpenFile(), System.IO.FileStream)
                Dim objreader As New System.IO.StreamReader(Ofs)
                Dim LineNum As Integer = 0

                InLineList.Clear()
                lvInScript.Items.Clear()

                While (objreader.Peek() > -1)
                    LineNum += 1
                    Dim NewListEntry As New MyIntlist(LineNum, objreader.ReadLine())
                    InLineList.Add(NewListEntry, NewListEntry.Idx)
                    lvInScript.Items.Add(LineNum).SubItems.Add(NewListEntry.Text)
                End While

                objreader.Close()
                Ofs.Close()
                LoadLVtext = True
            Else
                LoadLVtext = False
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptImport LoadLVtext")
            LoadLVtext = False
        End Try

    End Function

#End Region

#Region "Analyze"

    '****** Script Analyzer ******
    Private Sub btnAnalyze_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAnalyze.Click

        Try

        Catch ex As Exception
            LogError(ex, "frmScriptImport btnAnalyze_Click")
        End Try

    End Sub

#End Region

#Region "Display Script Analysis"

#End Region

End Class