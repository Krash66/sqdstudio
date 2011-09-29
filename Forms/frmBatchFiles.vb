Imports System.Windows.Forms

Public Class frmBatchFiles

    '/// Globals
    'Engine
    Dim ObjThis As clsEngine
    'File Streams
    Dim fsOut As System.IO.FileStream
    Dim objWriteOut As System.IO.StreamWriter
    Dim fsERR As System.IO.FileStream
    Dim objWriteERR As System.IO.StreamWriter
    'Paths and Names
    Dim LocCDCstore As String
    Dim NameCDCstore As String
    Dim LocCDCcap As String
    Dim NameCDCcap As String
    Dim LocSQDconf As String
    Dim LocParse As String
    Dim LocEng As String
    Dim LocScripts As String
    Dim NameODBC As String
    'Source Datastore
    Dim SrcDS As clsDatastore


#Region "Main process"

    Public Function ShowForm(ByVal eng As clsEngine) As Boolean

        Try
            ObjThis = eng
            If PopulateVariables() = False Then
                txtBatFiles.AppendText("There were problems, see log for details " & vbLf)
                txtBatFiles.AppendText("Fix incorrect properties and try again " & vbLf)
            End If

doAgain:
            Select Case Me.ShowDialog
                Case Windows.Forms.DialogResult.OK
                    ShowForm = True
                Case Else
                    GoTo doAgain
            End Select

        Catch ex As Exception
            LogError(ex, "frmBatchFiles MakeBatch")
            Return False
        End Try

    End Function

    Function PopulateVariables() As Boolean

        Dim Problem As String = ""
        Try
            Problem = "Problem with Engine Object"
            If ObjThis.Sources.Count < 1 Then
                SrcDS = Nothing
            Else
                SrcDS = ObjThis.Sources(1)
            End If

            Problem = "Problem with CDCstore Location"
            LocCDCstore = ObjThis.CDCdir
            If LocCDCstore.EndsWith("\") = False Then
                LocCDCstore = LocCDCstore & "\"
            End If

            Problem = "Problem with Source Datastore Physical Name"
            NameCDCstore = SrcDS.DsPhysicalSource
            'If NameCDCstore.EndsWith(".cab") = False Then
            '    NameCDCstore = NameCDCstore & ".cab"
            'End If

            Problem = "Problem with CDCcapture Location"
            LocCDCcap = LocCDCstore
            If LocCDCcap.EndsWith("\") = False Then
                LocCDCcap = LocCDCcap & "\"
            End If

            Problem = "Problem with CDCcapture Name"
            NameCDCcap = "mscapture.cab"

            Problem = "Problem with Executables Location"
            LocSQDconf = ObjThis.EXEdir
            If LocSQDconf.EndsWith("\") = False Then
                LocSQDconf = LocSQDconf & "\"
            End If
            LocParse = LocSQDconf
            LocEng = LocSQDconf

            Problem = "Problem with Script Directory, Check Environment"
            LocScripts = ObjThis.ObjSystem.Environment.LocalScriptDir
            If LocScripts.EndsWith("\") = False Then
                LocScripts = LocScripts & "\"
            End If

            Problem = "Problem with Engine connection"
            NameODBC = ObjThis.Connection.Database
            If NameODBC = "" Then
                NameODBC = "<database name>"
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles PopulateVariables", Problem)
            Return False
        End Try

    End Function

#End Region

#Region "Button Clicks"

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCreate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreate.Click
        MasterBuild()
    End Sub

#End Region

#Region "Build Files"

    Function MasterBuild() As Boolean

        Try
            Dim success As Boolean = True

            While success
                success = ApplyFile()
                success = DisplayFile()
                success = MountFile()
                success = UnMountFile()
                success = ModifyFile()
                success = RemoveAllFile()
                success = ParseFile()
                success = RunEngineFile()
                success = ConfigFile()
                success = StartFile()
                success = StopFile()
                Exit While
            End While

            Return success

        Catch ex As Exception
            LogError(ex, "frmbatchFiles MasterBuild")
            txtBatFiles.AppendText("Remaining Files Failed!!!!" & vbLf)
            Return False
        End Try

    End Function

    Function CreateFile(ByVal FullFileName As String) As Boolean

        Try
            '//Create new file and filestream for Output file
            If System.IO.File.Exists(FullFileName) Then
                If cbOverwrite.Checked = True Then
                    System.IO.File.Delete(FullFileName)
                Else
                    txtBatFiles.AppendText("Skipped!!!!... " & GetFileNameFromPath(FullFileName) & "  " & vbLf)
                    Return False
                    Exit Try
                End If
            End If
            fsOut = System.IO.File.Create(FullFileName)
            objWriteOut = New System.IO.StreamWriter(fsOut)

            Return True

        Catch ex As Exception
            LogError(ex, "modModeler CreateFile")
            txtBatFiles.AppendText("Failed to create file!!!!... " & GetFileNameFromPath(FullFileName) & "  " & vbLf)
            Return False
        End Try

    End Function

    Function wRemLine(Optional ByVal msg As String = "") As Boolean

        Try
            If msg = "" Then
                objWriteOut.WriteLine("REM")
            Else
                objWriteOut.WriteLine("REM ***" & msg & "***")
            End If

            wRemLine = True

        Catch ex As Exception
            LogError(ex, "modGenerate wRemLine")
            wRemLine = False
        End Try

    End Function

#End Region

#Region "File Building"

    Function ApplyFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Apply.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" Apply to SQDconf ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf apply " & LocCDCcap & NameCDCcap & " --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Apply.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles ApplyFile")
            txtBatFiles.AppendText("Created internal text failed!... Apply.bat " & vbLf)
            Return True
        End Try

    End Function

    Function DisplayFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Display.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" Display SQDconf Configuration ")
            wRemLine(" Displays properties of CDC store ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf display " & LocCDCstore & NameCDCstore & ".cab --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Display.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles DisplayFile")
            txtBatFiles.AppendText("Created internal text failed!... Display.bat " & vbLf)
            Return True
        End Try

    End Function

    Function MountFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Mount.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" Mount CDC store ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf mount " & LocCDCcap & NameCDCcap & " --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Mount.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles MountFile")
            txtBatFiles.AppendText("Created internal text failed!... Mount.bat " & vbLf)
            Return True
        End Try

    End Function

    Function UnMountFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\UnMount.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" UnMount CDC store ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf unmount " & LocCDCcap & NameCDCcap & " --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... UnMount.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles UnMountFile")
            txtBatFiles.AppendText("Created internal text failed!... UnMount.bat " & vbLf)
            Return True
        End Try

    End Function

    Function ModifyFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Modify.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" Generic Modify Capture ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf modify " & LocCDCcap & NameCDCcap & " --cdcstore=" & LocCDCstore & _
                                  NameCDCstore & ".cab")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Modify.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles ModifyFile")
            txtBatFiles.AppendText("Created internal text failed!... Modify.bat " & vbLf)
            Return True
        End Try

    End Function

    Function RemoveAllFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\RemoveAll.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If

            wRemLine()
            wRemLine(" Remove all tables from capture ")
            wRemLine(" Note: check schema and table names ")
            wRemLine()
            For Each sel As clsDSSelection In SrcDS.ObjSelections
                Dim TableName As String = sel.SelectionName
                objWriteOut.WriteLine(LocSQDconf & "sqdconf remove " & LocCDCcap & NameCDCcap & " --schema=dbo --table=" & TableName)
            Next

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... RemoveAll.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles RemoveAllFile")
            txtBatFiles.AppendText("Created internal text failed!... RemoveAll.bat " & vbLf)
            Return True
        End Try

    End Function

    Function ParseFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Parse.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If
            wRemLine()
            wRemLine(" Parse Engine Script: ")
            wRemLine(" " & LocScripts & ObjThis.EngineName & ".sqd ")
            wRemLine()
            objWriteOut.WriteLine(LocParse & "sqdparse " & _
                                  LocScripts & ObjThis.EngineName & ".sqd " & _
                                  LocScripts & ObjThis.EngineName & ".prc " & _
                                  "LIST=ALL A B C D E F G H 1")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Parse.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles ParseFile")
            txtBatFiles.AppendText("Created internal text failed!... Parse.bat " & vbLf)
            Return True
        End Try

    End Function

    Function RunEngineFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\RunEngine.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If
            wRemLine()
            wRemLine(" Run Engine Parsed Script: ")
            wRemLine(" " & LocScripts & ObjThis.EngineName & ".prc ")
            wRemLine()
            objWriteOut.WriteLine(LocEng & "sqdata " & _
                                  LocScripts & ObjThis.EngineName & ".prc " & _
                                  "--log-level=8 > " & _
                                  LocScripts & ObjThis.EngineName & ".rpt " & _
                                  "2>&1")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... RunEngine.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles RunEngineFile")
            txtBatFiles.AppendText("Created internal text failed!... RunEngine.bat " & vbLf)
            Return True
        End Try

    End Function

    Function ConfigFile() As Boolean

        Try
            Dim AccessHost As String = ObjThis.ObjSystem.Host.Trim
            If AccessHost = "" Then
                AccessHost = "localhost"
            End If

            Dim path As String = ObjThis.BATdir & "\Config.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If
            '/// Start Building Text
            wRemLine()
            wRemLine(" Create the SQLserver Capture ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf create " & LocCDCstore & NameCDCstore & _
                                  " --type=cdcstore --alias=" & NameCDCstore & _
                                  " --data-path=" & LocCDCcap)
            objWriteOut.WriteLine(LocSQDconf & "sqdconf create " & LocCDCcap & NameCDCcap & _
                                  " --type=mssql --database=" & NameODBC & _
                                  " --cdcstore=" & LocCDCstore & NameCDCstore & ".cab")


            wRemLine()
            wRemLine(" Add Tables to Capture ")
            wRemLine()
            For Each sel As clsDSSelection In SrcDS.ObjSelections
                Dim TableName As String = sel.SelectionName
                objWriteOut.WriteLine(LocSQDconf & "sqdconf add " & LocCDCcap & NameCDCcap & " --schema=dbo --table=" & TableName & _
                                      " --datastore=cdc://" & AccessHost & "/" & SrcDS.DsPhysicalSource & "/" & SrcDS.DatastoreName & _
                                      " --active")
            Next

            wRemLine()
            wRemLine(" Mount cdcstore ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf mount " & LocCDCcap & NameCDCcap & " --log-level=8")

            wRemLine()
            wRemLine(" time function to pause batch ")
            wRemLine("time")
            wRemLine()

            wRemLine()
            wRemLine(" Set to Start from beginning of log ")
            wRemLine(LocSQDconf & "sqdconf modify --lsn=00000000:00000000:0000 " & LocCDCcap & NameCDCcap)
            wRemLine()

            wRemLine()
            wRemLine(" Apply configuration ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf apply " & LocCDCcap & NameCDCcap & " --log-level=8")

            wRemLine()
            wRemLine(" Start Capture Agent ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf start " & LocCDCcap & NameCDCcap & " --log-level=8")
            '/// Finished building text

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Config.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles ConfigFile")
            txtBatFiles.AppendText("Created internal text failed!... Config.bat " & vbLf)
            Return True
        End Try

    End Function

    Function StartFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Start.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If
            wRemLine()
            wRemLine(" Start Capture ")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf start " & LocCDCcap & NameCDCcap & " --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Start.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles StartFile")
            txtBatFiles.AppendText("Created internal text failed!... Start.bat " & vbLf)
            Return True
        End Try

    End Function

    Function StopFile() As Boolean

        Try
            Dim path As String = ObjThis.BATdir & "\Stop.bat"
            If CreateFile(path) = False Then
                Return True
                Exit Function
            End If
            wRemLine()
            wRemLine(" Stop Capture")
            wRemLine()
            objWriteOut.WriteLine(LocSQDconf & "sqdconf stop " & LocCDCcap & NameCDCcap & " --log-level=8")

            objWriteOut.Close()
            fsOut.Close()
            txtBatFiles.AppendText("Created and Built!... Stop.bat " & vbLf)
            Return True

        Catch ex As Exception
            LogError(ex, "frmBatchFiles StopFile")
            txtBatFiles.AppendText("Created internal text failed!... Stop.bat " & vbLf)
            Return True
        End Try

    End Function

#End Region

End Class
