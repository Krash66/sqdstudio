Public Class frmScriptGen

    Private RetCode As clsRcode
    Private ScriptDirectory As String
    Private ObjEng As clsEngine
    Private ObjProc As clsTask
    Private ObjDS As clsDatastore
    Private debugSrc As Boolean = False
    Private ShowUID As Boolean = False
    Private IsEventFromCode As Boolean
    Private sourcelevel As enumMappingLevel
    Private targetlevel As enumMappingLevel
    Private ScriptType As enumGenType


    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(modHelp.HHId.H_Generate_Script)

    End Sub

    Public Sub MyCTL_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFolderPath.KeyDown, txtINL.KeyDown, txtRPT.KeyDown, txtSQD.KeyDown

        'Dim TypeTask As enumTaskType
        'TypeTask = objThis.TaskType
        Select Case e.KeyCode
            Case Keys.Escape
                'cmdCancel_Click(sender, New EventArgs)
            Case Keys.F1
                cmdHelp_Click(sender, New EventArgs)
        End Select
    End Sub

    Private Sub btnExpFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpFolder.Click

        Try
            If txtFolderPath.Text <> "" Then
                Shell("explorer.exe " & txtFolderPath.Text, AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnExpFolder_Click")
        End Try

    End Sub

    Private Sub btnFTP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFTP.Click

        Try
            Dim retStr As String = ""
            Dim frmFTP As frmFTPClient
            frmFTP = New frmFTPClient

            retStr = frmFTP.BrowseFile(GetDirFromPath(RetCode.SQDPath), ".*", enumCalledFrom.BY_SCRIPTGEN)


        Catch ex As Exception
            LogError(ex, "frmScriptGen btnFTP_Click")
        End Try

    End Sub

    Private Sub btnOpenSQD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenSQD.Click

        Try
            If txtSQD.Text <> "" Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(RetCode.SQDPath)

                'Shell("notepad.exe " & RetCode.SQDPath, AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnOpenSQD_Click")
        End Try

    End Sub

    Private Sub btnOpenTMP_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Try
            'If txtTMP.Text <> "" Then
            '    Dim OpenProcess As New System.Diagnostics.Process
            '    Process.Start(RetCode.TMPPath)

            '    'Shell("notepad.exe " & RetCode.TMPPath, AppWinStyle.NormalFocus)
            'End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnOpenTMP_Click")
        End Try

    End Sub

    Private Sub btnOpenINL_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenINL.Click

        Try
            If txtINL.Text <> "" Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(RetCode.INLPath)

                'Shell("notepad.exe " & RetCode.INLPath, AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnOpenINL_Click")
        End Try

    End Sub

    Private Sub btnOpenRPT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenRPT.Click

        Try
            If txtRPT.Text <> "" Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(RetCode.RPTPath)

                'Shell("notepad.exe " & RetCode.RPTPath, AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnOpenRPT_Click")
        End Try

    End Sub

    Public Sub OpenFormEng(ByVal EngObj As clsEngine, ByVal SavePath As String)

        Try
            Startload()

            Me.rbAllsrc.Checked = False
            Me.rbSelectsrc.Checked = True
            Me.rbFieldsrc.Checked = False
            Me.rbAlltgt.Checked = False
            Me.rbSelecttgt.Checked = True
            Me.rbFieldtgt.Checked = False
            sourcelevel = enumMappingLevel.ShowDesc
            targetlevel = enumMappingLevel.ShowDesc

            ObjEng = EngObj
            ScriptType = enumGenType.Eng
            ScriptDirectory = SavePath


            Endload()

            Me.Show()

        Catch ex As Exception
            LogError(ex, "frmScriptGen OpenScripts")
        End Try

    End Sub

    Public Sub OpenFormDS(ByVal DSObj As clsDatastore, ByVal SavePath As String)

        Try
            Startload()

            Me.rbAllsrc.Checked = False
            Me.rbSelectsrc.Checked = True
            Me.rbFieldsrc.Checked = False
            Me.rbAlltgt.Checked = False
            Me.rbSelecttgt.Checked = True
            Me.rbFieldtgt.Checked = False
            sourcelevel = enumMappingLevel.ShowDesc
            targetlevel = enumMappingLevel.ShowDesc

            ObjDS = DSObj
            ObjEng = ObjDS.Engine
            ScriptType = enumGenType.DS

            ScriptDirectory = SavePath

            Endload()

            Me.Show()

        Catch ex As Exception
            LogError(ex, "frmScriptGen OpenScripts")
        End Try

    End Sub

    Public Sub OpenFormProc(ByVal ProcObj As clsTask, ByVal SavePath As String)

        Try
            Startload()

            Me.rbAllsrc.Checked = False
            Me.rbSelectsrc.Checked = True
            Me.rbFieldsrc.Checked = False
            Me.rbAlltgt.Checked = False
            Me.rbSelecttgt.Checked = True
            Me.rbFieldtgt.Checked = False
            sourcelevel = enumMappingLevel.ShowDesc
            targetlevel = enumMappingLevel.ShowDesc

            ObjProc = ProcObj
            ObjEng = ObjProc.Engine
            ScriptType = enumGenType.Proc
            ScriptDirectory = SavePath

            Endload()

            Me.Show()

        Catch ex As Exception
            LogError(ex, "frmScriptGen OpenScripts")
        End Try

    End Sub

    Sub Startload()

        IsEventFromCode = True
        Me.gbPath.Enabled = False
        Me.gbStudioFiles.Enabled = False
        Me.gbParseFiles.Enabled = False
        Me.btnOpenOutput.Enabled = False

        cmdOk.Visible = False
        cmdCancel.Text = "Close"

    End Sub

    Sub Endload()

        IsEventFromCode = False

    End Sub

    Function FillResults(ByVal Parsed As Boolean) As Boolean

        Try
            lblResult.Text = "Script Generation resulted in " & RetCode.ErrorCount.ToString & " Errors" & Chr(13)
            lblSummary.Text = GetMsg()

            txtFolderPath.Text = GetDirFromPath(RetCode.SQDPath)
            txtSQD.Text = GetFileNameOnly(RetCode.SQDPath)
            'txtTMP.Text = GetFileNameOnly(RetCode.TMPPath)
            txtINL.Text = GetFileNameOnly(RetCode.INLPath)
            txtRPT.Text = GetFileNameOnly(RetCode.RPTPath)
            'txtPRC.Text = GetFileNameOnly(RetCode.PRCPath)

            btnFTP.Enabled = True
            Me.gbPath.Enabled = True
            Me.gbStudioFiles.Enabled = True
            Me.gbParseFiles.Enabled = Parsed
            Me.btnSQData.Enabled = Parsed

            Return True

        Catch ex As Exception
            LogError(ex, "frmScriptGen FillFormInfo")
            Return False
        End Try

    End Function

    Function GetMsg() As String

        Try
            Dim Message As String

            If RetCode.HasError = False Then
                '/// Good Script
                Message = "Script Generator wrote:" & Chr(13) & _
                    RetCode.SQDline & " lines to .SQD file and" & Chr(13) & _
                    RetCode.INLline & " lines to .INL file" & Chr(13) & _
                    RetCode.ReturnCode
                'RetCode.TMPline & " lines to .TMP file" & Chr(13) & _

                'If MsgBox("Script Generator wrote:" & Chr(13) & _
                '    RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
                '    RetCode.INLline & " lines to .INL file and" & Chr(13) & _
                '    RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
                '    RetCode.ReturnCode, MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.Yes Then
                '    Shell("notepad.exe " & Quote(RetCode.Path, """"), AppWinStyle.NormalFocus)
                'End If
            Else
                If RetCode.ErrorLocation = enumErrorLocation.ModGenSQDParse Then
                    '/// Error while Parsing Script
                    Message = "Script Caused an Error while Parsing:" & Chr(13) & _
                    RetCode.ReturnCode & Chr(13) & _
                    "Parser Return Code: " & RetCode.ParserRC.ToString & Chr(13) & _
                    "Generator wrote:" & Chr(13) & _
                    RetCode.SQDline & " lines to .SQD file and" & Chr(13) & _
                    RetCode.INLline & " lines to .INL file"  ' & Chr(13) & _
                    'RetCode.TMPline & " lines to .TMP file"
                    'If MsgBox("Script Caused an Error while Parsing:" & Chr(13) & Chr(13) & _
                    'RetCode.ReturnCode & Chr(13) & Chr(13) & _
                    '"Generator wrote:" & Chr(13) & _
                    'RetCode.SQDline & " lines to .SQD file," & Chr(13) & _
                    'RetCode.INLline & " lines to .INL file and" & Chr(13) & _
                    'RetCode.TMPline & " lines to .TMP file" & Chr(13) & Chr(13) & _
                    '"Would You Like to view the Script Report?", MsgBoxStyle.YesNo, _
                    '"Script Generation Parse Error") = MsgBoxResult.Yes Then
                    '    Shell("notepad.exe " & Quote(RetCode.ParserPath, """"), AppWinStyle.NormalFocus)
                    'End If
                Else
                    '/// Error in Windows (e.g. Object problems)
                    Message = "Script Generation Returned Errors:" & Chr(13) & _
                    "Generator wrote:" & Chr(13) & _
                    RetCode.SQDline & " lines to .SQD file and" & Chr(13) & _
                    RetCode.INLline & " lines to .INL file"   '& Chr(13) & _
                    'RetCode.TMPline & " lines to .TMP file"
                    'If MsgBox("Script Generation Returned The Following Error:" & Chr(13) & _
                    'RetCode.GetGUIErrorMsg() & Chr(13) & Chr(13) & _
                    '"Error Occured at:" & Chr(13) & _
                    '"SQD file line Number: " & RetCode.SQDline & Chr(13) & _
                    '"INL file Line Number: " & RetCode.INLline & Chr(13) & _
                    '"TMP file Line Number: " & RetCode.TMPline & Chr(13) & Chr(13) & _
                    '"Would You Like to view the Script?", MsgBoxStyle.YesNo, "Script Generation Error") = MsgBoxResult.Yes Then
                    '    Shell("notepad.exe " & Quote(RetCode.ErrorPath, """"), AppWinStyle.NormalFocus)
                    'End If
                End If
            End If

            GetMsg = Message

        Catch ex As Exception
            LogError(ex, "frmScriptGen GetMsg")
            GetMsg = ""
        End Try

    End Function

    Private Sub cbDebugSrcData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDebugSrcData.CheckedChanged

        If IsEventFromCode = True Then Exit Sub
        debugSrc = cbDebugSrcData.Checked

    End Sub

    Private Sub cbUID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbUID.CheckedChanged

        If IsEventFromCode = True Then Exit Sub
        ShowUID = cbUID.Checked

    End Sub

    Private Sub btnGenScript_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenScript.Click

        Try
            RetCode = New clsRcode

            Select Case ScriptType
                Case enumGenType.DS
                    RetCode = modGenerateV3.GenerateDSScriptV3(ObjDS, ScriptDirectory, True, debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Proc
                    RetCode = modGenerateV3.GenerateProcScriptV3(ObjProc, ScriptDirectory, True, debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Eng
                    RetCode = modGenerateV3.GenerateEngScriptV3(ObjEng, ScriptDirectory, True, debugSrc, ShowUID, sourcelevel, targetlevel)

            End Select

            FillResults(False)

            Log("********* Script Generation End : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnGenScript_Click")
        End Try

    End Sub

    Private Sub btnParse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParse.Click

        Try
            RetCode = New clsRcode

            Select Case ScriptType
                Case enumGenType.DS
                    RetCode = modGenerateV3.GenerateDSScriptV3(ObjDS, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Proc
                    RetCode = modGenerateV3.GenerateProcScriptV3(ObjProc, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Eng
                    RetCode = modGenerateV3.GenerateEngScriptV3(ObjEng, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

            End Select
            FillResults(True)

            Log("********* Script Generation End : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnParse_Click")
        End Try

    End Sub

    Private Sub btnParseOnly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParseOnly.Click

        Try
            RetCode = New clsRcode

            Select Case ScriptType
                Case enumGenType.DS
                    RetCode = modGenerateV3.GenerateDSScriptV3(ObjDS, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Proc
                    RetCode = modGenerateV3.GenerateProcScriptV3(ObjProc, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Eng
                    RetCode = modGenerateV3.GenerateEngScriptV3(ObjEng, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel, True)

            End Select
            FillResults(True)

            Log("********* Script Parse End ***********")

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnParseOnly_Click")
        End Try

    End Sub

    Private Sub btnParseSQD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParseSQD.Click

        Try
            RetCode = New clsRcode

            Select Case ScriptType
                Case enumGenType.DS
                    RetCode = modGenerateV3.GenerateDSScriptV3(ObjDS, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Proc
                    RetCode = modGenerateV3.GenerateProcScriptV3(ObjProc, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Eng
                    RetCode = modGenerateV3.GenerateEngScriptV3(ObjEng, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel, , True)

            End Select
            FillResults(True)

            Log("********* Script Generation End : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnParseSQD_Click")
        End Try

    End Sub

    Private Sub btnParseOnlySQD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnParseOnlySQD.Click

        Try
            RetCode = New clsRcode

            Select Case ScriptType
                Case enumGenType.DS
                    RetCode = modGenerateV3.GenerateDSScriptV3(ObjDS, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Proc
                    RetCode = modGenerateV3.GenerateProcScriptV3(ObjProc, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel)

                Case enumGenType.Eng
                    RetCode = modGenerateV3.GenerateEngScriptV3(ObjEng, ScriptDirectory, , debugSrc, ShowUID, sourcelevel, targetlevel, True, True)

            End Select
            FillResults(True)

            Log("********* Script Parse End ***********")

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnParseOnlySQD_Click")
        End Try

    End Sub

    Private Sub rbAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAllsrc.CheckedChanged

        If Me.rbAllsrc.Checked = True Then
            Me.rbSelectsrc.Checked = False
            Me.rbFieldsrc.Checked = False
            sourcelevel = enumMappingLevel.ShowAll
        End If

    End Sub

    Private Sub rbSelect_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSelectsrc.CheckedChanged

        If Me.rbSelectsrc.Checked = True Then
            Me.rbAllsrc.Checked = False
            Me.rbFieldsrc.Checked = False
            sourcelevel = enumMappingLevel.ShowDesc
        End If

    End Sub

    Private Sub rbField_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFieldsrc.CheckedChanged

        If Me.rbFieldsrc.Checked = True Then
            Me.rbAllsrc.Checked = False
            Me.rbSelectsrc.Checked = False
            sourcelevel = enumMappingLevel.ShowFld
        End If

    End Sub

    Private Sub rbAlltgt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbAlltgt.CheckedChanged

        If Me.rbAlltgt.Checked = True Then
            Me.rbSelecttgt.Checked = False
            Me.rbFieldtgt.Checked = False
            targetlevel = enumMappingLevel.ShowAll
        End If

    End Sub

    Private Sub rbSelecttgt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbSelecttgt.CheckedChanged

        If Me.rbSelecttgt.Checked = True Then
            Me.rbAlltgt.Checked = False
            Me.rbFieldtgt.Checked = False
            targetlevel = enumMappingLevel.ShowDesc
        End If

    End Sub

    Private Sub rbFieldtgt_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbFieldtgt.CheckedChanged

        If Me.rbFieldtgt.Checked = True Then
            Me.rbAlltgt.Checked = False
            Me.rbSelecttgt.Checked = False
            targetlevel = enumMappingLevel.ShowFld
        End If

    End Sub

    Private Sub btnSQData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSQData.Click

        Try
            Dim Success As String

            '/// Success is the full path to the SQData engine Log
            Success = RunSQData(IO.Path.Combine(txtFolderPath.Text, ObjEng.EngineName & ".PRC"))

            If Success <> "" Then
                Process.Start(Success)
                If System.IO.File.Exists(IO.Path.Combine(txtFolderPath.Text, "Output.dat")) = True Then
                    btnOpenOutput.Enabled = True
                End If
            End If

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnSQData_click")
        End Try

    End Sub

    Private Sub btnOpenOutput_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenOutput.Click

        Try
            '/// Open the output file
            Process.Start(IO.Path.Combine(txtFolderPath.Text, "Output.dat"))

        Catch ex As Exception
            LogError(ex, "frmScriptGen btnOpenOutput_Click")
        End Try

    End Sub

End Class
