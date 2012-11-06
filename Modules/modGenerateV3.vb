Imports System.IO
Public Module modGenerateV3

    '/// Created Sept.-Nov./2007 by Tom Karasch ///
    '/// This module generates the SQD ///
    '/// scripts in Version 3 syntax ///
    '/// and replaces sqdgnsqd.exe ///

#Region "Local Variables"

    Dim ObjThis As clsEngine
    Dim ObjProc As clsTask
    Dim ObjDS As clsDatastore
    Dim ObjSys As clsSystem

    Private semi As String = ";"

    Dim ScriptPath As String
    '/// Parser and Engine Paths
    Dim ParserPath As String
    Dim EnginePath As String
    Dim UseEXEpath As Boolean
    Dim EngEXEpath As String

    Dim doParse As Boolean
    Dim OnlyParse As Boolean
    Dim SQDParse As Boolean
    Dim SourceLevel As enumMappingLevel
    Dim TargetLevel As enumMappingLevel

    'FileStreams and stream writers
    Dim fsSQD As System.IO.FileStream
    Dim objWriteSQD As System.IO.StreamWriter
    Dim fsINL As System.IO.FileStream
    Dim objWriteINL As System.IO.StreamWriter
    Dim fsTMP As System.IO.FileStream
    Dim objWriteTMP As System.IO.StreamWriter
    Dim fsRPT As System.IO.FileStream
    Dim objWriteRPT As System.IO.StreamWriter
    Dim fsERR As System.IO.FileStream
    Dim objWriteERR As System.IO.StreamWriter
    Dim fsPRC As System.IO.FileStream
    Dim objwriteLOG As System.IO.StreamWriter
    Dim fsLOG As System.IO.FileStream
    Dim ResWordStream As System.IO.FileStream
    Dim objReadResWord As System.IO.StreamReader

    'FilePaths
    Dim pathERR As String
    Dim pathRPT As String
    Dim pathPRC As String
    Dim pathTMP As String
    Dim pathSQD As String
    Dim pathINL As String
    Dim pathLOG As String

    Dim strType As String
    Dim DBDalias As String
    Dim DBDfile As String

    Dim PrintUID As Boolean

    Dim ResWords As New Collection

    Dim StrList As New Collection
    Dim dbdList As New Collection
    '/// Added 7/2012
    Dim SQLList As New Collection

    Dim SynNew As Boolean
    Private DBGMap As Boolean = False
    Private OutMsg As Boolean = False

    '/// Added for Multitable SQLserver files 7/2012 by TK
    Private SQLsrcs As New Collection
    Private SQLtgts As New Collection

#End Region

#Region " Main Processes "

    Public Function GenerateEngScriptV3(ByVal EngObj As clsEngine, _
                                        ByVal SavePath As String, _
                                        Optional ByVal NoParse As Boolean = False, _
                                        Optional ByVal debug As Boolean = False, _
                                        Optional ByVal UseUID As Boolean = False, _
                                        Optional ByVal MappingLevel As enumMappingLevel = enumMappingLevel.ShowAll, _
                                        Optional ByVal TgtLevel As enumMappingLevel = enumMappingLevel.ShowAll, _
                                        Optional ByVal ParseOnly As Boolean = False, _
                                        Optional ByVal ParseSQD As Boolean = False, _
                                        Optional ByVal NewSyn As Boolean = True, _
                                        Optional ByVal MapDBG As Boolean = False, _
                                        Optional ByVal UseEXE As Boolean = False) _
                                        As clsRcode

        Log("********* Script Generation Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

        '/// initialize return code object
        Dim RC As clsRcode
        RC = New clsRcode
        RC.HasError = False
        RC.ErrorCount = 0
        RC.HasFatal = False
        RC.Path = ""
        RC.ErrorLocation = enumErrorLocation.NoErrors
        RC.ReturnCode = ""
        RC.ErrorPath = ""
        RC.ParseCode = enumParserReturnCode.NoCode
        RC.ParserPath = ""
        RC.GenType = enumGenType.Eng

        doParse = Not NoParse
        OnlyParse = ParseOnly
        SQDParse = ParseSQD
        PrintUID = UseUID
        SourceLevel = MappingLevel
        TargetLevel = TgtLevel
        SynNew = NewSyn
        DBGMap = MapDBG
        OutMsg = debug
        UseEXEpath = UseEXE


        Try
            ScriptPath = EngObj.ObjSystem.Environment.LocalScriptDir
            ObjThis = EngObj
            ObjSys = EngObj.ObjSystem
            EngObj.LoadMe()
            EngObj.LoadItems()
            EngObj.ObjSystem.LoadMe()
            If EngObj.Connection IsNot Nothing Then
                EngObj.Connection.LoadMe()
            End If

            If UseEXEpath = True Then
                ParserPath = ObjThis.EXEdir & "\sqdparse.exe"
                EnginePath = ObjThis.EXEdir & "\sqdata.exe"
            Else
                If EngObj.EngVersion <> "" Then
                    ParserPath = GetAppPath() & "sqdparse.exe" 'EngObj.EngVersion & "\" & 
                    EnginePath = GetAppPath() & "sqdata.exe"  'EngObj.EngVersion & "\" & 
                Else
                    ParserPath = GetAppPath() & "sqdparse.exe"
                    EnginePath = GetAppPath() & "sqdata.exe"
                End If
            End If

            RC.Path = ScriptPath
            RC.Name = ObjThis.Text

            Call PopulateResWords()
            StrList.Clear()
            dbdList.Clear()

            Log("********* Script File Creation Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")


            '*****************************************************
            '/// create files
            If CreateFiles(RC) = False Then
                GoTo ErrorGoTo
            End If

            '/// No errors so far so set RCpath of SQD file
            RC.Path = pathSQD
            RC.ErrorPath = pathSQD

            If ParseOnly = True Then
                GoTo ParseOnly
            End If
            '*****************************************************
            '/// write Headers
            If prtHeaderComment(RC) = False Then
                GoTo ErrorGoTo
            End If
            '/// Header Comments Take first 9 lines

            '/// write Engine description if any
            If ObjThis.EngineDescription IsNot Nothing Then
                If ObjThis.EngineDescription.Trim <> "" Then
                    wComment(RC, "Engine Description:")
                    wComment(RC, ObjThis.EngineDescription)
                    wComment(RC, "")
                End If
            End If

            '/// Now write Engine Header
            If prtHeader(RC) = False Then
                GoTo ErrorGoTo
            End If

            Log("********* Data Definition Section Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

            '*****************************************************
            '/// write Datastores, Structures, Joins, & Lookups
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        DATA DEFINITION SECTION")
            wComment(RC, "-------------------------------------------------------------")
            If prtDatastores(RC) = False Then
                GoTo ErrorGoTo
            End If

            'If prtJoins(RC) = False Then
            '    GoTo ErrorGoTo
            'End If

            If prtLookups(RC) = False Then
                GoTo ErrorGoTo
            End If

            Log("********* Field Specification Section Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
            '*****************************************************
            '/// Field Specs
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        FIELD SPECIFICATION SECTION")
            wComment(RC, "-------------------------------------------------------------")
            '/// Write Variables
            If prtVAR(RC) = False Then
                GoTo ErrorGoTo
            End If

            '/// write Global Datastore Field Specs
            If prtGlobal(RC) = False Then
                GoTo ErrorGoTo
            End If

            '/// write field attributes
            If prtFA(RC) = False Then
                GoTo ErrorGoTo
            End If

            Log("********* Procedure Section Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
            '*****************************************************
            '/// write Procs
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        PROCEDURE SECTION")
            wComment(RC, "-------------------------------------------------------------")
            If prtProcedures(RC) = False Then
                GoTo ErrorGoTo
            End If

            Log("********* Main Section Section Start : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
            '*****************************************************
            '/// write Main
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        MAIN SECTION")
            wComment(RC, "-------------------------------------------------------------")
            If prtMain(RC) = False Then
                GoTo ErrorGoTo
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

            '/// close file streams ******************************
            objWriteSQD.Close()
            fsSQD.Close()
            objWriteINL.Close()
            fsINL.Close()
            objWriteTMP.Close()
            fsTMP.Close()

            If RC.HasFatal = True Then
                GoTo ErrorGoTo2
            End If

            Log("********* Script Creation Complete : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

ParseOnly:
            '*****************************************************
            If NoParse = False Then
                '/// Run Parser
                If SQDParse = False Then
                    If CallParser(RC) = False Then
                        GoTo ErrorGoTo2
                    End If
                Else
                    If CallParserSQD(RC) = False Then
                        GoTo ErrorGoTo2
                    End If
                End If

            End If

ErrorGoTo2:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate Main")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
            If RC.ErrorLocation = enumErrorLocation.NoErrors Then
                RC.ErrorLocation = enumErrorLocation.ModGenWindows
                RC.ErrorPath = pathSQD
                RC.ObjInode = ObjThis
            End If
        Finally
            If NoParse = False Then
                objWriteRPT.Close()
                fsRPT.Close()
                objWriteERR.Close()
                fsERR.Close()
            End If

            '/// send Return Code Object Back to Main Program
            If RC.HasError = False Then
                RC.ReturnCode = "Script generated Successfully !!"
                RC.ErrorLocation = enumErrorLocation.NoErrors
            Else
                RC.HasError = True
            End If
            GenerateEngScriptV3 = RC

        End Try

    End Function

    Public Function GenerateProcScriptV3(ByVal ProcObj As clsTask, ByVal SavePath As String, Optional ByVal NoParse As Boolean = False, Optional ByVal debug As Boolean = False, Optional ByVal UseUID As Boolean = False, Optional ByVal MappingLevel As enumMappingLevel = enumMappingLevel.ShowAll, Optional ByVal TgtLevel As enumMappingLevel = enumMappingLevel.ShowAll, Optional ByVal NewSyn As Boolean = True) As clsRcode

        '/// initialize return code object
        Dim RC As clsRcode
        RC = New clsRcode
        RC.HasError = False
        RC.ErrorCount = 0
        RC.HasFatal = False
        RC.Path = ""
        RC.ErrorLocation = enumErrorLocation.NoErrors
        RC.ReturnCode = ""
        RC.ErrorPath = ""
        RC.ParseCode = enumParserReturnCode.NoCode
        RC.ParserPath = ""
        RC.GenType = enumGenType.Proc

        doParse = Not NoParse
        PrintUID = UseUID
        SourceLevel = MappingLevel
        TargetLevel = TgtLevel
        SynNew = NewSyn

        Try
            ObjProc = ProcObj
            ScriptPath = ObjProc.Engine.ObjSystem.Environment.LocalScriptDir

            ObjThis = ObjProc.Engine
            ObjThis.LoadMe()

            ObjSys = ObjThis.ObjSystem
            ObjSys.LoadMe()

            RC.Path = ScriptPath
            RC.Name = ObjThis.Text

            Call PopulateResWords()

            Log("********* Script Generation Start *********")

            '*****************************************************
            '/// create files
            If CreateFiles(RC) = False Then
                GoTo ErrorGoTo
            End If
            '/// No errors so far so set RCpath of SQD file
            RC.Path = pathSQD
            RC.ErrorPath = pathSQD

            '*****************************************************
            '/// write Headers
            If prtHeaderComment(RC) = False Then
                GoTo ErrorGoTo
            End If
            '/// Header Comments Take first 9 lines

            '            '/// write Engine description if any
            '            If ObjThis.EngineDescription.Trim <> "" Then
            '                wComment(RC, "Engine Description:")
            '                wComment(RC, ObjThis.EngineDescription)
            '                wComment(RC, "")
            '            End If

            '            '/// Now write Engine Header
            '            If prtHeader(RC) = False Then
            '                GoTo ErrorGoTo
            '            End If

            '            '*****************************************************
            '            '/// write Datastores, Structures, Joins, & Lookups
            '            wComment(RC, "-------------------------------------------------------------")
            '            wComment(RC, "        DATA DEFINITION SECTION")
            '            wComment(RC, "-------------------------------------------------------------")
            '            For Each SrcObj As INode In ObjProc.ObjSources
            '                If SrcObj.Type = NODE_SOURCEDATASTORE Then
            '                    If prtDS(RC, CType(SrcObj, clsDatastore)) = False Then
            '                        GoTo loop1
            '                    End If
            '                End If
            'loop1:      Next

            '            For Each TgtObj As INode In ObjProc.ObjTargets
            '                If TgtObj.Type = NODE_TARGETDATASTORE Then
            '                    If prtDS(RC, CType(TgtObj, clsDatastore)) = False Then
            '                        GoTo loop2
            '                    End If
            '                End If
            'loop2:      Next

            '            For Each SrcObj As INode In ObjProc.ObjSources
            '                If SrcObj.Type = NODE_JOIN Then
            '                    If prtJs(RC, CType(SrcObj, clsTask)) = False Then
            '                        GoTo loop3
            '                    End If
            '                End If
            'loop3:      Next

            '            For Each SrcObj As INode In ObjProc.ObjSources
            '                If SrcObj.Type = NODE_LOOKUP Then
            '                    If prtLUs(RC, CType(SrcObj, clsTask)) = False Then
            '                        GoTo loop4
            '                    End If
            '                End If
            'loop4:      Next

            '            '*****************************************************
            '            '/// Field Specs
            '            wComment(RC, "-------------------------------------------------------------")
            '            wComment(RC, "        FIELD SPECIFICATION SECTION")
            '            wComment(RC, "-------------------------------------------------------------")
            '            '/// Write Variables
            '            If prtVAR(RC) = False Then
            '                GoTo ErrorGoTo
            '            End If

            '            '/// write Global Datastore Field Specs
            '            If prtGlobal(RC) = False Then
            '                GoTo ErrorGoTo
            '            End If

            '            '/// write field attributes
            '            If prtFA(RC) = False Then
            '                GoTo ErrorGoTo
            '            End If

            '*****************************************************
            '/// write Procs
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        PROCEDURE SECTION")
            wComment(RC, "-------------------------------------------------------------")
            If PrtProcs(RC, ObjProc) = False Then
                GoTo ErrorGoTo
            End If


ErrorGoTo:  '/// send returnPath or enumreturncode

            '/// close file streams ******************************
            objWriteSQD.Close()
            fsSQD.Close()
            objWriteINL.Close()
            fsINL.Close()
            objWriteTMP.Close()
            fsTMP.Close()

            If RC.HasFatal = True Then
                GoTo ErrorGoTo2
            End If


            '*****************************************************
            If NoParse = False Then
                '/// Run Parser
                If CallParser(RC) = False Then
                    GoTo ErrorGoTo2
                End If
            End If

ErrorGoTo2:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate Main")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
            If RC.ErrorLocation = enumErrorLocation.NoErrors Then
                RC.ErrorLocation = enumErrorLocation.ModGenWindows
                RC.ErrorPath = pathSQD
                RC.ObjInode = ObjThis
            End If
        Finally
            If NoParse = False Then
                objWriteRPT.Close()
                fsRPT.Close()
                objWriteERR.Close()
                fsERR.Close()
            End If

            '/// send Return Code Object Back to Main Program
            If RC.HasError = False Then
                RC.ReturnCode = "Script generated Successfully !!"
                RC.ErrorLocation = enumErrorLocation.NoErrors
            Else
                RC.HasError = True
            End If

            GenerateProcScriptV3 = RC

        End Try

    End Function

    Public Function GenerateDSScriptV3(ByVal DSObj As clsDatastore, ByVal SavePath As String, Optional ByVal NoParse As Boolean = False, Optional ByVal debug As Boolean = False, Optional ByVal UseUID As Boolean = False, Optional ByVal SrcLevel As enumMappingLevel = enumMappingLevel.ShowAll, Optional ByVal TgtLevel As enumMappingLevel = enumMappingLevel.ShowAll, Optional ByVal NewSyn As Boolean = True) As clsRcode

        '/// initialize return code object
        Dim RC As clsRcode
        RC = New clsRcode
        RC.HasError = False
        RC.ErrorCount = 0
        RC.HasFatal = False
        RC.Path = ""
        RC.ErrorLocation = enumErrorLocation.NoErrors
        RC.ReturnCode = ""
        RC.ErrorPath = ""
        RC.ParseCode = enumParserReturnCode.NoCode
        RC.ParserPath = ""
        RC.GenType = enumGenType.DS

        doParse = Not NoParse
        PrintUID = UseUID
        SourceLevel = SrcLevel
        TargetLevel = TgtLevel
        SynNew = NewSyn

        Try
            ObjDS = DSObj
            ScriptPath = ObjDS.Engine.ObjSystem.Environment.LocalScriptDir
            ObjThis = ObjDS.Engine
            ObjThis.LoadMe()

            ObjSys = ObjThis.ObjSystem
            ObjSys.LoadMe()

            RC.Path = ScriptPath
            RC.Name = ObjThis.Text

            Call PopulateResWords()

            Log("********* Script Generation Start *********")

            '*****************************************************
            '/// create files
            If CreateFiles(RC) = False Then
                GoTo ErrorGoTo
            End If
            '/// No errors so far so set RCpath of SQD file
            RC.Path = pathSQD
            RC.ErrorPath = pathSQD

            '*****************************************************
            '/// write Headers
            If prtHeaderComment(RC) = False Then
                GoTo ErrorGoTo
            End If
            '/// Header Comments Take first 9 lines

            '/// write Engine description if any
            If ObjThis.EngineDescription.Trim <> "" Then
                wComment(RC, "Engine Description:")
                wComment(RC, ObjThis.EngineDescription)
                wComment(RC, "")
            End If

            '/// Now write Engine Header
            If prtHeader(RC) = False Then
                GoTo ErrorGoTo
            End If

            '*****************************************************
            '/// write Datastores, Structures, Joins, & Lookups
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        DATA DEFINITION SECTION")
            wComment(RC, "-------------------------------------------------------------")
            If prtDS(RC, DSObj) = False Then
                GoTo ErrorGoTo
            End If

            '*****************************************************
            '/// Field Specs
            wComment(RC, "-------------------------------------------------------------")
            wComment(RC, "        FIELD SPECIFICATION SECTION")
            wComment(RC, "-------------------------------------------------------------")
            '/// Write Variables

            '/// write Global Datastore Field Specs
            If prtGlobal(RC) = False Then
                GoTo ErrorGoTo
            End If

            '/// write field attributes
            If prtFA(RC) = False Then
                GoTo ErrorGoTo
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

            '/// close file streams ******************************
            objWriteSQD.Close()
            fsSQD.Close()
            objWriteINL.Close()
            fsINL.Close()
            objWriteTMP.Close()
            fsTMP.Close()

            If RC.HasFatal = True Then
                GoTo ErrorGoTo2
            End If


            '*****************************************************
            If NoParse = False Then
                '/// Run Parser
                If CallParser(RC) = False Then
                    GoTo ErrorGoTo2
                End If
            End If

ErrorGoTo2:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate Main")
            RC.HasError = True
            RC.ErrorCount += 1
            If RC.ReturnCode = "" Then
                RC.ReturnCode = ex.Message
            End If
            If RC.ErrorLocation = enumErrorLocation.NoErrors Then
                RC.ErrorLocation = enumErrorLocation.ModGenWindows
                RC.ErrorPath = pathSQD
                RC.ObjInode = ObjThis
            End If
        Finally
            If NoParse = False Then
                objWriteRPT.Close()
                fsRPT.Close()
                objWriteERR.Close()
                fsERR.Close()
            End If

            '/// send Return Code Object Back to Main Program
            If RC.HasError = False Then
                RC.ReturnCode = "Script generated Successfully !!"
                RC.ErrorLocation = enumErrorLocation.NoErrors
            Else
                RC.HasError = True
            End If
            GenerateDSScriptV3 = RC

        End Try

    End Function

#End Region

#Region " Parser function "

    '//// Function to parse scripts and capture output to files
    Function CallParser(ByRef rc As clsRcode) As Boolean

        Try
            Dim FORargs As String = String.Format("{0}{1}", Quote(GetFileNameFromPath(pathINL), """") & " ", _
              Quote(GetFileNameFromPath(pathPRC), """") & " LIST=ALL A B C D E F G H 1")
            ' >" & Quote(pathRPT, """") & " 2>" & Quote(pathERR, """"))
            Dim args As String = FORargs
            Dim si As New ProcessStartInfo()

            Log("*** Parser Run Started : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

            '/// Set all Process Start Information
            si.WorkingDirectory = GetDirFromPath(pathPRC)  'GetAppPath() '
            si.Arguments = args

            si.FileName = ParserPath

            si.UseShellExecute = False
            si.CreateNoWindow = True
            '// Output will be redirected to .RPT file
            si.RedirectStandardOutput = True
            si.RedirectStandardError = True

            Using myProcess As New System.Diagnostics.Process()
                myProcess.StartInfo = si

                Log("*** " & si.FileName & args)

                '//Create a new process to parse Script file and dump to RPT file and Error File
                myProcess.Start()

                Dim OutStr As String = ""
                Dim ErrStr As String = ""

                '/// split output into multiple threads and capture each output stream into separate buffers
                OutputToEnd(myProcess, OutStr, ErrStr)

                ''//wait until task is done
                myProcess.WaitForExit()

                Select Case myProcess.ExitCode
                    Case 8
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Parser Returned: Script Generation Error" ' & Chr(13)
                        '&  "Would you like to see the report?"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case 4
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Warning
                        rc.ReturnCode = "Parser Returned: Script Generation Warning"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case 1
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Parser Returned: There is a problem with the PATH"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case Is > 0
                        rc.HasError = True
                        rc.ParserPath = GetAppLog() & "sqdparse.log"
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Script generated with errors,"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = Quote(GetAppLog() & "sqdparse.log", """")
                        rc.ParserRC = myProcess.ExitCode

                    Case 0
                        rc.ParserPath = pathINL
                        rc.ParseCode = enumParserReturnCode.OK
                        rc.ReturnCode = "Script generated and Parsed Successfully !!"
                        rc.ErrorLocation = enumErrorLocation.NoErrors
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode
                        Log("*** Script file saved at : " & pathINL)

                    Case Else
                        rc.HasError = True
                        Log("return code >> " & myProcess.ExitCode.ToString & _
                            Chr(13) & "return path >> " & pathSQD)
                End Select


                objWriteRPT.Write(ErrStr)
                objWriteERR.Write(ErrStr)
                objWriteRPT.Write(OutStr)


                Log("*** Parser Return Code = " & myProcess.ExitCode & " *********")
                Log("*** Parser Report file saved at : " & pathRPT)
                Log("*** Parser Run Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                myProcess.Close()

            End Using

ErrorGoTo:  '/// send returnPath or enumreturncode


        Catch ex As Exception
            LogError(ex, "modGenerate CallParser")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.ReturnCode = ex.Message
            rc.ErrorPath = pathRPT
            rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
            rc.ParserPath = ""
            rc.ParseCode = enumParserReturnCode.NoCode
            rc.ObjInode = ObjThis
        End Try

        CallParser = Not rc.HasFatal

    End Function

    '//// Function to parse scripts and capture output to files
    Function CallParserSQD(ByRef rc As clsRcode) As Boolean

        Try
            Dim FORargs As String = String.Format("{0}{1}", Quote(GetFileNameFromPath(pathSQD), """") & _
            " ", Quote(GetFileNameFromPath(pathPRC), """") & " LIST=ALL A B C D E F G H 1")
            '' >" & Quote(pathRPT, """") & " 2>" & Quote(pathERR, """"))
            Dim args As String = FORargs
            Dim si As New ProcessStartInfo()

            Log("*** Parser Run Started : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

            '/// Set all Process Start Information
            si.WorkingDirectory = GetDirFromPath(pathPRC)  'GetAppPath()
            si.Arguments = args

            si.FileName = ParserPath

            si.UseShellExecute = False
            si.CreateNoWindow = True
            '// Output will be redirected to .RPT file
            si.RedirectStandardOutput = True
            si.RedirectStandardError = True

            Using myProcess As New System.Diagnostics.Process()
                myProcess.StartInfo = si

                Log("*** " & si.FileName & args)

                '//Create a new process to parse Script file and dump to RPT file and Error File
                myProcess.Start()

                Dim OutStr As String = ""
                Dim ErrStr As String = ""

                '/// split output into multiple threads and capture each output stream into separate buffers
                OutputToEnd(myProcess, OutStr, ErrStr)

                ''//wait until task is done
                myProcess.WaitForExit()

                Select Case myProcess.ExitCode
                    Case 8
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Parser Returned: Script Generation Error" ' & Chr(13)
                        '&  "Would you like to see the report?"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case 4
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Warning
                        rc.ReturnCode = "Parser Returned: Script Generation Warning"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case 1
                        rc.HasError = True
                        rc.ParserPath = pathRPT
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Parser Returned: There is a problem with the PATH"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode

                    Case Is > 0
                        rc.HasError = True
                        rc.ParserPath = GetAppLog() & "sqdparse.log"
                        rc.ParseCode = enumParserReturnCode.Failed
                        rc.ReturnCode = "Script generated with errors,"
                        rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
                        rc.Path = Quote(GetAppLog() & "sqdparse.log", """")
                        rc.ParserRC = myProcess.ExitCode

                    Case 0
                        rc.ParserPath = pathSQD
                        rc.ParseCode = enumParserReturnCode.OK
                        rc.ReturnCode = "Script generated Successfully !!"
                        rc.ErrorLocation = enumErrorLocation.NoErrors
                        rc.Path = pathSQD
                        rc.ParserRC = myProcess.ExitCode
                        Log("*** Script file saved at : " & pathSQD)

                    Case Else
                        rc.HasError = True
                        Log("return code >> " & myProcess.ExitCode.ToString & Chr(13) & "return path >> " & pathSQD)
                End Select

                'Error First so Parser Header is at Top

                objWriteRPT.Write(ErrStr)
                objWriteERR.Write(ErrStr)
                objWriteRPT.Write(OutStr)


                Log("*** Parser Return Code = " & myProcess.ExitCode & " *********")
                Log("*** Parser Report file saved at : " & pathRPT)
                Log("*** Parser Run Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                myProcess.Close()

            End Using

ErrorGoTo:  '/// send returnPath or enumreturncode


        Catch ex As Exception
            LogError(ex, "modGenerate CallParserSQD")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.ReturnCode = ex.Message
            rc.ErrorPath = pathRPT
            rc.ErrorLocation = enumErrorLocation.ModGenSQDParse
            rc.ParserPath = ""
            rc.ParseCode = enumParserReturnCode.NoCode
            rc.ObjInode = ObjThis
        End Try

        CallParserSQD = Not rc.HasFatal

    End Function

#End Region

#Region " Create Output Files and Write Headers "

    Function CreateFiles(ByRef rc As clsRcode) As Boolean

        Dim ScriptName As String = ""

        Select Case rc.GenType
            Case enumGenType.DS
                ScriptName = "DS-" & ObjDS.DatastoreName
            Case enumGenType.Eng
                ScriptName = ObjThis.EngineName
            Case enumGenType.Proc
                ScriptName = "P-" & ObjProc.TaskName
        End Select

        pathSQD = ScriptPath & "\" & ScriptName & ".sqd"
        pathINL = ScriptPath & "\" & ScriptName & ".inl"
        pathTMP = ScriptPath & "\" & ScriptName & ".tmp"
        rc.SQDPath = pathSQD
        rc.INLPath = pathINL
        rc.TMPPath = pathTMP

        If OnlyParse = False Then
            Try
                '//Create new file and filestream for SQD file
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".sqd") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".sqd")
                End If
                fsSQD = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".sqd")
                pathSQD = fsSQD.Name
                rc.SQDPath = pathSQD
                objWriteSQD = New System.IO.StreamWriter(fsSQD)

            Catch ex As Exception
                LogError(ex, "modGenerate makeSQD")
                rc.HasError = True
                rc.ErrorCount += 1
                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".sqd"
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                GoTo ErrorGoTo
            End Try

            Try
                '//Create new file and filestream for INL file
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".inl") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".inl")
                End If
                fsINL = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".inl")
                pathINL = fsINL.Name
                rc.INLPath = pathINL
                objWriteINL = New System.IO.StreamWriter(fsINL)

            Catch ex As Exception
                LogError(ex, "modGenerate makeINL")
                rc.HasError = True
                rc.ErrorCount += 1
                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".inl"
                GoTo ErrorGoTo
            End Try

            '//Create New file and filestream for TMP file
            Try
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".tmp") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".tmp")
                End If
                fsTMP = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".tmp")
                pathTMP = fsTMP.Name
                rc.TMPPath = pathTMP
                objWriteTMP = New System.IO.StreamWriter(fsTMP)

            Catch ex As Exception
                LogError(ex, "modGenerate makeTMP")
                rc.HasError = True
                rc.ErrorCount += 1
                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".tmp"
                GoTo ErrorGoTo
            End Try
        End If
        '/// Create the Report file that the parser will generate
        If doParse = True Then

            Try
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".rpt") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".rpt")
                End If
                fsRPT = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".rpt")
                pathRPT = fsRPT.Name
                rc.RPTPath = pathRPT
                objWriteRPT = New System.IO.StreamWriter(fsRPT)


            Catch ex As Exception
                LogError(ex, "modGenerate makeRPT")
                rc.HasError = True
                rc.ErrorCount += 1

                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".rpt"
                GoTo ErrorGoTo
            End Try

            '/// Create the Error file that the parser will generate if it causes Windows Errors
            Try
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".ERR") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".ERR")
                End If
                fsERR = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".ERR")
                pathERR = fsERR.Name
                objWriteERR = New System.IO.StreamWriter(fsERR)


            Catch ex As Exception
                LogError(ex, "modGenerate makeERR")
                rc.HasError = True
                rc.ErrorCount += 1
                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".ERR"
                GoTo ErrorGoTo
            End Try

            '//// Create the Parse file (.PRC) that the Parser will Generate
            Try
                If System.IO.File.Exists(ScriptPath & "\" & ScriptName & ".prc") Then
                    System.IO.File.Delete(ScriptPath & "\" & ScriptName & ".prc")
                End If
                fsPRC = System.IO.File.Create(ScriptPath & "\" & ScriptName & ".prc")
                pathPRC = fsPRC.Name
                rc.PRCPath = pathPRC

            Catch ex As Exception
                LogError(ex, "modGenerate makePRC")
                rc.HasError = True
                rc.ErrorCount += 1
                rc.HasFatal = True
                rc.ObjInode = ObjThis
                rc.ReturnCode = ex.Message
                rc.ErrorLocation = enumErrorLocation.ModGenFileCreation
                rc.ErrorPath = ScriptPath & "\" & ScriptName & ".prc"
                GoTo ErrorGoTo
            End Try
        End If
        '----------------------------------------------------------------------
        '//// Create the Mapping Log file (.LOG) that will populate with 
        '//// mapping sources and targets and their types
        '/// *** THIS IS FOR DEBUGGING ***
        'Try
        '    If System.IO.File.Exists(ScriptPath & "\" & ObjThis.EngineName & "MapType.log") Then
        '        System.IO.File.Delete(ScriptPath & "\" & ObjThis.EngineName & "MapType.log")
        '    End If
        '    fsLOG = System.IO.File.Create(ScriptPath & "\" & ObjThis.EngineName & "MapType.log")
        '    pathLOG = fsLOG.Name
        '    objwriteLOG = New System.IO.StreamWriter(fsLOG)

        'Catch ex As Exception
        '    LogError(ex, "modGenerate makeLOG")
        '    rc.HasError = True
        '    rc.ObjInode = ObjThis
        '    rc.ReturnCode = ex.Message
        '    rc.ErrorLocation = enumErrorLocation.FileCreation
        '    rc.ErrorPath = ScriptPath & "\" & ObjThis.EngineName & "MapType.log"
        '    GoTo ErrorGoTo
        'End Try
        '------------------------------------------------------------------------

        '/// Jump to here on error
ErrorGoTo:
        If doParse = True Then
            fsPRC.Close()
            'fsRPT.Close()
        End If

        '/// send returnPath or enumreturncode
        CreateFiles = Not rc.HasFatal

    End Function

    Function prtHeaderComment(ByRef rc As clsRcode) As Boolean

        Try
            Dim i As Integer
            For i = 0 To 7
                Select Case i
                    Case 0, 7
                        wHeaderComments(rc, i, "-------------------------------------------")
                    Case 1, 6
                        wHeaderComments(rc, i, "--                                       --")
                    Case 2
                        wHeaderComments(rc, i, "--  GENERATED ON:  ")
                    Case 3
                        wHeaderComments(rc, i, "--       PROJECT:  ", ObjThis.Project.ProjectName)
                    Case 4
                        wHeaderComments(rc, i, "--   ENVIRONMENT:  ", ObjThis.ObjSystem.Environment.EnvironmentName)
                    Case 5
                        wHeaderComments(rc, i, "--        SYSTEM:  ", ObjThis.ObjSystem.SystemName)
                End Select
            Next

        Catch ex As Exception
            LogError(ex, "modGenerate prtHeader")
            rc.HasError = True
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenHead
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
        End Try

        prtHeaderComment = Not rc.HasFatal

    End Function

    Function prtHeader(ByRef rc As clsRcode) As Boolean

        Try
            wHeader(rc, 1, "JOBNAME ")
            If ObjThis.ReportFile <> "" Then
                wHeader(rc, 2, "REPORT " & ObjThis.ReportFile & " EVERY ")
            Else
                If ObjThis.ReportEvery <> 0 Then
                    wHeader(rc, 2, "REPORT EVERY ")
                End If
            End If
            If ObjThis.CommitEvery <> 0 Then
                wHeader(rc, 3, "COMMIT EVERY ")
            End If
            If ObjThis.DateFormat IsNot Nothing Then
                If ObjThis.DateFormat.Trim <> "" Then
                    wHeader(rc, 4, "DATEFORMAT ")
                End If
            End If

            wHeader(rc, 5, "RDBMS ")

        Catch ex As Exception
            LogError(ex, "modGenerate prtHeader")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Printing Header Info"
            Log(rc.LocalErrorMsg)
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenHead
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
        End Try

        prtHeader = Not rc.HasFatal

    End Function

#End Region

#Region " Object Functions "

    Function prtDatastores(ByRef rc As clsRcode) As Boolean

        Try
            SQLsrcs.Clear()
            SQLtgts.Clear()

            '/// First Sources
            For Each ds As clsDatastore In ObjThis.Sources
                ds.LoadMe()
                ds.LoadItems()

                If ds.UseMTD = True And ds.MTDfile <> "" Then
                    'New logic for SQLserver MTD
                    SQLsrcs.Add(ds, ds.DatastoreName)
                Else
                    '// original logic
                    If ds.IsLookUp = False Then
                        If prtDS(rc, ds) = False Then
                            GoTo ErrorGoTo
                        End If
                    Else
                        If prtLUDS(rc, ds) = False Then
                            GoTo ErrorGoTo
                        End If
                    End If
                End If
            Next 'Source

            '/// Now process DS with MTD added 7/2012 by TK
            If SQLsrcs.Count > 0 Then
                If prtSrcSQLserverDS(rc) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            '/// Then Targets
            For Each ds As clsDatastore In ObjThis.Targets
                'wComment(rc, "---------------------------------------------------------------------")
                ds.LoadMe()
                ds.LoadItems()

                If ds.UseMTD = True And ds.MTDfile <> "" Then
                    'New logic for SQLserver MTD
                    SQLtgts.Add(ds, ds.DatastoreName)
                Else
                    '// original logic
                    If prtDS(rc, ds) = False Then
                        GoTo ErrorGoTo
                    End If
                End If
            Next 'Target

            If SQLtgts.Count > 0 Then
                If prtTgtSQLserverDS(rc) = False Then
                    GoTo ErrorGoTo
                End If
            End If


ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtDatastores")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Occured Writing Datastores"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenDS
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtDatastores = Not rc.HasFatal

    End Function

    Function prtDS(ByRef rc As clsRcode, Optional ByVal InputDS As clsDatastore = Nothing) As Boolean

        'Dim j As Integer
        Dim ds As clsDatastore
        'Dim dssel As clsDSSelection
        Dim str As clsStructure

        Try
            If InputDS Is Nothing Then
                ds = ObjDS
            Else
                ds = InputDS
            End If

            If ds.DatastoreType = enumDatastore.DS_INCLUDE Then
                If prtInclude(rc, ds) = False Then
                    GoTo ErrorGoTo
                Else
                    GoTo ErrorGoTo
                End If
            End If

            
            '/// Normal Datastore dump
            '// loop through all structures in Datastore and write to files
            For Each dssel As clsDSSelection In ds.ObjSelections
                str = dssel.ObjStructure
                str.LoadMe()

                If StrList.Contains(str.StructureName) = False Then
                    StrList.Add(str, str.StructureName)
                    '/// if IMSDBD write IMSDBname line before structures
                    If str.StructureType = enumStructure.STRUCT_COBOL_IMS Then   'j = 0 And
                        If dbdList.Contains(str.IMSDBName) = False Then
                            dbdList.Add(str, str.IMSDBName)
                            If wIMSprefix(rc, str, dssel) = False Then
                                GoTo ErrorGoTo
                            End If
                        End If
                    End If
                    If str.StructureDescription IsNot Nothing Then
                        If str.StructureDescription.Trim <> "" Then
                            wComment(rc, str.StructureDescription)
                        End If
                    End If

                    '/// write Structures
                    If wStructSQD(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    If wStructINL(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    If wStructTMP(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    '/// if IMSDBD then write Segment Desc after each structure
                    If str.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                        If wIMSsuffix(rc, str) = False Then
                            GoTo ErrorGoTo
                        End If
                    End If
                    '/// list all field Descriptions in each structure
                    For Each fld As clsField In str.ObjFields
                        If fld.FieldDesc IsNot Nothing Then
                            If fld.FieldDesc.Trim <> "" Then
                                wComment(rc, str.StructureName & "." & fld.FieldName & " has description: " & fld.FieldDesc)
                            End If
                        End If
                    Next
                    wComment(rc, " ")
                End If
            Next
            '/// Do Unit of Work description, if necessary on Sources Only
            If ds.DsUOW.Trim <> "" And ds.DatastoreType = enumDatastore.DS_DB2CDC Then
                If wDSuow(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            '/// If Exception Datastore, then write Description for it here
            If ds.ExceptionDatastore.Trim <> "" Then
                If wDummy(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            '/// If Restart Datastore, then write Restart Datastore Here
            'If ds.Restart.Trim <> "" And ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
            '    If wDummyReSTRT(rc, ds) = False Then
            '        GoTo ErrorGoTo
            '    End If
            'End If
            '/// now write Datastore Description
            If ds.DatastoreType = enumDatastore.DS_INCLUDE Then
                If prtInclude(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            Else
                If wDS(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            If wDSattribException(rc, ds) = False Then
                GoTo ErrorGoTo
            End If

            If wDSkey(rc, ds) = False Then
                GoTo ErrorGoTo
            End If

            '/// now write datastore special attributes
            If wDSattrib(rc, ds) = False Then
                GoTo ErrorGoTo
            End If

            If ds.DsDirection = DS_DIRECTION_TARGET Then
                If wDSsuffix(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            wSemiLine(rc)

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtDS")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Occured Writing Datastores"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenDS
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtDS = Not rc.HasFatal

    End Function

    '// New for MSsqlServer MTD files 7/2012 by TK
    Function prtSrcSQLserverDS(ByRef rc As clsRcode) As Boolean

        Try
            For Each ds As clsDatastore In SQLsrcs
                Dim sqlMTDfile As String = ""

                '/// First Get the MTD file Name
                If ds.MTDfile <> "" Then
                    sqlMTDfile = ds.MTDfile
                Else
                    GoTo ErrorGoTo
                End If

                '/// Now print the MTD file to the script using "sqlDescList" for the Alias Statement
                Dim FORstr1SQD As String = String.Format("{0}{1}", "DESCRIPTION MSSQL ", Quote(sqlMTDfile))
                Dim FORstr1INL As String = String.Format("{0}", "DESCRIPTION MSSQL ")
                Dim PSix1INL As String = String.Format("{0}", "/+")
                Dim PSix2INL As String = String.Format("{0}", "+/")

                '/// Read in MTD and Print to INL, Print first lines for SQD
                Dim objReadStr As System.IO.StreamReader
                Dim InStream As System.IO.FileStream
                Dim TempString As String

                wBlankLine(rc)
                'sqd
                objWriteSQD.WriteLine(FORstr1SQD)
                AddToLineNo(rc, False, , False)
                'inl
                objWriteINL.WriteLine(FORstr1INL)
                AddToLineNo(rc, False, , False)
                objWriteINL.WriteLine(PSix1INL)
                AddToLineNo(rc, False, , False)


                If System.IO.File.Exists(sqlMTDfile) = True Then
                    InStream = System.IO.File.OpenRead(sqlMTDfile)
                    objReadStr = New System.IO.StreamReader(InStream)
                Else
                    rc.HasError = True
                    'rc.ObjInode = struct
                    rc.ErrorLocation = enumErrorLocation.ModGenStruct
                    rc.ReturnCode = "File " & rc.Path & " does not exist"
                    rc.ErrorPath = pathINL
                    GoTo ErrorGoTo
                End If
                '/// now read and write the structure file
                While (objReadStr.Peek() > -1)
                    TempString = objReadStr.ReadLine

                    objWriteINL.WriteLine(TempString)

nextTempstr:        AddToLineNo(rc, False, , False)
                End While

                '/// Write closing to INL
SubstGoTo:      objWriteINL.WriteLine(PSix2INL)
                AddToLineNo(rc, False, , False)

                '/// Write ALIAS Statement
                Dim count As Integer = 1

                For Each dssel As clsDSSelection In ds.ObjSelections
                    Dim dmlPath As String = dssel.ObjStructure.fPath1
                    Dim DescAlias As String = dssel.SelectionName
                    Dim FORaliasFirst As String = String.Format("{0,14}{1}{2}{3}", "ALIAS(", dmlPath, " AS ", DescAlias)
                    Dim FORaliasmid As String = String.Format("{0,14}{1}{2}{3}", ",", dmlPath, " AS ", DescAlias)
                    Dim FORaliasLast As String = String.Format("{0,14}{1}{2}{3}{4}", ",", dmlPath, " AS ", DescAlias, ");")

                    If count = 1 Then
                        objWriteINL.WriteLine(FORaliasFirst)
                        AddToLineNo(rc, False, , False)
                        objWriteSQD.WriteLine(FORaliasFirst)
                        AddToLineNo(rc, False, , False)
                    ElseIf count < ds.ObjSelections.Count Then
                        objWriteINL.WriteLine(FORaliasmid)
                        AddToLineNo(rc, False, , False)
                        objWriteSQD.WriteLine(FORaliasmid)
                        AddToLineNo(rc, False, , False)
                    Else
                        objWriteINL.WriteLine(FORaliasLast)
                        AddToLineNo(rc, False, , False)
                        objWriteSQD.WriteLine(FORaliasLast)
                        AddToLineNo(rc, False, , False)
                    End If
                    count += 1
                Next
                wBlankLine(rc)
                wComment(rc, " ")
                '/// Normal Datastore dump
                '/// If Exception Datastore, then write Description for it here
                If ds.ExceptionDatastore.Trim <> "" Then
                    If wDummy(rc, ds) = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                '/// now write Datastore Description
                If wDS(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// attributes exceptions
                If wDSattribException(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// Keys
                If wDSkey(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// now write datastore special attributes
                If wDSattrib(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// Last write suffix
                If wDSsuffix(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If

                wSemiLine(rc)
                wBlankLine(rc)
            Next

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtSrcSQLserverDS")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Occured Writing MSSQLserver Datastores"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenDS
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtSrcSQLserverDS = Not rc.HasFatal

    End Function

    '// New for MSsqlServer MTD files 7/2012 by TK
    Function prtTgtSQLserverDS(ByRef rc As clsRcode) As Boolean

        Try
            Dim sqlDescList As New Collection
            Dim sqlMTDfile As String = ""

            '/// First Get the MTD file Name
            For Each ds As clsDatastore In SQLtgts
                If ds.MTDfile <> "" Then
                    sqlMTDfile = ds.MTDfile
                    Exit For
                End If
            Next

            '/// Next Build list of Target Descriptions in Target Datastores
            For Each ds As clsDatastore In SQLtgts
                For Each dssel As clsDSSelection In ds.ObjSelections
                    sqlDescList.Add(dssel, dssel.SelectionName)
                Next
            Next

            '/// Now print the MTD file to the script using "sqlDescList" for the Alias Statement
            Dim FORstr1SQD As String = String.Format("{0}{1}", "DESCRIPTION MSSQL ", Quote(sqlMTDfile))
            Dim FORstr1INL As String = String.Format("{0}", "DESCRIPTION MSSQL ")
            Dim PSix1INL As String = String.Format("{0}", "/+")
            Dim PSix2INL As String = String.Format("{0}", "+/")

            '/// Read in MTD and Print to INL, Print first lines for SQD
            Dim objReadStr As System.IO.StreamReader
            Dim InStream As System.IO.FileStream
            Dim TempString As String

            wBlankLine(rc)
            wComment(rc, " ")
            'sqd
            objWriteSQD.WriteLine(FORstr1SQD)
            AddToLineNo(rc, False, , False)
            'inl
            objWriteINL.WriteLine(FORstr1INL)
            AddToLineNo(rc, False, , False)
            objWriteINL.WriteLine(PSix1INL)
            AddToLineNo(rc, False, , False)


            If System.IO.File.Exists(sqlMTDfile) = True Then
                InStream = System.IO.File.OpenRead(sqlMTDfile)
                objReadStr = New System.IO.StreamReader(InStream)
            Else
                rc.HasError = True
                'rc.ObjInode = struct
                rc.ErrorLocation = enumErrorLocation.ModGenStruct
                rc.ReturnCode = "File " & rc.Path & " does not exist"
                rc.ErrorPath = pathINL
                GoTo ErrorGoTo
            End If
            '/// now read and write the structure file
            While (objReadStr.Peek() > -1)
                TempString = objReadStr.ReadLine
                
                objWriteINL.WriteLine(TempString)

nextTempstr:    AddToLineNo(rc, False, , False)
            End While

            '/// Write closing to INL
SubstGoTo:  objWriteINL.WriteLine(PSix2INL)
            AddToLineNo(rc, False, , False)

            '/// Write ALIAS Statement
            Dim count As Integer = 1

            For Each dssel As clsDSSelection In sqlDescList
                Dim dmlPath As String = dssel.ObjStructure.fPath1
                Dim DescAlias As String = dssel.SelectionName
                Dim FORaliasFirst As String = String.Format("{0,14}{1}{2}{3}", "ALIAS(", dmlPath, " AS ", DescAlias)
                Dim FORaliasmid As String = String.Format("{0,14}{1}{2}{3}", ",", dmlPath, " AS ", DescAlias)
                Dim FORaliasLast As String = String.Format("{0,14}{1}{2}{3}{4}", ",", dmlPath, " AS ", DescAlias, ");")

                If count = 1 Then
                    objWriteINL.WriteLine(FORaliasFirst)
                    AddToLineNo(rc, False, , False)
                    objWriteSQD.WriteLine(FORaliasFirst)
                    AddToLineNo(rc, False, , False)
                ElseIf count < sqlDescList.Count Then
                    objWriteINL.WriteLine(FORaliasmid)
                    AddToLineNo(rc, False, , False)
                    objWriteSQD.WriteLine(FORaliasmid)
                    AddToLineNo(rc, False, , False)
                Else
                    objWriteINL.WriteLine(FORaliasLast)
                    AddToLineNo(rc, False, , False)
                    objWriteSQD.WriteLine(FORaliasLast)
                    AddToLineNo(rc, False, , False)
                End If
                count += 1
            Next
            wBlankLine(rc)
            '/// Normal Datastore dump
            For Each ds As clsDatastore In SQLtgts
                wComment(rc, " ")
                '/// If Exception Datastore, then write Description for it here
                If ds.ExceptionDatastore.Trim <> "" Then
                    If wDummy(rc, ds) = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                '/// now write Datastore Description
                If wDS(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// attributes exceptions
                If wDSattribException(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// Keys
                If wDSkey(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// now write datastore special attributes
                If wDSattrib(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                '/// Last write suffix
                If wDSsuffix(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If

                wSemiLine(rc)
                wBlankLine(rc)
            Next

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtTgtSQLserverDS")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Occured Writing MSSQLserver Datastores"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenDS
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtTgtSQLserverDS = Not rc.HasFatal

    End Function

    Function prtLUDS(ByRef rc As clsRcode, Optional ByVal InputDS As clsDatastore = Nothing) As Boolean

        Dim j As Integer
        Dim ds As clsDatastore
        Dim dssel As clsDSSelection
        Dim str As clsStructure

        Try
            If InputDS Is Nothing Then
                ds = ObjDS
            Else
                ds = InputDS
            End If
            If ds.DatastoreType = enumDatastore.DS_INCLUDE Then
                If prtInclude(rc, ds) = False Then
                    GoTo ErrorGoTo
                Else
                    GoTo ErrorGoTo
                End If
            End If
            '// loop through all structures in Datastore and write to files
            For j = 0 To ds.ObjSelections.Count - 1
                dssel = CType(ds.ObjSelections(j), clsDSSelection)
                str = dssel.ObjStructure
                If StrList.Contains(str.StructureName) = False Then
                    StrList.Add(str, str.StructureName)
                    '/// if IMSDBD write IMSDBname line before structures
                    If str.StructureType = enumStructure.STRUCT_COBOL_IMS Then   'j = 0 And
                        If dbdList.Contains(str.IMSDBName) = False Then
                            dbdList.Add(str, str.IMSDBName)
                            If wIMSprefix(rc, str, dssel) = False Then
                                GoTo ErrorGoTo
                            End If
                        End If
                    End If
                    If str.StructureDescription IsNot Nothing Then
                        If str.StructureDescription.Trim <> "" Then
                            wComment(rc, str.StructureDescription)
                        End If
                    End If

                    '/// write Structures
                    If wStructSQD(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    If wStructINL(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    If wStructTMP(rc, str, dssel) = False Then
                        GoTo ErrorGoTo
                    End If
                    '/// if IMSDBD then write Segment Desc after each structure
                    If str.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                        If wIMSsuffix(rc, str) = False Then
                            GoTo ErrorGoTo
                        End If
                    End If
                    '/// list all field Descriptions in each structure
                    For Each fld As clsField In str.ObjFields
                        If fld.FieldDesc IsNot Nothing Then
                            If fld.FieldDesc.Trim <> "" Then
                                wComment(rc, str.StructureName & "." & fld.FieldName & " has description: " & fld.FieldDesc)
                            End If
                        End If
                    Next
                    wComment(rc, " ")
                End If
            Next
            '/// Do Unit of Work description, if necessary on Sources Only
            If ds.DsUOW.Trim <> "" And ds.DatastoreType = enumDatastore.DS_DB2CDC Then
                If wDSuow(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            '/// If Exception Datastore, then write Description for it here
            If ds.ExceptionDatastore.Trim <> "" Then
                If wDummy(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            '/// now write Datastore Description
            If ds.DatastoreType = enumDatastore.DS_INCLUDE Then
                If prtInclude(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            Else
                If wDS(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            '////////////////////////////////////////////////////////////////////
            If wDSkey(rc, ds, True) = False Then
                GoTo ErrorGoTo
            End If
            '/////////////////////////////////////////////////////////////////////
            If wLookDS(rc, ds) = False Then
                GoTo ErrorGoTo
            End If

            If wDSkey(rc, ds) = False Then
                GoTo ErrorGoTo
            End If
            '/// now write datastore special attributes
            If wDSattrib(rc, ds) = False Then
                GoTo ErrorGoTo
            End If
            If ds.DsDirection = DS_DIRECTION_TARGET Then
                If wDSsuffix(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            wSemiLine(rc)

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtDSLU")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error Occured Writing Datastores"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenDS
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtLUDS = Not rc.HasFatal

    End Function

    Function prtGlobal(ByRef rc As clsRcode) As Boolean

        Try
            For Each ds As clsDatastore In ObjThis.Sources
                If ds.ExtTypeChar <> "" Then
                    If wGlobal(rc, ds, "EXTTYPECHAR") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                If ds.IfNullChar <> "" Then
                    If wGlobal(rc, ds, "IFNULLCHAR") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                If ds.InValidChar <> "" Then
                    If wGlobal(rc, ds, "INVALIDCHAR") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                If ds.ExtTypeNum <> "" Then
                    If wGlobal(rc, ds, "EXTTYPENUM") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                If ds.IfNullNum <> "" Then
                    If wGlobal(rc, ds, "IFNULLNUM") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
                If ds.InValidNum <> "" Then
                    If wGlobal(rc, ds, "INVALIDNUM") = False Then
                        GoTo ErrorGoTo
                    End If
                End If
            Next

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtGlobal")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Global Variables"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenStruct
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtGlobal = Not rc.HasFatal

    End Function

    Function prtFA(ByRef rc As clsRcode) As Boolean

        Try
            Dim i As Integer
            Dim FldName As String = ""

            For Each ds As clsDatastore In ObjThis.Sources
                For Each dsSel As clsDSSelection In ds.ObjSelections
                    dsSel.LoadMe()

                    For Each Fld As clsField In dsSel.DSSelectionFields
                        'dsSel.LoadFieldAttr(Fld)

                        FldName = ds.DatastoreName & "." & dsSel.SelectionName & "." & Fld.FieldName
                        For i = 2 To 8
                            If Fld.GetFieldAttr(i).ToString.Trim <> "" Then
                                If wFA(rc, FldName, Fld, i) = False Then
                                    GoTo ErrorGoTo
                                End If
                            End If
                        Next
                    Next
                Next
            Next

            For Each ds As clsDatastore In ObjThis.Targets
                For Each dsSel As clsDSSelection In ds.ObjSelections
                    dsSel.LoadMe()

                    For Each Fld As clsField In dsSel.DSSelectionFields
                        'dsSel.LoadFieldAttr(Fld)

                        FldName = ds.DatastoreName & "." & dsSel.SelectionName & "." & Fld.FieldName
                        For i = 2 To 8
                            If Fld.GetFieldAttr(i).ToString.Trim <> "" Then
                                If wFA(rc, FldName, Fld, i) = False Then
                                    GoTo ErrorGoTo
                                End If
                            End If
                        Next
                    Next
                Next
            Next


ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtFA")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Field attributes"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenStruct
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtFA = Not rc.HasFatal

    End Function

    Function prtVAR(ByRef rc As clsRcode) As Boolean

        Try
            If ObjThis.Variables.Count > 0 Then
                For Each var As clsVariable In ObjThis.Variables
                    var.LoadMe()

                    If var.VariableDescription.Trim <> "" Then
                        wComment(rc, var.VariableDescription)
                    End If
                    If wVariable(rc, var) = False Then
                        GoTo ErrorGoTo
                    End If
                Next
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtVar")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Variables"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenVar
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtVAR = Not rc.HasFatal

    End Function

    '    Function prtJoins(ByRef rc As clsRcode) As Boolean

    '        Try
    '            If ObjThis.Gens.Count > 0 Then
    '                For Each Join As clsTask In ObjThis.Gens
    '                    Join.LoadMe()

    '                    If prtJs(rc, Join) = False Then
    '                        GoTo ErrorGoTo
    '                    End If
    'ErrorGoTo:      Next
    '            End If

    '            '/// send returnPath or enumreturncode

    '        Catch ex As Exception
    '            LogError(ex, "modGenerate prtJoins")
    '            rc.HasError = True
    '            rc.ErrorCount += 1
    '            rc.LocalErrorMsg = "Error while writing Joins"
    '            If rc.ReturnCode = "" Then
    '                rc.ReturnCode = ex.Message
    '            End If
    '            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
    '                rc.ErrorLocation = enumErrorLocation.ModGenJoin
    '                rc.ErrorPath = pathSQD
    '                rc.ObjInode = ObjThis
    '            End If
    '            ErrorComment(rc)
    '        End Try

    '        prtJoins = Not rc.HasFatal

    'End Function

    Function prtJs(ByRef rc As clsRcode, ByVal Join As clsTask) As Boolean

        Try
            Dim Map As clsMapping
            Dim i As Integer
            Dim SourceName As String = ""
            Dim First As Boolean = False
            Dim ds As clsDatastore = CType(Join.ObjSources(0), clsDatastore)

            If Join.TaskDescription.Trim <> "" Then
                wComment(rc, Join.TaskDescription)
            End If
            If wJoin(rc, Join, ds.DatastoreName) = False Then
                GoTo ErrorGoTo
            End If
            For i = 0 To Join.ObjMappings.Count - 1
                Map = Join.ObjMappings(i)
                If i = 0 Or First = True Then
                    SourceName = Map.SourceDataStore
                    First = True
                Else
                    First = False
                End If

                If CType(Map, clsMapping).SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                    If CType(Map, clsMapping).TargetType = enumMappingType.MAPPING_TYPE_NONE Then
                        If Map.MappingDesc.Trim <> "" Then
                            wComment(rc, Map.MappingDesc)
                        End If
                        If wMapSrcOnly(rc, Map, First) = False Then
                            GoTo ErrorGoTo
                        End If
                    End If
                    First = False
                Else

                    MsgBox("found the non field: " & Map.MappingSource, MsgBoxStyle.OkOnly, MsgTitle)
                    First = True
                End If
            Next
            If wJoin(rc, Join, ds.DatastoreName, False) = False Then
                GoTo ErrorGoTo
            End If
            If wDSkey(rc, ds) = False Then
                GoTo ErrorGoTo
            End If
            wSemiLine(rc)

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtJs")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Joins"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenJoin
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtJs = Not rc.HasFatal

    End Function

    Function prtLookups(ByRef rc As clsRcode) As Boolean

        Try
            If ObjThis.Lookups.Count > 0 Then
                For Each LU As clsTask In ObjThis.Lookups
                    LU.LoadMe()

                    If prtLUs(rc, LU) = False Then
                        GoTo ErrorGoTo
                    End If
                Next
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtLookups")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Lookups"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenLU
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtLookups = Not rc.HasFatal

    End Function

    Function prtLUs(ByRef rc As clsRcode, ByVal LU As clsTask) As Boolean

        Try
            Dim Map As clsMapping
            Dim i As Integer
            Dim SourceName As String = ""
            Dim First As Boolean = False

            If LU.ObjSources.Count = 1 And LU.ObjTargets.Count = 0 Then
                Dim ds As clsDatastore = CType(LU.ObjSources(0), clsDatastore)
                If LU.TaskDescription.Trim <> "" Then
                    wComment(rc, LU.TaskDescription)
                End If
                If wLookup(rc, LU) = False Then
                    GoTo ErrorGoTo
                End If
                For i = 0 To LU.ObjMappings.Count - 1
                    Map = LU.ObjMappings(i)
                    If i = 0 Or First = True Then
                        SourceName = Map.SourceDataStore
                        First = True
                    Else
                        First = False
                    End If

                    If CType(Map, clsMapping).SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                        If CType(Map, clsMapping).TargetType = enumMappingType.MAPPING_TYPE_NONE Then
                            If Map.MappingDesc.Trim <> "" Then
                                wComment(rc, Map.MappingDesc)
                            End If
                            If wMapSrcOnly(rc, Map, First) = False Then
                                GoTo ErrorGoTo
                            End If
                        End If
                        First = False
                    Else

                        MsgBox("found the non field: " & Map.MappingSource, MsgBoxStyle.OkOnly, MsgTitle)
                        First = True
                    End If
                Next
                If wLookup(rc, LU, SourceName, False) = False Then
                    GoTo ErrorGoTo
                End If
                If wDSkey(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
                wSemiLine(rc)
            End If

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtLookups")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Lookups"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenLU
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtLUs = Not rc.HasFatal

    End Function

    Function prtProcedures(ByRef rc As clsRcode) As Boolean

        Try
            For Each proc As clsTask In ObjThis.Procs
                proc.LoadMe()
                'proc.LoadDatastores()
                'proc.LoadMappings(True)

                If PrtProcs(rc, proc) = False Then
                    GoTo ErrorGoTo1
                End If
ErrorGoTo1: Next

            '            For Each Rproc As clsTask In ObjThis.Mains
            '                Rproc.LoadMe()
            '                'Rproc.LoadDatastores()
            '                'Rproc.LoadMappings(True)

            '                ' IsMapped is True if it is the Main Procedure
            '                If Rproc.TaskType = enumTaskType.TASK_GEN Then
            '                    If PrtProcs(rc, Rproc) = False Then
            '                        GoTo ErrorGoTo2
            '                    End If
            '                End If
            'ErrorGoTo2: Next

            '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtProcedures")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Procedures"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenProc
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtProcedures = Not rc.HasFatal

    End Function

    Function PrtProcs(ByRef rc As clsRcode, ByVal Proc As clsTask) As Boolean

        Dim Map As clsMapping
        Dim i As Integer
        Dim SourceName As String = ""
        'Dim First As Boolean = False
        Dim flag As Boolean = False
        Dim flag2 As Boolean = False

        Try
            'wComment(rc, "---------------------------------------------------------------------")
            If Proc.TaskType = enumTaskType.TASK_IncProc Then
                If prtInclude(rc, Proc) = False Then
                    GoTo ErrorGoTo
                Else
                    GoTo ErrorGoTo
                End If
            Else
                If Proc.TaskDescription IsNot Nothing Then
                    If Proc.TaskDescription.Trim <> "" Then
                        wComment(rc, Proc.TaskDescription)
                    End If
                End If
                If wProc(rc, Proc) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            If DBGMap = True Then
                If wDebugHead(rc, Proc) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            '/// Print Debug OutMsg if debug script
            If OutMsg = True Then
                For i = 0 To Proc.ObjMappings.Count - 1
                    Map = CType(Proc.ObjMappings(i), clsMapping)
                    If Map.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                        If wDebug(rc, Map) = False Then
                            GoTo ErrorGoTo
                        End If
                    End If
                Next
            Else
                For Each mapping As clsMapping In Proc.ObjMappings
                    If mapping.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                        If mapping.OutMsg = enumOutMsg.PrintOutMsg Then
                            If wDebug(rc, mapping, False) = False Then
                                GoTo ErrorGoTo
                            End If
                        End If
                    End If
                Next
            End If

            flag = False
            For i = 0 To Proc.ObjMappings.Count - 1
                Map = CType(Proc.ObjMappings(i), clsMapping)

                If flag = False Then
                    If Map.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                        'fix this
                        SourceName = Map.SourceDataStore
                        flag = True
                    Else
                        For Each src As clsDatastore In Proc.ObjSources
                            If Map.SourceType = enumMappingType.MAPPING_TYPE_FUN Then
                                If CType(Map.MappingSource, clsSQFunction).SQFunctionWithInnerText. _
                                ToString.Contains(src.DatastoreName) = True Then
                                    SourceName = src.DatastoreName
                                    flag = True
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                    If SourceName = "" Then
                        SourceName = CType(Proc.ObjSources(0), clsDatastore).DatastoreName
                    End If
                End If
                'If i = 0 Then
                '    First = True
                'Else
                '    First = False
                'End If
                '___________________________________** FOR DEBUGGING, Not "debug" ** ________________________________
                '/// Process all of the Mappings
                '/// FOR DEBUG: Log all Mapping Types of Mappings
                '/// Log file is "[enginename]MapType.log"
                'If Map.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                '    Dim str1 As String = String.Format("{0,-35}{1}", "Source --> " & _
                '    CType(Map.MappingSource, clsField).FieldName, _
                '    "  Type --> " & Map.SourceType.ToString())
                '    objwriteLOG.WriteLine(str1)

                '    Dim str2 As String = String.Format("{0,-35}{1}", "Target --> " & _
                '    CType(Map.MappingTarget, clsField).FieldName, _
                '    "  Type --> " & Map.TargetType.ToString())
                '    objwriteLOG.WriteLine(str2)
                'Else
                '    Dim str3 As String = String.Format("{0,-35}{1}", "Source --> " & _
                '    CType(Map.MappingSource, INode).Text, _
                '    "  Type --> " & Map.SourceType.ToString())
                '    objwriteLOG.WriteLine(str3)

                '    Dim str4 As String = String.Format("{0,-35}{1}", "Target --> " & _
                '    CType(Map.MappingTarget, INode).Text, _
                '    "  Type --> " & Map.TargetType.ToString())
                '    objwriteLOG.WriteLine(str4)
                'End If
                '____________________________________________________________________________________

                Select Case Map.SourceType

                    Case enumMappingType.MAPPING_TYPE_FIELD
                        If Map.SourceDataStore = SourceName Then
                            If Map.TargetType = enumMappingType.MAPPING_TYPE_FIELD Or Map.TargetType = enumMappingType.MAPPING_TYPE_VAR _
                            Or Map.TargetType = enumMappingType.MAPPING_TYPE_WORKVAR Then
                                If wMap(rc, Map) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            Else
                                '/// skip the mapping, because there is no target
                            End If
                        Else
                            '/// Must be LookUp or Join
                            For Each LU As clsTask In ObjThis.Lookups
                                If CType(LU.ObjSources(0), clsDatastore).DatastoreName = Map.SourceDataStore Then
                                    If wLUfield(rc, LU, Map) = False Then   ', First
                                        GoTo ErrorGoTo
                                    End If
                                    flag2 = True
                                    GoTo mapTgt
                                Else
                                    flag2 = False
                                End If
                            Next
                            For Each Join As clsTask In ObjThis.Gens
                                If CType(Join.ObjSources(0), clsDatastore).DatastoreName = Map.SourceDataStore Then
                                    If wLUfield(rc, Join, Map) = False Then   ', First
                                        GoTo ErrorGoTo
                                    End If
                                    flag2 = True
                                    GoTo mapTgt
                                Else
                                    flag2 = False
                                End If
                            Next
mapTgt:                     '/// Now Map Target
                            If flag2 = False Then
                                If wMap(rc, Map) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            End If
                        End If

                    Case enumMappingType.MAPPING_TYPE_FUN, enumMappingType.MAPPING_TYPE_JOIN, _
                    enumMappingType.MAPPING_TYPE_LOOKUP, enumMappingType.MAPPING_TYPE_PROC, _
                    enumMappingType.MAPPING_TYPE_VAR, enumMappingType.MAPPING_TYPE_WORKVAR, _
                    enumMappingType.MAPPING_TYPE_NONE, enumMappingType.MAPPING_TYPE_CONSTANT


                        If Map.IsMapped = "0" Then
                            If Map.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
                                If wMap(rc, Map, True) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            ElseIf Map.TargetType = enumMappingType.MAPPING_TYPE_VAR Or _
                            Map.TargetType = enumMappingType.MAPPING_TYPE_WORKVAR Then
                                If wMapFunWTgt(rc, Map, True) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            Else
                                If wMapFun(rc, Map) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            End If
                        Else
                            If Proc.TaskType = enumTaskType.TASK_GEN Then
                                If wMapFun(rc, Map) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            Else
                                '/// map function with target
                                '**** for debugging of Function Mappings
                                If DBGMap = True Then
                                    If Map IsNot Nothing Then
                                        Log("map obj exists")
                                    Else
                                        Log("*** map does not exist ***")
                                    End If
                                End If

                                If wMapFunWTgt(rc, Map, False) = False Then   ', First
                                    GoTo ErrorGoTo
                                End If
                            End If
                        End If

                        '/// I think these are all unused, although they
                        '/// are referenced throughout the program. I think
                        '/// they were never used as intended when the 
                        '/// enumeration was created. TK 10/29/07
                        'Case enumMappingType.MAPPING_TYPE_JOIN
                        'Case enumMappingType.MAPPING_TYPE_LOOKUP
                        'Case enumMappingType.MAPPING_TYPE_PROC
                        'Case enumMappingType.MAPPING_TYPE_VAR
                        'Case enumMappingType.MAPPING_TYPE_WORKVAR
                        'Case enumMappingType.MAPPING_TYPE_NONE
                        'Case enumMappingType.MAPPING_TYPE_CONSTANT
                        'Case Else

                End Select
            Next
            If wProc(rc, Proc, False, SourceName) = False Then
                GoTo ErrorGoTo
            End If

ErrorGoTo:
        Catch ex As Exception
            LogError(ex, "modGenerate prtProcs")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Procedures"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenProc
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        PrtProcs = Not rc.HasFatal

    End Function

    Function prtMain(ByRef rc As clsRcode) As Boolean

        Try
            Dim srcName As String = ""
            'Dim main As String = ObjThis.Main
            Dim NoOfMains As Integer = 0 'ObjThis.Mains.Count
            Dim Counter As Integer = 0
            Dim Last As Boolean = False
            Dim first As Boolean = False

            For Each main As clsTask In ObjThis.Mains
                If main.TaskType = enumTaskType.TASK_MAIN Then
                    NoOfMains += 1
                End If
            Next
            '            If ObjThis.Sources.Count >= 1 Then
            '                Counter = ObjThis.Sources.Count
            'repeat:         If CType(ObjThis.Sources(Counter), clsDatastore).IsLookUp = False Then
            '                    srcName = CType(ObjThis.Sources(Counter), clsDatastore).DatastoreName
            '                Else
            '                    Counter = Counter - 1
            '                    GoTo repeat
            '                End If
            '            End If
            For Each main As clsTask In ObjThis.Mains
                main.LoadMe()
                'main.LoadDatastores()
                'main.LoadMappings(True)
                If main.TaskType <> enumTaskType.TASK_MAIN Then
                    GoTo GoToNext
                End If

                Counter += 1
                If Counter = NoOfMains Then
                    Last = True
                End If
                If Counter = 1 Then
                    first = True
                Else
                    first = False
                End If

                srcName = ""
                '/// modified 12/12 to correct 'no source' problem
                For Each src As clsDatastore In ObjThis.Sources
                    If CType(CType(main.ObjMappings(0), clsMapping).MappingSource, clsSQFunction). _
                    SQFunctionWithInnerText.Contains(src.DatastoreName) = True Then
                        Exit For
                    End If
                Next


                '/// this finds the source datastore for each Main Proc
                '/// Added 12/12/07 by TK to correct for loop above
                If srcName = "" Then
                    If main.ObjSources.Count > 0 Then
                        srcName = CType(main.ObjSources(0), INode).Text
                    Else
                        rc.HasError = True
                        rc.ObjInode = main
                        rc.ErrorLocation = enumErrorLocation.ModGenMain
                        rc.ErrorPath = pathSQD
                        rc.ReturnCode = "Main procedure has no Source Datastore specified" & Chr(13) & "Please view the Log"
                        Log("Main procedure has no Source Datastore specified")
                        Log("To correct: right-click Main Procedure and 'add/remove datastores'")
                        Log("Check the datastore that pertains to the Main Procedure")
                        Exit Try
                    End If
                End If

                If main.TaskDescription.Trim <> "" Then
                    wComment(rc, main.TaskDescription)
                End If
                If wMain(rc, main, srcName, True, True) = False Then
                    GoTo ErrorGoTo
                End If
                If Counter > 0 And Counter < NoOfMains Then
                    If wUnion(rc) = False Then
                        GoTo ErrorGoTo
                    End If
                End If
GoToNext:
            Next

ErrorGoTo:  '/// send returnPath or enumreturncode

        Catch ex As Exception
            LogError(ex, "modGenerate prtMain")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Main Procedure"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenMain
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtMain = Not rc.HasFatal

    End Function

    Function prtInclude(ByRef rc As clsRcode, ByVal Obj As INode) As Boolean

        Try
            Dim EngObj As clsEngine

            Select Case Obj.Type
                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    Dim DSobj As clsDatastore = CType(Obj, clsDatastore)
                    EngObj = DSobj.Engine
                    If wInclude(rc, GetIncPath(DSobj.DsPhysicalSource, EngObj)) = False Then
                        GoTo errorgoto
                    End If

                Case NODE_PROC
                    Dim ProcObj As clsTask = CType(Obj, clsTask)
                    EngObj = ProcObj.Engine
                    If wInclude(rc, GetIncPath(ProcObj.TaskDescription, EngObj)) = False Then
                        GoTo errorgoto
                    End If
            End Select
errorgoto:

        Catch ex As Exception
            LogError(ex, "modGenerate prtInclude")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Include"
            If rc.ReturnCode = "" Then
                rc.ReturnCode = ex.Message
            End If
            If rc.ErrorLocation = enumErrorLocation.NoErrors Then
                rc.ErrorLocation = enumErrorLocation.ModGenProc
                rc.ErrorPath = pathSQD
                rc.ObjInode = ObjThis
            End If
            ErrorComment(rc)
        End Try

        prtInclude = Not rc.HasFatal

    End Function

#End Region

#Region " Write to file Functions "

    Function wHeaderComments(ByRef rc As clsRcode, ByVal LineNo As Integer, ByVal instr As String, Optional ByVal ArgStr As String = "") As Boolean

        Dim EndStr As String = "  --"
        Dim FormatTopBot As String = String.Format("{0,42}", instr)
        Dim FormatNorm As String = String.Format("{0,19}{1,-20}{2,4}", instr, ArgStr, EndStr)
        Dim FormatDateTime As String = String.Format("{0,19}{1,-20:MM/dd/yyyy HH:mm:ss}{2,4}", instr, Now(), EndStr)

        Try
            Select Case LineNo
                Case 0, 1, 6, 7
                    objWriteSQD.WriteLine(FormatTopBot)
                    objWriteINL.WriteLine(FormatTopBot)
                    objWriteTMP.WriteLine(FormatTopBot)
                    AddToLineNo(rc)
                Case 3, 4, 5
                    objWriteSQD.WriteLine(FormatNorm)
                    objWriteINL.WriteLine(FormatNorm)
                    objWriteTMP.WriteLine(FormatNorm)
                    AddToLineNo(rc)
                Case 2
                    objWriteSQD.WriteLine(FormatDateTime)
                    objWriteINL.WriteLine(FormatDateTime)
                    objWriteTMP.WriteLine(FormatDateTime)
                    AddToLineNo(rc)
            End Select

        Catch ex As Exception
            LogError(ex, "modGenerate wHeader")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Header Comments"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenHead
            rc.ErrorPath = pathSQD
            ErrorComment(rc)
        End Try

        wHeaderComments = Not rc.HasFatal

    End Function

    Function wHeader(ByRef rc As clsRcode, ByVal LineNo As Integer, ByVal instr As String) As Boolean

        Try
            Dim ConnType As String = ""

            If LineNo < 5 Then
                Dim VarStr As String = ""
                Select Case LineNo
                    Case 1
                        VarStr = ObjThis.EngineName
                    Case 2
                        If ObjThis.ReportEvery.ToString.Trim <> "" Then
                            VarStr = ObjThis.ReportEvery.ToString
                        End If
                    Case 3
                        If ObjThis.CommitEvery.ToString.Trim <> "" Then
                            If ObjThis.ForceCommit = False Then
                                VarStr = ObjThis.CommitEvery.ToString
                            Else
                                VarStr = ObjThis.CommitEvery.ToString & " FORCE"
                            End If
                        End If
                    Case 4
                        If ObjThis.DateFormat.Trim <> "" Then
                            VarStr = ObjThis.DateFormat.Trim
                        End If
                End Select
                objWriteSQD.WriteLine(QuoteRes(instr) & QuoteRes(VarStr) & semi)
                objWriteINL.WriteLine(QuoteRes(instr) & QuoteRes(VarStr) & semi)
                objWriteTMP.WriteLine(QuoteRes(instr) & QuoteRes(VarStr) & semi)
                AddToLineNo(rc)
            Else
                If ObjThis.Connection IsNot Nothing Then
                    wBlankLine(rc)
                    If ObjThis.Connection.ConnectionType = "ODBCwithSubstitutionVariable" Then
                        ConnType = "ODBC"
                    ElseIf ObjThis.Connection.ConnectionType = "DB2" Then
                        If ObjThis.ObjSystem.OSType = "z/OS" Then
                            ConnType = ""
                        Else
                            ConnType = "NATIVEDB2"
                        End If
                    Else
                        ConnType = ObjThis.Connection.ConnectionType
                    End If
                    If PrintUID = True Then
                        objWriteSQD.WriteLine(instr & ConnType & " " & _
                        ObjThis.Connection.Database & " " & ObjThis.Connection.UserId & " " & _
                        ObjThis.Connection.Password & semi)
                        objWriteINL.WriteLine(instr & ConnType & " " & _
                        ObjThis.Connection.Database & " " & ObjThis.Connection.UserId & " " & _
                        ObjThis.Connection.Password & semi)
                        objWriteTMP.WriteLine(instr & ConnType & " " & _
                        ObjThis.Connection.Database & " " & ObjThis.Connection.UserId & " " & _
                        ObjThis.Connection.Password & semi)
                    Else
                        objWriteSQD.WriteLine(instr & ConnType & " " & _
                        QuoteRes(ObjThis.Connection.Database) & " dummy dummy " & semi)
                        objWriteINL.WriteLine(instr & ConnType & " " & _
                        QuoteRes(ObjThis.Connection.Database) & " dummy dummy " & semi)
                        objWriteTMP.WriteLine(instr & ConnType & " " & _
                        QuoteRes(ObjThis.Connection.Database) & " dummy dummy " & semi)
                    End If

                    AddToLineNo(rc)
                End If
                wBlankLine(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wHeader")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Engine Info"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenHead
            rc.ErrorPath = pathSQD
            ErrorComment(rc)
        End Try

        wHeader = Not rc.HasFatal

    End Function

    Function wDS(ByRef rc As clsRcode, ByRef ds As clsDatastore) As Boolean

        Try
            Dim commaORspace As String = " "
            Dim selName As String = ""
            Dim MQstr As String = ds.Engine.ObjSystem.QueueManager
            Dim TCPport As String = ds.DsPort.Trim
            If TCPport = "" Then
                If ds.DsPort.Trim <> "" Then
                    TCPport = ds.DsPort.Trim
                ElseIf ObjSys.Port.Trim <> "" Then
                    TCPport = ObjSys.Port.Trim
                End If
            End If
            Dim DSname As String = ds.DsPhysicalSource
            Dim i As Integer

            Dim AccessHost As String = "" '= ObjSys.Host.Trim
            If ds.DsHostName.Trim <> "" Then
                AccessHost = ds.DsHostName.Trim
            ElseIf ObjSys.Host.Trim <> "" Then
                AccessHost = ObjSys.Host.Trim
            End If

            If AccessHost <> "" And TCPport <> "" And ds.DsAccessMethod <> DS_ACCESSMETHOD_MQSERIES Then
                AccessHost = AccessHost & ":" & TCPport
            End If

            If ds.DsDirection = DS_DIRECTION_SOURCE Then
                If SynNew = False Then
                    '*******OLD SYNTAX *******
                    Select Case ds.DsAccessMethod
                        Case DS_ACCESSMETHOD_FILE
                            DSname = ds.DsPhysicalSource
                        Case DS_ACCESSMETHOD_IP
                            DSname = ds.DsPhysicalSource & ":" & TCPport.Trim & "@TCP"
                        Case DS_ACCESSMETHOD_MQSERIES
                            If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                            Strings.Left(DSname, 3) = "dD:" Then
                                DSname = ds.DsPhysicalSource
                            Else
                                If MQstr.Trim = "" Then
                                    DSname = ds.DsPhysicalSource & "@MQS"
                                Else
                                    DSname = ds.DsPhysicalSource & "#" & MQstr.Trim & "@MQS"
                                End If
                            End If
                        Case DS_ACCESSMETHOD_CDCSTORE
                            DSname = "cdc:///" & ds.DsPhysicalSource & "/" '& ds.DatastoreName '& ":" & TCPport.Trim
                        Case DS_ACCESSMETHOD_VSAM
                            DSname = ds.DsPhysicalSource
                        Case Else
                            DSname = ds.DsPhysicalSource
                    End Select

                    If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                        If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                        Strings.Left(DSname, 3) = "dD:" Or DSname.Contains("@MQS") = True Then
                            DSname = DSname
                        Else
                            DSname = "'" & DSname & "'"
                        End If
                    End If
                Else
                    '*********New Syntax ************
                    Select Case ds.DsAccessMethod
                        Case DS_ACCESSMETHOD_FILE
                            DSname = "file:///" & Quote(ds.DsPhysicalSource) '& "/" & ds.DatastoreName
                        Case DS_ACCESSMETHOD_IP
                            DSname = "tcp://" & AccessHost & "/" & ds.DsPhysicalSource & "/" & ds.Engine.EngineName '":" & TCPport.Trim '"/"& "/" & ds.DatastoreName & 
                        Case DS_ACCESSMETHOD_MQSERIES
                            If Strings.Left(DSname.ToUpper, 3) = "DD:" Then  'Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or Strings.Left(DSname, 3) = "dD:"
                                DSname = ds.DsPhysicalSource
                            Else
                                If MQstr.Trim = "" Then
                                    DSname = "mqs:///" & ds.DsPhysicalSource '& "@MQS" " & AccessHost & "
                                Else
                                    DSname = "mqs://" & MQstr.Trim & "/" & ds.DsPhysicalSource '& "#" & MQstr.Trim & "@MQS"& AccessHost & "/"
                                End If
                            End If
                        Case DS_ACCESSMETHOD_CDCSTORE
                            DSname = "cdc://" & AccessHost & "/" & ds.DsPhysicalSource & "/" & ds.Engine.EngineName 'ds.DatastoreName '& "/" & ds.DatastoreName '& ":" & TCPport.Trim
                        Case DS_ACCESSMETHOD_VSAM
                            DSname = "vsam://" & AccessHost & "/" & ds.DsPhysicalSource
                        Case Else
                            '*************** CDCStore ??? ***********************
                            DSname = ds.DsPhysicalSource
                    End Select

                    'If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                    '    If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                    '    Strings.Left(DSname, 3) = "dD:" Or DSname.Contains("@MQS") = True Then
                    '        DSname = DSname
                    '    Else
                    '        DSname = "'" & DSname & "'"
                    '    End If
                    'End If

                End If
            Else
                If SynNew = False Then
                    '*******OLD SYNTAX *******
                    Select Case ds.DsAccessMethod
                        Case DS_ACCESSMETHOD_FILE
                            DSname = ds.DsPhysicalSource
                        Case DS_ACCESSMETHOD_IP
                            DSname = ds.DsPhysicalSource & ":" & TCPport.Trim & "@TCP"
                        Case DS_ACCESSMETHOD_MQSERIES
                            If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                            Strings.Left(DSname, 3) = "dD:" Then
                                DSname = ds.DsPhysicalSource
                            Else
                                If MQstr.Trim = "" Then
                                    DSname = ds.DsPhysicalSource & "@MQS"
                                Else
                                    DSname = ds.DsPhysicalSource & "#" & MQstr.Trim & "@MQS"
                                End If
                            End If
                        Case DS_ACCESSMETHOD_CDCSTORE
                            DSname = "cdc:///" & ds.DsPhysicalSource & "/" '& ds.DatastoreName '& ":" & TCPport.Trim
                        Case DS_ACCESSMETHOD_VSAM
                            DSname = ds.DsPhysicalSource
                        Case Else
                            DSname = ds.DsPhysicalSource
                    End Select

                    If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                        If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                        Strings.Left(DSname, 3) = "dD:" Or DSname.Contains("@MQS") = True Then
                            DSname = DSname
                        Else
                            DSname = "'" & DSname & "'"
                        End If
                    End If
                Else
                    '*********New Syntax ************
                    Select Case ds.DsAccessMethod
                        Case DS_ACCESSMETHOD_FILE
                            DSname = "file:///" & Quote(ds.DsPhysicalSource) '& "/" & ds.DatastoreName
                        Case DS_ACCESSMETHOD_IP
                            DSname = "tcp://" & AccessHost & "/" & ds.DsPhysicalSource & "/" & ds.Engine.EngineName '":" & TCPport.Trim '"/"& "/" & ds.DatastoreName & 
                        Case DS_ACCESSMETHOD_MQSERIES
                            If Strings.Left(DSname.ToUpper, 3) = "DD:" Then  'Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or Strings.Left(DSname, 3) = "dD:"
                                DSname = ds.DsPhysicalSource
                            Else
                                If MQstr.Trim = "" Then
                                    DSname = "mqs:///" & ds.DsPhysicalSource '& "@MQS"" & AccessHost & "
                                Else
                                    DSname = "mqs://" & MQstr.Trim & "/" & ds.DsPhysicalSource '& "#" & MQstr.Trim & "@MQS"& AccessHost & "/"
                                End If
                            End If
                        Case DS_ACCESSMETHOD_CDCSTORE
                            DSname = "cdc://" & AccessHost & "/" & ds.DsPhysicalSource & "/" & ds.Engine.EngineName 'ds.DatastoreName '& "/" & ds.DatastoreName '& ":" & TCPport.Trim
                        Case DS_ACCESSMETHOD_VSAM
                            DSname = "vsam://" & AccessHost & "/" & ds.DsPhysicalSource
                        Case Else
                            '*************** CDCStore ??? ***********************
                            DSname = ds.DsPhysicalSource
                    End Select

                    'If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                    '    If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                    '    Strings.Left(DSname, 3) = "dD:" Or DSname.Contains("@MQS") = True Then
                    '        DSname = DSname
                    '    Else
                    '        DSname = "'" & DSname & "'"
                    '    End If
                    'End If

                End If
            End If
            

            '/// define formatted strings
            Dim FORds1 As String
            If SynNew = True Then
                FORds1 = String.Format("{0}", "DATASTORE " & QuoteRes(DSname))
            Else
                FORds1 = String.Format("{0}", "DATASTORE " & Quote(QuoteRes(DSname)))
            End If

            Dim FORdsForOf As String = String.Format("{0,12}{1}{2}", " ", "OF ", GetDStype(ds.DatastoreType))
            Dim FORds2 As String = String.Format("{0,12}{1}{2}", " ", "AS ", QuoteRes(ds.DatastoreName))
            Dim FORds3 As String = String.Format("{0,12}{1}", " ", "DESCRIBED BY")

            If ds.DatastoreDescription.Trim <> "" Then
                wComment(rc, ds.DatastoreDescription)
                wComment(rc, "")
            End If

            objWriteSQD.WriteLine(FORds1)
            objWriteINL.WriteLine(FORds1)
            objWriteTMP.WriteLine(FORds1)
            'AddToLineNo(rc)
            objWriteSQD.WriteLine(FORdsForOf)
            objWriteINL.WriteLine(FORdsForOf)
            objWriteTMP.WriteLine(FORdsForOf)
            AddToLineNo(rc)

            If ds.DatastoreType = enumDatastore.DS_DELIMITED Then
                'wBlankLine(rc)
                If wDSdelimited(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If
            If ds.DatastoreType = enumDatastore.DS_BINARY Or ds.DatastoreType = enumDatastore.DS_DELIMITED Then
                If wASCII(rc, ds) = False Then
                    GoTo ErrorGoTo
                End If
            End If

            objWriteSQD.WriteLine(FORds2)
            objWriteINL.WriteLine(FORds2)
            objWriteTMP.WriteLine(FORds2)
            AddToLineNo(rc)
            objWriteSQD.WriteLine(FORds3)
            objWriteINL.WriteLine(FORds3)
            objWriteTMP.WriteLine(FORds3)
            AddToLineNo(rc)
            For i = 0 To ds.ObjSelections.Count - 1
                If i = 0 Then
                    commaORspace = " "
                Else
                    commaORspace = ","
                End If
                selName = CType(ds.ObjSelections(i), clsDSSelection).ObjStructure.StructureName
                Dim FORds4 As String = String.Format("{0,21}{1}{2}", " ", commaORspace, QuoteRes(selName))
                objWriteSQD.WriteLine(FORds4)
                objWriteINL.WriteLine(FORds4)
                objWriteTMP.WriteLine(FORds4)
                AddToLineNo(rc)
                commaORspace = ","
            Next
            '/// Now add UOW if Used
            If ds.DsUOW.Trim <> "" And ds.DatastoreType = enumDatastore.DS_DB2CDC Then
                Dim FORds5 As String = String.Format("{0,21}{1}{2}", " ", commaORspace, "UOW")
                objWriteSQD.WriteLine(FORds5)
                objWriteINL.WriteLine(FORds5)
                objWriteTMP.WriteLine(FORds5)
                AddToLineNo(rc)
            End If

ErrorGoTo:
        Catch ex As Exception
            LogError(ex, "modGenerate wDS")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDS = Not rc.HasFatal

    End Function

    Function wDSkey(ByRef rc As clsRcode, ByRef ds As clsDatastore, Optional ByVal isLUDS As Boolean = False) As Boolean

        Try
            Dim i, j As Integer
            Dim dssel As clsDSSelection
            Dim fld As clsField
            Dim prefix As String = " "
            Dim first As Boolean = True

            If ds.DsDirection <> DS_DIRECTION_SOURCE Or ds.IsLookUp = True Then
                For i = 0 To ds.ObjSelections.Count - 1
                    dssel = CType(ds.ObjSelections(i), clsDSSelection)
                    For j = 0 To dssel.DSSelectionFields.Count - 1
                        fld = CType(dssel.DSSelectionFields(j), clsField)
                        If fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                            Dim ForKey As String
                            Dim FldPar As String
                            If ds.IsLookUp = True And isLUDS = False Then
                                FldPar = ds.DatastoreName
                            Else
                                FldPar = fld.ParentStructureName
                            End If
                            If first = True Then
                                '/// Print "key is" first
                                Dim ForKeyIs As String = String.Format("{0,12}{1}", " ", "KEY IS")
                                objWriteSQD.Write(ForKeyIs)
                                objWriteINL.Write(ForKeyIs)
                                objWriteTMP.Write(ForKeyIs)
                                AddToLineNo(rc)
                                ForKey = String.Format("{0,1}{1}{2}", prefix, FldPar, "." & fld.FieldName)
                                first = False
                            Else
                                prefix = ","
                                ForKey = String.Format("{0,19}{1}{2}", prefix, FldPar, "." & fld.FieldName)
                            End If
                            objWriteSQD.WriteLine(ForKey)
                            objWriteINL.WriteLine(ForKey)
                            objWriteTMP.WriteLine(ForKey)
                            AddToLineNo(rc)
                        End If
                    Next
                Next
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wDSkey")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore Key fields"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSkey = Not rc.HasFatal

    End Function

    Function wDummy(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            Dim commaORspace As String = " "
            Dim selName As String = ""
            Dim MQstr As String = ds.Engine.ObjSystem.QueueManager
            Dim TCPport As String = ds.DsPort.Trim
            Dim EXname As String = ds.ExceptionDatastore
            Dim AccessHost As String = ds.DsHostName.Trim
            If TCPport = "" Then
                TCPport = ObjSys.Port.Trim
            End If
            'Dim DSEXname As String = ds.ExceptionDatastore
            'Dim i As Integer

            If AccessHost = "" Then
                AccessHost = ObjSys.Host.Trim
            End If
            If AccessHost = "" Then
                AccessHost = "localhost"
            End If

            If SynNew = False Then
                '*******OLD SYNTAX *******
                Select Case ds.DsAccessMethod
                    Case DS_ACCESSMETHOD_FILE
                        If ds.ExceptionDatastore <> "" Then
                            EXname = ds.ExceptionDatastore
                        End If
                    Case DS_ACCESSMETHOD_IP
                        EXname = ds.ExceptionDatastore & ":" & TCPport.Trim & "@TCP"
                    Case DS_ACCESSMETHOD_MQSERIES
                        If Strings.Left(EXname, 3) = "DD:" Or Strings.Left(EXname, 3) = "dd:" Or Strings.Left(EXname, 3) = "Dd:" Or _
                        Strings.Left(EXname, 3) = "dD:" Then
                            EXname = ds.ExceptionDatastore
                        Else
                            If MQstr.Trim = "" Then
                                EXname = ds.ExceptionDatastore & "@MQS"
                            Else
                                EXname = ds.ExceptionDatastore & "#" & MQstr.Trim & "@MQS"
                            End If
                        End If
                    Case DS_ACCESSMETHOD_VSAM
                        EXname = ds.ExceptionDatastore
                    Case Else
                        EXname = ds.ExceptionDatastore
                End Select
                'If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                '    If Strings.Left(EXname, 3) = "DD:" Or Strings.Left(EXname, 3) = "dd:" Or Strings.Left(EXname, 3) = "Dd:" Or _
                '    Strings.Left(EXname, 3) = "dD:" Or EXname.Contains("@MQS") = True Then
                '        EXname = EXname
                '    Else
                '        EXname = "'" & EXname & "'"
                '    End If
                'End If
            Else
                '*********New Syntax ************
                Select Case ds.DsAccessMethod
                    Case DS_ACCESSMETHOD_FILE
                        EXname = "file:///" & Quote(ds.ExceptionDatastore) '& "/" & ds.DatastoreName
                    Case DS_ACCESSMETHOD_IP
                        EXname = "tcp://" & AccessHost & ":" & TCPport.Trim '"/" & ds.DsPhysicalSource & "/" & ds.DatastoreName & 
                    Case DS_ACCESSMETHOD_MQSERIES
                        If Strings.Left(EXname, 3) = "DD:" Or Strings.Left(EXname, 3) = "dd:" Or Strings.Left(EXname, 3) = "Dd:" Or _
                        Strings.Left(EXname, 3) = "dD:" Then
                            EXname = ds.ExceptionDatastore
                        Else
                            If MQstr.Trim = "" Then
                                EXname = "mqs://" & AccessHost & "/" & ds.ExceptionDatastore '& "@MQS"
                            Else
                                EXname = "mqs://" & AccessHost & "/" & MQstr.Trim & "/" & ds.ExceptionDatastore '& "#" & MQstr.Trim & "@MQS"
                            End If
                        End If
                    Case DS_ACCESSMETHOD_CDCSTORE
                        EXname = "cdc://" & AccessHost & "/" & ds.ExceptionDatastore & "/" & ds.DatastoreName '& "/" & ds.DatastoreName '& ":" & TCPport.Trim
                    Case DS_ACCESSMETHOD_VSAM
                        EXname = "vsam://" & AccessHost & "/" & ds.ExceptionDatastore
                    Case Else
                        EXname = ds.ExceptionDatastore
                End Select

                'If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
                '    If Strings.Left(DSname, 3) = "DD:" Or Strings.Left(DSname, 3) = "dd:" Or Strings.Left(DSname, 3) = "Dd:" Or _
                '    Strings.Left(DSname, 3) = "dD:" Or DSname.Contains("@MQS") = True Then
                '        EXname = EXname
                '    Else
                '        EXname = "'" & EXname & "'"
                '    End If
                'End If
            End If


            Dim ForDummy As String
            If SynNew = True Then
                ForDummy = String.Format("{0}{1}{2}", "DATASTORE " & QuoteRes(EXname), " OF BINARY AS " & _
            QuoteRes("EX_" & ds.DatastoreName), " DESCRIBED BY DUMMY" & semi)
            Else
                ForDummy = String.Format("{0}{1}{2}", "DATASTORE " & Quote(QuoteRes(EXname)), " OF BINARY AS " & _
            QuoteRes("EX_" & ds.DatastoreName), " DESCRIBED BY DUMMY" & semi)
            End If

            wComment(rc, "-------------- *** Exception Datastore *** --------------")
            objWriteSQD.WriteLine(ForDummy)
            objWriteINL.WriteLine(ForDummy)
            objWriteTMP.WriteLine(ForDummy)
            AddToLineNo(rc)
            wBlankLine(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wDummy")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Exception Datastore"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDummy = Not rc.HasFatal

    End Function

    'Function wDummyReSTRT(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

    '    Try
    '        Dim commaORspace As String = " "
    '        Dim selName As String = ""
    '        Dim MQstr As String = ds.Engine.ObjSystem.QueueManager
    '        Dim TCPport As String = ds.DsPort
    '        Dim RSname As String = ds.Restart

    '        Select Case ds.DsAccessMethod
    '            Case DS_ACCESSMETHOD_FILE
    '                If ds.Restart <> "" Then
    '                    RSname = ds.Restart
    '                End If
    '            Case DS_ACCESSMETHOD_IP
    '                RSname = ds.Restart & ":" & TCPport.Trim & "@TCP"
    '            Case DS_ACCESSMETHOD_MQSERIES
    '                If Strings.Left(RSname, 3) = "DD:" Or Strings.Left(RSname, 3) = "dd:" Or Strings.Left(RSname, 3) = "Dd:" Or _
    '                Strings.Left(RSname, 3) = "dD:" Then
    '                    RSname = ds.Restart
    '                Else
    '                    If MQstr.Trim = "" Then
    '                        RSname = ds.Restart & "@MQS"
    '                    Else
    '                        RSname = ds.Restart & "#" & MQstr.Trim & "@MQS"
    '                    End If
    '                End If

    '            Case DS_ACCESSMETHOD_VSAM
    '                RSname = ds.Restart
    '            Case Else
    '                RSname = ds.Restart
    '        End Select
    '        If ds.DsAccessMethod = DS_ACCESSMETHOD_MQSERIES Or ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
    '            If Strings.Left(RSname, 3) = "DD:" Or Strings.Left(RSname, 3) = "dd:" Or Strings.Left(RSname, 3) = "Dd:" Or _
    '            Strings.Left(RSname, 3) = "dD:" Or RSname.Contains("@MQS") = True Then
    '                RSname = RSname
    '            Else
    '                RSname = "'" & RSname & "'"
    '            End If
    '        End If

    '        Dim ForRstrtDummy As String = String.Format("{0}{1}{2}", "DATASTORE '" & QuoteRes(RSname), "' OF BINARY AS CDCRSRT", _
    '        " DESCRIBED BY DUMMY" & semi)

    '        wComment(rc, "-------------- *** Restart Datastore *** --------------")
    '        objWriteSQD.WriteLine(ForRstrtDummy)
    '        objWriteINL.WriteLine(ForRstrtDummy)
    '        objWriteTMP.WriteLine(ForRstrtDummy)
    '        AddToLineNo(rc)
    '        wBlankLine(rc)

    '    Catch ex As Exception
    '        LogError(ex, "modGenerate wDummyReSTRT")
    '        rc.HasError = True
    '        rc.ErrorCount += 1
    '        rc.LocalErrorMsg = "Error while writing Restart Datastore"
    '        rc.ReturnCode = ex.Message
    '        rc.ErrorLocation = enumErrorLocation.ModGenDS
    '        rc.ErrorPath = pathSQD
    '        rc.ObjInode = ds
    '        ErrorComment(rc)
    '    End Try

    '    wDummyReSTRT = Not rc.HasFatal

    'End Function

    Function wDSdelimited(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            Dim bYes As Boolean = True

            '/// write character delimiter
            If ds.TextQualifier <> "" Then
                bYes = True
                If ds.TextQualifier.Trim = DS_DELIMITER_CR Or ds.TextQualifier.Trim = DS_DELIMITER_LF Then
                    bYes = False
                End If
            Else
                bYes = False
            End If
            If bYes = True Then
                Dim ForDStextQ As String = String.Format("{0,12}{1}", " ", "CHRDEL(" & Quote(ds.TextQualifier) & ")")
                objWriteSQD.WriteLine(ForDStextQ)
                objWriteINL.WriteLine(ForDStextQ)
                objWriteTMP.WriteLine(ForDStextQ)
                AddToLineNo(rc)
            End If

            '/// write record delimiter
            bYes = True
            If ds.RowDelimiter <> "" Then
                bYes = True
                If ds.RowDelimiter.Trim = DS_DELIMITER_CR Or ds.RowDelimiter.Trim = DS_DELIMITER_LF Then
                    bYes = False
                End If
            Else
                bYes = False
            End If
            If bYes = True Then
                Dim ForDSrowDel As String = String.Format("{0,12}{1}", " ", "RECDEL(" & Quote(ds.RowDelimiter.Trim) & ")")
                objWriteSQD.WriteLine(ForDSrowDel)
                objWriteINL.WriteLine(ForDSrowDel)
                objWriteTMP.WriteLine(ForDSrowDel)
                AddToLineNo(rc)
            End If

            '/// write column delimiter
            bYes = True
            If ds.ColumnDelimiter <> "" Then
                bYes = True
                If ds.ColumnDelimiter.Trim = DS_DELIMITER_CR Or ds.ColumnDelimiter.Trim = DS_DELIMITER_LF Then
                    bYes = False
                End If
            Else
                bYes = False
            End If
            If bYes = True Then
                Dim ForDScolDel As String = String.Format("{0,12}{1}", " ", "COLDEL(" & Quote(ds.ColumnDelimiter) & ")")
                objWriteSQD.WriteLine(ForDScolDel)
                objWriteINL.WriteLine(ForDScolDel)
                objWriteTMP.WriteLine(ForDScolDel)
                AddToLineNo(rc)
            End If


        Catch ex As Exception
            LogError(ex, "modGenerate wDSdelimited")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Delimited Datastore"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSdelimited = Not rc.HasFatal

    End Function

    Function wASCII(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            Dim strCharCode As String

            If ds.DsDirection = DS_DIRECTION_SOURCE And ds.Engine.ObjSystem.OSType = "z/OS" Then
                If ds.DsCharacterCode = DS_CHARACTERCODE_ASCII Then
                    strCharCode = "ASCII"
                    Dim ForCC As String = String.Format("{0,12}{1}", " ", strCharCode)
                    objWriteSQD.WriteLine(ForCC)
                    objWriteINL.WriteLine(ForCC)
                    objWriteTMP.WriteLine(ForCC)
                    AddToLineNo(rc)
                End If
            ElseIf ds.DsDirection = DS_DIRECTION_SOURCE And ds.Engine.ObjSystem.OSType = "UNIX" Then
                If ds.DsCharacterCode = DS_CHARACTERCODE_EBCDIC Then
                    strCharCode = "EBCDIC"
                    Dim ForCC As String = String.Format("{0,12}{1}", " ", strCharCode)
                    objWriteSQD.WriteLine(ForCC)
                    objWriteINL.WriteLine(ForCC)
                    objWriteTMP.WriteLine(ForCC)
                    AddToLineNo(rc)
                End If
            End If


        Catch ex As Exception
            LogError(ex, "modGenerate wASCII")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Delimited Datastore"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wASCII = Not rc.HasFatal

    End Function

    Function wDSuow(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            Dim MQstr As String = ds.Engine.ObjSystem.QueueManager
            Dim TCPport As String = ds.DsPort
            Dim DSname As String = ""

            Select Case ds.DsAccessMethod
                Case DS_ACCESSMETHOD_FILE
                    Exit Try
                Case DS_ACCESSMETHOD_IP
                    DSname = ds.DsUOW & ":" & TCPport.Trim & "@TCP"
                Case DS_ACCESSMETHOD_MQSERIES
                    If MQstr.Trim = "" Then
                        DSname = ds.DsUOW & "@MQS"
                    Else
                        DSname = ds.DsUOW & "#" & MQstr.Trim & "@MQS"
                    End If
                Case DS_ACCESSMETHOD_VSAM
                    DSname = ds.DsUOW
            End Select

            Dim ForDummy As String = String.Format("{0}{1}", "DESCRIPTION DB2CDCURQ '" & QuoteRes(DSname), "' AS UOW" & semi)

            wComment(rc, "-------------- *** Unit Of Work Description *** ---------------")
            objWriteSQD.WriteLine(ForDummy)
            objWriteINL.WriteLine(ForDummy)
            objWriteTMP.WriteLine(ForDummy)
            AddToLineNo(rc)
            wBlankLine(rc)


        Catch ex As Exception
            LogError(ex, "modGenerate wDSuow")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore UOW"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSuow = Not rc.HasFatal

    End Function

    Function wDSattrib(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            'If ds.Restart.Trim <> "" And ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
            '    Dim ForRestrt As String = String.Format("{0,12}{1}", " ", "RESTART CDCRSRT")
            '    objWriteSQD.WriteLine(ForRestrt)
            '    objWriteINL.WriteLine(ForRestrt)
            '    objWriteTMP.WriteLine(ForRestrt)
            '    AddToLineNo(rc)
            'End If

            'If ds.Restart.Trim <> "" And ds.Poll.Trim <> "" And ds.DsAccessMethod = DS_ACCESSMETHOD_VSAM Then
            '    Dim ForPoll As String = String.Format("{0,12}{1}", " ", "POLL " & ds.Poll)
            '    objWriteSQD.WriteLine(ForPoll)
            '    objWriteINL.WriteLine(ForPoll)
            '    objWriteTMP.WriteLine(ForPoll)
            '    AddToLineNo(rc)
            'End If
            '/// Add UOW description to datastore if necessary
            'If ds.DatastoreType = enumDatastore.DS_DB2CDC Then
            '    If ds.DsUOW.Trim <> "" Then
            '        Dim FORds5 As String = String.Format("{0,12}{1}", " ", "CDCUOW UOW_" & ds.DatastoreName)
            '        objWriteSQD.WriteLine(FORds5)
            '        objWriteINL.WriteLine(FORds5)
            '        objWriteTMP.WriteLine(FORds5)
            '        AddToLineNo(rc)
            '    End If
            'End If

            'If ds.IsOrdered.Trim = "1" Then
            '    Dim ForIsOrd As String = String.Format("{0,12}{1}", " ", "ORDERED")
            '    objWriteSQD.WriteLine(ForIsOrd)
            '    objWriteINL.WriteLine(ForIsOrd)
            '    objWriteTMP.WriteLine(ForIsOrd)
            '    AddToLineNo(rc)
            'End If
            If ds.IsIMSPathData.Trim = "1" Then
                Dim ForIsImsPath As String = String.Format("{0,12}{1}", " ", "IMSPATHDATA")
                objWriteSQD.WriteLine(ForIsImsPath)
                objWriteINL.WriteLine(ForIsImsPath)
                objWriteTMP.WriteLine(ForIsImsPath)
                AddToLineNo(rc)
            End If
            'If ds.IsBeforeImage.Trim = "1" Then
            '    Dim ForIsBefore As String = String.Format("{0,12}{1}", " ", "BEFOREIMG")
            '    objWriteSQD.WriteLine(ForIsBefore)
            '    objWriteINL.WriteLine(ForIsBefore)
            '    objWriteTMP.WriteLine(ForIsBefore)
            '    AddToLineNo(rc)
            'End If
            If ds.IsSkipChangeCheck = "1" Then
                Dim ForIsSkipChg As String = String.Format("{0,12}{1}", " ", "BYPASS CHGCHECK")
                objWriteSQD.WriteLine(ForIsSkipChg)
                objWriteINL.WriteLine(ForIsSkipChg)
                objWriteTMP.WriteLine(ForIsSkipChg)
                AddToLineNo(rc)
            End If

            'If ds.ExceptionDatastore.Trim <> "" Then
            '    Dim ForIsOrd As String = String.Format("{0,12}{1}", " ", "EXCEPTION " & QuoteRes("EX_" & ds.DatastoreName))
            '    objWriteSQD.WriteLine(ForIsOrd)
            '    objWriteINL.WriteLine(ForIsOrd)
            '    objWriteTMP.WriteLine(ForIsOrd)
            '    AddToLineNo(rc)
            'End If

        Catch ex As Exception
            LogError(ex, "modGenerate wDSattrib")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore Attributes"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSattrib = Not rc.HasFatal

    End Function

    Function wDSattribException(ByRef rc As clsRcode, ByVal ds As clsDatastore) As Boolean

        Try
            If ds.ExceptionDatastore.Trim <> "" Then
                Dim ForExDS As String = String.Format("{0,12}{1}", " ", "EXCEPTION " & QuoteRes("EX_" & ds.DatastoreName))
                objWriteSQD.WriteLine(ForExDS)
                objWriteINL.WriteLine(ForExDS)
                objWriteTMP.WriteLine(ForExDS)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wDSattribException")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore Attributes Exception"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSattribException = Not rc.HasFatal

    End Function

    Function wDSsuffix(ByRef rc As clsRcode, ByRef ds As clsDatastore) As Boolean

        Try
            If ds.OperationType <> "" Then
                Dim FORds3 As String = String.Format("{0,12}{1}", " ", "FOR " & GetOperType(ds.OperationType))

                objWriteSQD.WriteLine(FORds3)
                objWriteINL.WriteLine(FORds3)
                objWriteTMP.WriteLine(FORds3)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wDSsuffix")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore Suffix"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wDSsuffix = Not rc.HasFatal

    End Function

    Function wStructSQD(ByRef rc As clsRcode, ByVal struct As clsStructure, ByVal DSsel As clsDSSelection) As Boolean

        Try
            Dim objEng As clsEngine = DSsel.ObjDatastore.Engine
            Dim objSys As clsSystem = objEng.ObjSystem

            objSys.LoadMe()

            Dim FORstr1 As String = String.Format("{0}{1}{2}", "DESCRIPTION ", _
            GetStrType(struct.StructureType), Quote(GetStrPath(struct, objEng)))
            Dim FORstr2 As String = String.Format("{0}{1}", "AS ", QuoteRes(struct.StructureName))
            'Dim PSix As String = String.Format("{0}", "/+")
            'Dim PSix2 As String = String.Format("{0}", "+/")

            If struct.StructureType = enumStructure.STRUCT_REL_DML Or struct.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                wStructDML(rc, struct, "SQD", DSsel)
                objWriteSQD.Write(FORstr2)
                AddToLineNo(rc, , False, False)
            Else
                If objSys.OSType = "z/OS" Then
                    objWriteSQD.Write(FORstr1 & " ")
                    objWriteSQD.Write(FORstr2)
                    AddToLineNo(rc, , False, False)
                Else
                    objWriteSQD.WriteLine(FORstr1)
                    AddToLineNo(rc, , False, False)
                    objWriteSQD.Write("               " & FORstr2)
                    AddToLineNo(rc, , False, False)
                End If
            End If

            If struct.StructureType <> enumStructure.STRUCT_COBOL_IMS And struct.StructureType <> enumStructure.STRUCT_IMS Then
                objWriteSQD.WriteLine(semi)
            Else
                objWriteSQD.WriteLine()
            End If
            AddToLineNo(rc, , False, False)

            objWriteSQD.WriteLine()
            AddToLineNo(rc, , False, False)

        Catch ex As Exception
            LogError(ex, "modGenerate wStructSQD")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Descriptions to SQD file"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenStruct
            rc.ErrorPath = pathSQD
            rc.ObjInode = struct
            ErrorComment(rc)
        End Try

        wStructSQD = Not rc.HasFatal

    End Function

    Function wStructINL(ByRef rc As clsRcode, ByVal struct As clsStructure, ByVal dssel As clsDSSelection) As Boolean

        Try
            Dim FORstr1 As String = String.Format("{0}{1}", "DESCRIPTION ", GetStrType(struct.StructureType))
            Dim PSix As String = String.Format("{0}", "/+")

            Dim FORstr2 As String
            If struct.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                FORstr2 = String.Format("{0,15}{1}", "AS ", QuoteRes(struct.StructureName))
            Else
                FORstr2 = String.Format("{0,15}{1}{2}", "AS ", QuoteRes(struct.StructureName), semi)
            End If

            Dim objReadStr As System.IO.StreamReader
            Dim InStream As System.IO.FileStream
            Dim TempString As String
            'Dim FORdml2 As String = String.Format("{0,25}{1}{2}", "SELECT * FROM ", struct.fPath1, semi)
            Dim PSix2 As String = String.Format("{0}", "+/")

            '/// now write file or dml "select"
            If struct.StructureType = enumStructure.STRUCT_REL_DML Or struct.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                'objWriteINL.WriteLine(FORdml2)
                'AddToLineNo(rc, False, , False)
                'objWriteINL.WriteLine(PSix2)
                'AddToLineNo(rc, False, , False)
                wStructDML(rc, struct, "INL", dssel)
            Else
                objWriteINL.WriteLine(FORstr1)
                AddToLineNo(rc, False, , False)
                objWriteINL.WriteLine(PSix)
                AddToLineNo(rc, False, , False)
                If System.IO.File.Exists(struct.fPath1) = True Then
                    InStream = System.IO.File.OpenRead(struct.fPath1)
                    objReadStr = New System.IO.StreamReader(InStream)
                Else
                    rc.HasError = True
                    rc.ObjInode = struct
                    rc.ErrorLocation = enumErrorLocation.ModGenStruct
                    rc.ReturnCode = "File " & rc.Path & " does not exist"
                    rc.ErrorPath = pathINL
                    If struct.fPath1.StartsWith("%") = True Then
                        GoTo SubstGoTo
                    Else
                        rc.Path = struct.fPath1
                        GoTo ErrorGoTo
                    End If
                End If
                '/// now read and write the structure file
                While (objReadStr.Peek() > -1)
                    TempString = objReadStr.ReadLine
                    If struct.StructureType = enumStructure.STRUCT_COBOL_IMS Then
                        objWriteINL.WriteLine(TempString)
                    Else
                        'added 10/19/11 for making inline file correct for subsets
                        'no undefined dssel fields are written
                        'For Each fld As clsField In dssel.DSSelectionFields
                        '    Dim fldname As String = fld.FieldName
                        '    If TempString.Contains(fldname) Then
                        objWriteINL.WriteLine(TempString)
                        '        GoTo nextTempstr
                        '    End If
                        'Next
                    End If
nextTempstr:        AddToLineNo(rc, False, , False)
                End While
SubstGoTo:      objWriteINL.WriteLine(PSix2)
                AddToLineNo(rc, False, , False)
            End If
ErrorGoTo:
            objWriteINL.WriteLine(FORstr2)
            AddToLineNo(rc, False, , False)
            objWriteINL.WriteLine()
            AddToLineNo(rc, False, , False)


        Catch ex As Exception
            LogError(ex, "modGenerate wStructINL")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Description to INL file"
            rc.ReturnCode = ex.Message
            rc.ErrorPath = pathINL
            ErrorComment(rc)
        End Try

        wStructINL = Not rc.HasFatal

    End Function

    Function wStructTMP(ByRef rc As clsRcode, ByVal struct As clsStructure, ByVal DSsel As clsDSSelection) As Boolean

        Try
            Dim FORstr1 As String = String.Format("{0}{1}", "DESCRIPTION " & GetStrType(struct.StructureType), _
            Quote(GetStrPath(struct, DSsel.ObjDatastore.Engine)))
            Dim FORstr2 As String = String.Format("{0,15}{1}", "AS ", QuoteRes(struct.StructureName))
            Dim PSix As String = String.Format("{0}", "/+")
            Dim PSix2 As String = String.Format("{0}", "+/")
            'Dim FORdml1 As String = String.Format("{0}{1}", "DESCRIPTION ", GetStrType(struct.StructureType))
            'Dim FORdml2 As String = String.Format("{0,25}{1}{2}", "SELECT * FROM ", struct.fPath1, semi)

            If struct.StructureType = enumStructure.STRUCT_REL_DML Or struct.StructureType = enumStructure.STRUCT_REL_DML_FILE Then
                'objWriteTMP.WriteLine(FORdml1)
                'AddToLineNo(rc, False, False)
                'objWriteTMP.WriteLine(PSix)
                'AddToLineNo(rc, False, False)
                'objWriteTMP.WriteLine(FORdml2)
                'AddToLineNo(rc, False, False)
                'objWriteTMP.WriteLine(PSix2)
                'AddToLineNo(rc, False, False)
                wStructDML(rc, struct, "TMP", DSsel)
            Else
                objWriteTMP.WriteLine(FORstr1)
                AddToLineNo(rc, False, False)
            End If
            objWriteTMP.Write(FORstr2)
            If struct.StructureType <> enumStructure.STRUCT_COBOL_IMS And struct.StructureType <> enumStructure.STRUCT_IMS Then
                objWriteTMP.WriteLine(semi)
            Else
                objWriteTMP.WriteLine()
            End If
            AddToLineNo(rc, False, False)
            objWriteTMP.WriteLine()
            AddToLineNo(rc, False, False)

        Catch ex As Exception
            LogError(ex, "modGenerate wStructTMP")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Description to TMP file"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenStruct
            rc.ErrorPath = pathTMP
            rc.ObjInode = struct
            If struct.fPath1.StartsWith("%") = False Then
                rc.Path = struct.fPath1
            End If
            ErrorComment(rc)
        End Try

        wStructTMP = Not rc.HasFatal

    End Function

    Function wStructDML(ByRef rc As clsRcode, ByVal struct As clsStructure, ByVal CalledFrom As String, ByVal dssel As clsDSSelection) As Boolean

        Try
            Dim Selection As String = ""
            Dim BuildString As New System.Text.StringBuilder
            Dim fld As clsField

            BuildString.Append("SELECT ")
            For i As Integer = 0 To struct.ObjFields.Count - 1
                fld = struct.ObjFields(i)
                BuildString.Append(fld.FieldName) 'QuoteRes()
                If i < struct.ObjFields.Count - 1 Then
                    BuildString.Append(", ")
                End If
            Next
            BuildString.Append(" FROM ")

            Dim PSix As String = String.Format("{0}", "/+")
            Dim PSix2 As String = String.Format("{0}", "+/")
            Dim FORdml1 As String = String.Format("{0}{1}", "DESCRIPTION ", GetStrType(struct.StructureType))
            Dim FORdml2 As String = String.Format("{0,25}{1}{2}", BuildString.ToString, _
            GetStrPath(struct, dssel.ObjDatastore.Engine), semi)   'Quote()

            Select Case CalledFrom
                Case "TMP"
                    objWriteTMP.WriteLine(FORdml1)
                    AddToLineNo(rc, False, False)
                    objWriteTMP.WriteLine(PSix)
                    AddToLineNo(rc, False, False)
                    objWriteTMP.WriteLine(FORdml2)
                    AddToLineNo(rc, False, False)
                    objWriteTMP.WriteLine(PSix2)
                    AddToLineNo(rc, False, False)
                    objWriteTMP.WriteLine()
                    AddToLineNo(rc, False, False)
                Case "SQD"
                    objWriteSQD.WriteLine(FORdml1)
                    AddToLineNo(rc, False, False)
                    objWriteSQD.WriteLine(PSix)
                    AddToLineNo(rc, False, False)
                    objWriteSQD.WriteLine(FORdml2)
                    AddToLineNo(rc, False, False)
                    objWriteSQD.WriteLine(PSix2)
                    AddToLineNo(rc, False, False)
                    objWriteSQD.WriteLine()
                    AddToLineNo(rc, False, False)
                Case "INL"
                    objWriteINL.WriteLine(FORdml1)
                    AddToLineNo(rc, False, False)
                    objWriteINL.WriteLine(PSix)
                    AddToLineNo(rc, False, False)
                    objWriteINL.WriteLine(FORdml2)
                    AddToLineNo(rc, False, False)
                    objWriteINL.WriteLine(PSix2)
                    AddToLineNo(rc, False, False)
                    objWriteINL.WriteLine()
                    AddToLineNo(rc, False, False)
            End Select


        Catch ex As Exception
            LogError(ex, "modGenerate wStructDML")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing DML Description, check ODBC Connection"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenStruct
            rc.ErrorPath = pathTMP
            rc.ObjInode = struct
            If struct.fPath1.StartsWith("%") = False Then
                rc.Path = struct.fPath1
            End If
            ErrorComment(rc)
        End Try

        wStructDML = Not rc.HasFatal

    End Function

    Function wIMSprefix(ByRef RC As clsRcode, ByVal struct As clsStructure, ByVal dsSel As clsDSSelection) As Boolean

        Try
            Dim objEng As clsEngine = dsSel.ObjDatastore.Engine
            Dim objSys As clsSystem = objEng.ObjSystem

            Dim DBDformatSQD As String = String.Format("{0}{1}", "DESCRIPTION IMSDBD " & Quote(GetStrPath(struct, objEng, True)), _
            " AS " & struct.IMSDBName & semi)
            Dim DBDformatTMP As String = String.Format("{0}{1}", "DESCRIPTION IMSDBD " & Quote(GetStrPath(struct, objEng, True)), _
            " AS " & QuoteRes(struct.IMSDBName) & semi)
            Dim PSix As String = String.Format("{0}", "/+")
            Dim PSix2 As String = String.Format("{0}", "+/")
            Dim DBDformatINL As String = String.Format("{0}", "DESCRIPTION IMSDBD")
            Dim DBDformatINLend As String = String.Format("{0,12}{1}{2}", " ", "AS " & QuoteRes(struct.IMSDBName), semi)
            Dim objReadStr As System.IO.StreamReader
            Dim InStream As System.IO.FileStream
            Dim TempString As String

            objWriteSQD.WriteLine(DBDformatSQD)
            objWriteINL.WriteLine(DBDformatINL)
            objWriteTMP.WriteLine(DBDformatTMP)
            AddToLineNo(RC)
            objWriteINL.WriteLine(PSix)
            AddToLineNo(RC, False, , False)

            If System.IO.File.Exists(struct.fPath2) = True Then
                InStream = System.IO.File.OpenRead(struct.fPath2)
                objReadStr = New System.IO.StreamReader(InStream)
            Else
                RC.HasError = True
                If struct.fPath1.StartsWith("%") = False Then
                    RC.Path = struct.fPath2
                End If
                RC.ErrorLocation = enumErrorLocation.ModGenStruct
                'RC.ReturnCode = "File " & RC.ErrorPath & " does not exist"
                If struct.fPath2.StartsWith("%") = True Then
                    GoTo SubstGoTo
                Else
                    GoTo ErrorGoTo
                End If
            End If

            While (objReadStr.Peek() > -1)
                TempString = objReadStr.ReadLine
                objWriteINL.WriteLine(TempString)
                AddToLineNo(RC, False, , False)
            End While
SubstGoTo:  objWriteINL.WriteLine(PSix2)
            AddToLineNo(RC, False, , False)
            objWriteINL.WriteLine(DBDformatINLend)
            AddToLineNo(RC, False, , False)
            objWriteINL.WriteLine()
            AddToLineNo(RC, False, , False)

ErrorGoTo:
        Catch ex As Exception
            LogError(ex, "modGenerate wIMSprefix")
            RC.HasError = True
            RC.ErrorCount += 1
            RC.LocalErrorMsg = "Error while writing IMS Header"
            RC.ReturnCode = ex.Message
            RC.ObjInode = struct
            RC.ErrorPath = pathINL
            ErrorComment(RC)
        End Try

        wIMSprefix = Not RC.HasFatal

    End Function

    Function wIMSsuffix(ByRef RC As clsRcode, ByVal struct As clsStructure) As Boolean

        Try
            Dim Line1format As String = String.Format("{0,12}{1}{2}", " ", "FOR SEGMENT ", QuoteRes(struct.SegmentName))
            Dim Line2Format As String = String.Format("{0,12}{1}{2}", " ", "IN DATABASE ", QuoteRes(struct.IMSDBName) & semi)

            objWriteSQD.WriteLine(Line1format)
            objWriteINL.WriteLine(Line1format)
            objWriteTMP.WriteLine(Line1format)
            AddToLineNo(RC)
            objWriteSQD.WriteLine(Line2Format)
            objWriteINL.WriteLine(Line2Format)
            objWriteTMP.WriteLine(Line2Format)
            AddToLineNo(RC)
            wBlankLine(RC)

        Catch ex As Exception
            LogError(ex, "modGenerate wIMSprefix")
            RC.HasError = True
            RC.ErrorCount += 1
            RC.LocalErrorMsg = "Error while writing IMS Segment info"
            RC.ReturnCode = ex.Message
            RC.ErrorLocation = enumErrorLocation.ModGenStruct
            RC.ErrorPath = pathSQD
            If struct.fPath2.StartsWith("%") = False Then
                RC.Path = struct.fPath2
            End If
            ErrorComment(RC)
        End Try

        wIMSsuffix = Not RC.HasFatal

    End Function

    Function wProc(ByRef rc As clsRcode, ByVal Proc As clsTask, Optional ByVal First As Boolean = True, Optional ByVal SourceDS As String = "") As Boolean

        Try
            Dim FORprocBegin As String = String.Format("{0}{1}{2}", "CREATE PROC ", QuoteRes(Proc.TaskName), " AS SELECT")
            Dim FORprocEnd As String = String.Format("{0}{1}", "FROM ", QuoteRes(SourceDS) & semi)

            If First = True Then
                wBlankLine(rc)
                objWriteSQD.WriteLine(FORprocBegin)
                objWriteINL.WriteLine(FORprocBegin)
                objWriteTMP.WriteLine(FORprocBegin)
                AddToLineNo(rc)
                objWriteSQD.WriteLine("{")
                objWriteINL.WriteLine("{")
                objWriteTMP.WriteLine("{")
                AddToLineNo(rc)
            Else
                objWriteSQD.WriteLine("}")
                objWriteINL.WriteLine("}")
                objWriteTMP.WriteLine("}")
                AddToLineNo(rc)
                objWriteSQD.WriteLine(FORprocEnd)
                objWriteINL.WriteLine(FORprocEnd)
                objWriteTMP.WriteLine(FORprocEnd)
                AddToLineNo(rc)
                wBlankLine(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wProc")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Procedure"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenProc
            rc.ErrorPath = pathSQD
            rc.ObjInode = Proc
            ErrorComment(rc)
        End Try

        wProc = Not rc.HasFatal

    End Function

    Function wDebugHead(ByRef rc As clsRcode, ByVal proc As clsTask) As Boolean

        Try
            Dim FORout As String = String.Format("{0}{1}{2}", "--OUTMSG(0,'-------------------------- ", QuoteRes(proc.TaskName), _
            " Beginning --------------------')")

            objWriteSQD.WriteLine(FORout)
            objWriteINL.WriteLine(FORout)
            objWriteTMP.WriteLine(FORout)

        Catch ex As Exception
            LogError(ex, "modGenerate wDebugHead")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = ("Error while writing Debug OutMessages")
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenProc
            rc.ErrorPath = pathSQD
            rc.ObjInode = proc
            ErrorComment(rc)
        End Try

        wDebugHead = Not rc.HasFatal

    End Function

    Function wDebug(ByRef rc As clsRcode, ByVal map As clsMapping, Optional ByVal commentOut As Boolean = True) As Boolean

        Try
            Dim SrcStr As String = map.SourceDataStore & "." & map.SourceParent & "." & CType(map.MappingSource, clsField).FieldName
            Dim FORout As String

            If commentOut = True Then
                FORout = String.Format("{0}{1}{2}", "--OUTMSG(0,STRING('", SrcStr & " = ',", SrcStr & "))")
            Else
                FORout = String.Format("{0}{1}{2}", "OUTMSG(0,STRING('", SrcStr & " = ',", SrcStr & "))")
            End If

            objWriteSQD.WriteLine(FORout)
            objWriteINL.WriteLine(FORout)
            objWriteTMP.WriteLine(FORout)

        Catch ex As Exception
            LogError(ex, "modGenerate wDebug")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Debug Outmessages"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenProc
            rc.ErrorPath = pathSQD
            rc.ObjInode = map
            ErrorComment(rc)
        End Try

        wDebug = Not rc.HasFatal

    End Function

    Function wGlobal(ByRef rc As clsRcode, ByVal ds As clsDatastore, ByVal AttrType As String) As Boolean

        Try
            Dim Command As String = ""
            Dim Arg1 As String = ""
            Dim Arg2 As String = ""

            Select Case AttrType
                Case "EXTTYPECHAR"
                    Command = "IFSPACE"
                    Arg1 = ds.DatastoreName & ".ALLCHAR"
                    Arg2 = ds.ExtTypeChar  '"DT" & 

                Case "IFNULLCHAR"
                    Command = "IFNULL"
                    Arg1 = ds.DatastoreName & ".ALLCHAR"
                    Arg2 = ds.IfNullChar

                Case "INVALIDCHAR"
                    Command = "INVALID"
                    Arg1 = ds.DatastoreName & ".ALLCHAR"
                    Arg2 = ds.InValidChar

                Case "EXTTYPENUM"
                    Command = "IFSPACE"
                    Arg1 = ds.DatastoreName & ".ALLNUM"
                    Arg2 = ds.ExtTypeNum   '"DT" & 

                Case "IFNULLNUM"
                    Command = "IFNULL"
                    Arg1 = ds.DatastoreName & ".ALLNUM"
                    Arg2 = ds.IfNullNum

                Case "INVALIDNUM"
                    Command = "INVALID"
                    Arg1 = ds.DatastoreName & ".ALLNUM"
                    Arg2 = ds.InValidNum
            End Select

            Dim FORout As String = String.Format("{0}{1}{2}", Command & " ", Arg1 & " ", Arg2 & semi)

            objWriteSQD.WriteLine(FORout)
            objWriteINL.WriteLine(FORout)
            objWriteTMP.WriteLine(FORout)
            wBlankLine(rc)


        Catch ex As Exception
            LogError(ex, "modGenerate wGlobal")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Datastore-wide field Properties"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenDS
            rc.ErrorPath = pathSQD
            rc.ObjInode = ds
            ErrorComment(rc)
        End Try

        wGlobal = Not rc.HasFatal

    End Function

    Function wFA(ByRef rc As clsRcode, ByVal fldName As String, ByVal fld As clsField, ByVal FAtype As enumFieldAttributes) As Boolean

        Try
            Dim Command As String = ""
            Dim Arg1 As String = ""
            Dim Arg2 As String = ""

            Select Case FAtype
                Case enumFieldAttributes.ATTR_DATEFORMAT
                    Command = "DATEFORMAT"
                    Arg1 = fldName
                    Arg2 = fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT).ToString

                Case enumFieldAttributes.ATTR_EXTTYPE
                    Command = "EXTTYPE"
                    Arg1 = fldName
                    Arg2 = "DT" & fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE).ToString

                Case enumFieldAttributes.ATTR_IDENTVAL
                    Command = "IDENT"
                    Arg1 = fldName
                    Arg2 = fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL).ToString

                Case enumFieldAttributes.ATTR_INITVAL
                    Command = "INITIALIZE"
                    Arg1 = fldName
                    Arg2 = QuoteAttr(fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL).ToString)

                Case enumFieldAttributes.ATTR_INVALID
                    Command = "INVALID"
                    Arg1 = fldName
                    Arg2 = fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID).ToString
                    If Arg2 <> "SETNULL" And _
                    Arg2 <> "SETSPACE" And _
                    Arg2 <> "SETZERO" And _
                    Arg2 <> "SETEXP" Then
                        Arg2 = "'" & Arg2 & "'"
                    End If

                Case enumFieldAttributes.ATTR_LABEL
                    Command = "LABEL"
                    Arg1 = fldName
                    Arg2 = QuoteAttr(fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL).ToString)

                Case enumFieldAttributes.ATTR_RETYPE
                    Command = "RETYPE"
                    Arg1 = fldName
                    Arg2 = "DT" & fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE).ToString
                Case Else
                    '/// these attributes are defined in the structure itself
                    '/// not in the datastore.
            End Select
            Dim FORout As String = String.Format("{0}{1}{2}", Command & " ", QuoteRes(Arg1) & " ", Arg2 & semi)

            objWriteSQD.WriteLine(FORout)
            objWriteINL.WriteLine(FORout)
            objWriteTMP.WriteLine(FORout)
            wBlankLine(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wFA")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Field Attributes"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenMap
            rc.ErrorPath = pathSQD
            rc.ObjInode = fld
            ErrorComment(rc)
        End Try

        wFA = Not rc.HasFatal

    End Function

    Function wJoin(ByRef rc As clsRcode, ByVal Join As clsTask, ByVal dsName As String, Optional ByVal First As Boolean = True) As Boolean

        Try
            Dim FORJoinBegin As String = String.Format("{0}{1}{2}", "CREATE JOIN ", QuoteRes(Join.TaskName), " AS SELECT")
            Dim FORJoinEnd As String = String.Format("{0}{1}", "FROM ", QuoteRes(dsName))

            If First = True Then
                wBlankLine(rc)
                objWriteSQD.WriteLine(FORJoinBegin)
                objWriteINL.WriteLine(FORJoinBegin)
                objWriteTMP.WriteLine(FORJoinBegin)
                AddToLineNo(rc)
            Else
                objWriteSQD.WriteLine(FORJoinEnd)
                objWriteINL.WriteLine(FORJoinEnd)
                objWriteTMP.WriteLine(FORJoinEnd)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wJoin")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Join"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenJoin
            rc.ErrorPath = pathSQD
            rc.ObjInode = Join
            ErrorComment(rc)
        End Try

        wJoin = Not rc.HasFatal

    End Function

    Function wLookup(ByRef rc As clsRcode, ByVal LU As clsTask, Optional ByVal dsName As String = "", Optional ByVal First As Boolean = True) As Boolean

        Try
            Dim FORLUBegin As String = String.Format("{0}{1}{2}", "CREATE LOOKUP ", QuoteRes(LU.TaskName), " AS SELECT")
            Dim FORLUEnd As String = String.Format("{0}{1}", "FROM ", QuoteRes(dsName))

            If First = True Then
                wBlankLine(rc)
                objWriteSQD.WriteLine(FORLUBegin)
                objWriteINL.WriteLine(FORLUBegin)
                objWriteTMP.WriteLine(FORLUBegin)
                AddToLineNo(rc)
            Else
                objWriteSQD.WriteLine(FORLUEnd)
                objWriteINL.WriteLine(FORLUEnd)
                objWriteTMP.WriteLine(FORLUEnd)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wLookup")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Lookup"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenLU
            rc.ErrorPath = pathSQD
            rc.ObjInode = LU
            ErrorComment(rc)
        End Try

        wLookup = Not rc.HasFatal

    End Function

    Function wLookDS(ByRef rc As clsRcode, ByVal DS As clsDatastore) As Boolean

        Try
            Dim FORLUBegin As String = String.Format("{0}{1}{2}", "CREATE LOOKUP ", QuoteRes(DS.DatastoreName), " AS SELECT")
            Dim FORLUEnd As String = String.Format("{0,12}{1}{2}", " ", "FROM ", QuoteRes(DS.DatastoreName))
            Dim prefix As String = " "
            Dim first As Boolean = True
            Dim ForSel As String

            wSemiLine(rc)
            objWriteSQD.WriteLine(FORLUBegin)
            objWriteINL.WriteLine(FORLUBegin)
            objWriteTMP.WriteLine(FORLUBegin)
            AddToLineNo(rc)

            For Each dssel As clsDSSelection In DS.ObjSelections
                For Each fld As clsField In dssel.DSSelectionFields
                    If first = True Then
                        ForSel = String.Format("{0,19}{1}{2}", prefix, DS.DatastoreName, "." & fld.FieldName)
                        first = False
                    Else
                        prefix = ","
                        ForSel = String.Format("{0,19}{1}{2}", prefix, DS.DatastoreName, "." & fld.FieldName)
                    End If
                    objWriteSQD.WriteLine(ForSel)
                    objWriteINL.WriteLine(ForSel)
                    objWriteTMP.WriteLine(ForSel)
                    AddToLineNo(rc)
                Next
            Next

            objWriteSQD.WriteLine(FORLUEnd)
            objWriteINL.WriteLine(FORLUEnd)
            objWriteTMP.WriteLine(FORLUEnd)
            AddToLineNo(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wLookup")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Lookup"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenLU
            rc.ErrorPath = pathSQD
            rc.ObjInode = DS
            ErrorComment(rc)
        End Try

        wLookDS = Not rc.HasFatal

    End Function

    Function wLUfield(ByRef rc As clsRcode, ByVal LU As clsTask, ByVal map As clsMapping) As Boolean  ', ByVal first As Boolean

        Try
            Dim Prefix As String = "  "
            Dim SrcStr As String = ""
            Dim TgtStr As String = ""
            Dim FORtgt As String
            Dim FORsrc As String
            Dim SrcLen As Integer
            Dim TgtLen As Integer
            Dim TotLen As Integer
            Dim SrcLong As Boolean
            Dim TgtLong As Boolean

            Select Case SourceLevel
                Case enumMappingLevel.ShowAll
                    SrcStr = "LOOKFLD(" & LU.TaskName & ", " & map.SourceDataStore & "." & map.SourceParent & "." & _
                    CType(map.MappingSource, clsField).FieldName & ")"
                Case enumMappingLevel.ShowDesc
                    SrcStr = "LOOKFLD(" & LU.TaskName & ", " & map.SourceParent & "." & _
                    CType(map.MappingSource, clsField).FieldName & ")"
                Case enumMappingLevel.ShowFld
                    SrcStr = "LOOKFLD(" & LU.TaskName & ", " & CType(map.MappingSource, clsField).FieldName & ")"
            End Select
            Select Case TargetLevel
                Case enumMappingLevel.ShowAll
                    TgtStr = map.TargetDataStore & "." & map.TargetParent & "." & CType(map.MappingTarget, clsField).FieldName
                Case enumMappingLevel.ShowDesc
                    TgtStr = map.TargetParent & "." & CType(map.MappingTarget, clsField).FieldName
                Case enumMappingLevel.ShowFld
                    TgtStr = CType(map.MappingTarget, clsField).FieldName
            End Select

            If ObjSys.OSType = "z/OS" Then
                SrcStr = SrcStr.Trim
                TgtStr = TgtStr.Trim
                SrcLen = SrcStr.Length
                TgtLen = TgtStr.Length
                TotLen = SrcLen + TgtLen

                If SrcLen > 33 Then
                    SrcLong = True
                Else
                    SrcLong = False
                End If
                If TgtLen > 33 Then
                    TgtLong = True
                Else
                    TgtLong = False
                End If

                If TgtLong = False Then
                    FORtgt = String.Format("{0,2}{1,-33}", Prefix, TgtStr)
                Else
                    FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                End If

                If SrcLong = False Then
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,38}{1}", " = ", SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                Else
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,70}", " = " & SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                End If
                If TotLen > 66 Then
                    objWriteSQD.WriteLine(FORtgt)
                    objWriteINL.WriteLine(FORtgt)
                    objWriteTMP.WriteLine(FORtgt)
                    AddToLineNo(rc)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(rc)
                Else
                    objWriteSQD.Write(FORtgt)
                    objWriteINL.Write(FORtgt)
                    objWriteTMP.Write(FORtgt)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(rc)
                End If
            Else
                FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                objWriteSQD.Write(FORtgt)
                objWriteINL.Write(FORtgt)
                objWriteTMP.Write(FORtgt)
                objWriteSQD.WriteLine(FORsrc)
                objWriteINL.WriteLine(FORsrc)
                objWriteTMP.WriteLine(FORsrc)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wLUfield")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing LookFld function"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenProc
            rc.ErrorPath = pathSQD
            rc.ObjInode = map
            ErrorComment(rc)
        End Try

        wLUfield = Not rc.HasFatal

    End Function

    Function wMap(ByRef RC As clsRcode, ByVal map As clsMapping, Optional ByVal NoSrc As Boolean = False) As Boolean  ', ByVal First As Boolean

        Try
            Dim Prefix As String = "  "
            Dim SrcStr As String = ""
            Dim TgtStr As String = ""
            Dim FORtgt As String
            Dim FORsrc As String
            Dim SrcLen As Integer
            Dim TgtLen As Integer
            Dim TotLen As Integer
            Dim SrcLong As Boolean
            Dim TgtLong As Boolean

            If NoSrc = True Then
                SrcStr = "*** No Source Present ***"
                RC.HasError = True
                RC.ErrorCount += 1
                RC.LocalErrorMsg = "Error - Target with no Source"
                RC.ErrorLocation = enumErrorLocation.ModGenMap
                RC.ErrorPath = pathSQD
                RC.ObjInode = map
                ErrorComment(RC, True)
            Else
                If map.SourceType = enumMappingType.MAPPING_TYPE_FIELD Then
                    Select Case SourceLevel
                        Case enumMappingLevel.ShowAll
                            SrcStr = map.SourceDataStore & "." & map.SourceParent & "." & CType(map.MappingSource, clsField).FieldName
                        Case enumMappingLevel.ShowDesc
                            SrcStr = map.SourceParent & "." & CType(map.MappingSource, clsField).FieldName
                        Case enumMappingLevel.ShowFld
                            SrcStr = CType(map.MappingSource, clsField).FieldName
                    End Select
                Else
                    If map.SourceType = enumMappingType.MAPPING_TYPE_VAR Or map.SourceType = enumMappingType.MAPPING_TYPE_WORKVAR Then
                        SrcStr = CType(map.MappingSource, clsVariable).VariableName
                    End If
                End If

            End If

            If map.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
                Select Case TargetLevel
                    Case enumMappingLevel.ShowAll
                        TgtStr = map.TargetDataStore & "." & map.TargetParent & "." & CType(map.MappingTarget, clsField).FieldName
                    Case enumMappingLevel.ShowDesc
                        TgtStr = map.TargetParent & "." & CType(map.MappingTarget, clsField).FieldName
                    Case enumMappingLevel.ShowFld
                        TgtStr = CType(map.MappingTarget, clsField).FieldName
                End Select
            Else
                If map.TargetType = enumMappingType.MAPPING_TYPE_VAR Or map.TargetType = enumMappingType.MAPPING_TYPE_WORKVAR Then
                    TgtStr = CType(map.MappingTarget, clsVariable).VariableName
                End If
            End If

            SrcLen = SrcStr.Length
            TgtLen = TgtStr.Length
            TotLen = SrcLen + TgtLen

            If SrcLen > 33 Then
                SrcLong = True
            Else
                SrcLong = False
            End If

            If ObjSys.OSType = "z/OS" Then
                If TgtLen > 33 Then
                    TgtLong = True
                Else
                    TgtLong = False
                End If

                If TgtLong = False Then
                    FORtgt = String.Format("{0,2}{1,-33}", Prefix, TgtStr)
                Else
                    FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                End If

                If SrcLong = False Then
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,38}{1}", " = ", SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                Else
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,70}", " = " & SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                End If

                If map.MappingDesc IsNot Nothing Then
                    If map.MappingDesc.Trim <> "" Then
                        wComment(RC, map.MappingDesc)
                    End If
                End If

                If TotLen > 66 Then
                    objWriteSQD.WriteLine(FORtgt)
                    objWriteINL.WriteLine(FORtgt)
                    objWriteTMP.WriteLine(FORtgt)
                    AddToLineNo(RC)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(RC)
                Else
                    objWriteSQD.Write(FORtgt)
                    objWriteINL.Write(FORtgt)
                    objWriteTMP.Write(FORtgt)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(RC)
                End If
            Else
                If TgtLen > 40 Then
                    TgtLong = True
                Else
                    TgtLong = False
                End If
                If TgtLong = False Then
                    FORtgt = String.Format("{0,2}{1,-40}", Prefix, TgtStr)
                Else
                    FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                End If
                FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                If map.MappingDesc IsNot Nothing Then
                    If map.MappingDesc.Trim <> "" Then
                        wComment(RC, map.MappingDesc)
                    End If
                End If
                objWriteSQD.Write(FORtgt)
                objWriteINL.Write(FORtgt)
                objWriteTMP.Write(FORtgt)
                objWriteSQD.WriteLine(FORsrc)
                objWriteINL.WriteLine(FORsrc)
                objWriteTMP.WriteLine(FORsrc)
                AddToLineNo(RC)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wMap")
            RC.HasError = True
            RC.ErrorCount += 1
            RC.LocalErrorMsg = "Error while writing Mappings"
            RC.ReturnCode = ex.Message
            RC.ErrorLocation = enumErrorLocation.ModGenMap
            RC.ErrorPath = pathSQD
            RC.ObjInode = map
            ErrorComment(RC)
        End Try

        wMap = Not RC.HasFatal

    End Function

    Function wMapSrcOnly(ByRef RC As clsRcode, ByVal map As clsMapping, ByVal First As Boolean) As Boolean

        Try
            Dim Prefix As String
            Dim SrcStr As String = ""

            Select Case SourceLevel
                Case enumMappingLevel.ShowAll
                    SrcStr = map.SourceDataStore & "." & map.SourceParent & "." & CType(map.MappingSource, clsField).FieldName  '
                Case enumMappingLevel.ShowDesc
                    SrcStr = map.SourceParent & "." & CType(map.MappingSource, clsField).FieldName  '
                Case enumMappingLevel.ShowFld
                    SrcStr = CType(map.MappingSource, clsField).FieldName  '
            End Select

            If First = True Then
                Prefix = "  "
            Else
                Prefix = " ,"
            End If
            Dim FORsrc As String = String.Format("{0,2}{1}", Prefix, SrcStr)

            If map.MappingDesc.Trim <> "" Then
                wComment(RC, map.MappingDesc)
            End If
            objWriteSQD.WriteLine(FORsrc)
            objWriteINL.WriteLine(FORsrc)
            objWriteTMP.WriteLine(FORsrc)
            AddToLineNo(RC)

        Catch ex As Exception
            LogError(ex, "modGenerate wMapSrcOnly")
            RC.HasError = True
            RC.ErrorCount += 1
            RC.LocalErrorMsg = "Error while writing Mapping with Source only"
            RC.ReturnCode = ex.Message
            RC.ErrorLocation = enumErrorLocation.ModGenMap
            RC.ErrorPath = pathSQD
            RC.ObjInode = map
            ErrorComment(RC)
        End Try

        wMapSrcOnly = Not RC.HasFatal

    End Function

    Function wMapFun(ByRef rc As clsRcode, ByVal map As clsMapping) As Boolean

        Try
            Dim Prefix As String = "  "
            Dim SrcStr As String = ""
            Dim FORsrc As String

            Select Case map.SourceType
                Case enumMappingType.MAPPING_TYPE_CONSTANT
                    SrcStr = CType(map.MappingSource, Integer)
                Case enumMappingType.MAPPING_TYPE_FUN
                    SrcStr = CType(map.MappingSource, clsSQFunction).SQFunctionWithInnerText
                Case enumMappingType.MAPPING_TYPE_VAR
                    SrcStr = CType(map.MappingSource, clsVariable).VariableName
                Case enumMappingType.MAPPING_TYPE_WORKVAR
                    SrcStr = CType(map.MappingSource, clsVariable).VariableName
            End Select

            FORsrc = String.Format("{0,2}{1}", Prefix, SrcStr)

            If map.MappingDesc.Trim <> "" Then
                wComment(rc, map.MappingDesc)
            End If
            objWriteSQD.WriteLine(FORsrc)
            objWriteINL.WriteLine(FORsrc)
            objWriteTMP.WriteLine(FORsrc)
            AddToLineNo(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wMapFun")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Mapping Function"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenMap
            rc.ErrorPath = pathSQD
            rc.ObjInode = map
            ErrorComment(rc)
        End Try

        wMapFun = Not rc.HasFatal

    End Function

    Function wMapFunWTgt(ByRef rc As clsRcode, ByVal map As clsMapping, Optional ByVal NoSrc As Boolean = False) As Boolean   ', ByVal first As Boolean

        Try
            Dim Prefix As String = "  "
            Dim SrcStr As String = ""
            Dim TgtStr As String = ""
            Dim FORtgt As String = ""
            Dim FORsrc As String = ""
            Dim SrcLen As Integer = 0
            Dim TgtLen As Integer = 0
            Dim TotLen As Integer = 0
            Dim SrcLong As Boolean
            Dim TgtLong As Boolean

            '**** for debugging of Function Mappings
            If DBGMap = True Then
                Log("Mapping Source Type = " & map.SourceType.ToString)
                If map.SourceType = enumMappingType.MAPPING_TYPE_FUN Then
                    Log("Source = " & CType(map.MappingSource, clsSQFunction).SQFunctionWithInnerText)
                Else
                    Log("Source = " & map.MappingSource.ToString)
                End If
                Log("Mapping Target Type = " & map.TargetType.ToString)
                If map.TargetType = enumMappingType.MAPPING_TYPE_FIELD Then
                    Log("Target = " & CType(map.MappingTarget, clsField).FieldName)
                Else
                    Log("Target = " & map.MappingTarget.ToString)
                End If
                If rc IsNot Nothing Then
                    Log("rc object is Good")
                Else
                    Log("***** No rc object *****")
                End If
                Log("Mapping Level = " & map.IsMapped)
                Log("No Src = " & NoSrc.ToString)
            End If

            Select Case map.TargetType
                Case enumMappingType.MAPPING_TYPE_CONSTANT
                    TgtStr = CType(map.MappingTarget, clsVariable).VariableName
                Case enumMappingType.MAPPING_TYPE_FUN
                    TgtStr = CType(map.MappingTarget, clsSQFunction).SQFunctionWithInnerText
                Case enumMappingType.MAPPING_TYPE_VAR
                    TgtStr = CType(map.MappingTarget, clsVariable).VariableName
                Case enumMappingType.MAPPING_TYPE_WORKVAR
                    TgtStr = CType(map.MappingTarget, clsVariable).VariableName
                Case enumMappingType.MAPPING_TYPE_FIELD
                    Select Case TargetLevel
                        Case enumMappingLevel.ShowAll
                            TgtStr = map.TargetDataStore & "." & map.TargetParent & "." & _
                            CType(map.MappingTarget, clsField).FieldName
                        Case enumMappingLevel.ShowDesc
                            TgtStr = map.TargetParent & "." & CType(map.MappingTarget, clsField).FieldName
                        Case enumMappingLevel.ShowFld
                            TgtStr = CType(map.MappingTarget, clsField).FieldName
                    End Select
            End Select
            If DBGMap = True Then
                Log("TgtStr = " & TgtStr)
            End If


            If NoSrc = True Then
                SrcStr = "*** No Source Present ***"
                rc.HasError = True
                rc.ErrorCount += 1
                rc.LocalErrorMsg = "Error - Target with no Source"
                rc.ErrorLocation = enumErrorLocation.ModGenMap
                rc.ErrorPath = pathSQD
                rc.ObjInode = map
                ErrorComment(rc, True)
            Else
                Select Case map.SourceType
                    Case enumMappingType.MAPPING_TYPE_CONSTANT
                        SrcStr = CType(map.MappingSource, clsVariable).VariableName
                    Case enumMappingType.MAPPING_TYPE_FUN
                        SrcStr = CType(map.MappingSource, clsSQFunction).SQFunctionWithInnerText
                    Case enumMappingType.MAPPING_TYPE_VAR
                        SrcStr = CType(map.MappingSource, clsVariable).VariableName
                    Case enumMappingType.MAPPING_TYPE_WORKVAR
                        SrcStr = CType(map.MappingSource, clsVariable).VariableName
                End Select
            End If
            If DBGMap = True Then
                Log("SrcStr = " & SrcStr)
            End If


            SrcLen = SrcStr.Length
            TgtLen = TgtStr.Length
            TotLen = SrcLen + TgtLen

            If SrcLen > 33 Then
                SrcLong = True
            Else
                SrcLong = False
            End If

            If ObjSys.OSType = "z/OS" Then
                If TgtLen > 33 Then
                    TgtLong = True
                Else
                    TgtLong = False
                End If

                If TgtLong = False Then
                    FORtgt = String.Format("{0,2}{1,-33}", Prefix, TgtStr)
                Else
                    FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                End If

                If SrcLong = False Then
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,38}{1}", " = ", SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                Else
                    If TotLen > 66 Then
                        FORsrc = String.Format("{0,70}", " = " & SrcStr)
                    Else
                        FORsrc = String.Format("{0}{1}", " = ", SrcStr)
                    End If
                End If

                If map.MappingDesc IsNot Nothing Then
                    If map.MappingDesc <> "" Then
                        wComment(rc, map.MappingDesc)
                    End If
                End If

                If TotLen > 66 Then
                    objWriteSQD.WriteLine(FORtgt)
                    objWriteINL.WriteLine(FORtgt)
                    objWriteTMP.WriteLine(FORtgt)
                    AddToLineNo(rc)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(rc)
                Else
                    objWriteSQD.Write(FORtgt)
                    objWriteINL.Write(FORtgt)
                    objWriteTMP.Write(FORtgt)
                    objWriteSQD.WriteLine(FORsrc)
                    objWriteINL.WriteLine(FORsrc)
                    objWriteTMP.WriteLine(FORsrc)
                    AddToLineNo(rc)
                End If
            Else
                If TgtLen > 40 Then
                    TgtLong = True
                Else
                    TgtLong = False
                End If
                If TgtLong = False Then
                    FORtgt = String.Format("{0,2}{1,-40}", Prefix, TgtStr)
                Else
                    FORtgt = String.Format("{0,2}{1}", Prefix, TgtStr)
                End If
                FORsrc = String.Format("{0}{1}", " = ", SrcStr)

                If map.MappingDesc IsNot Nothing Then
                    If DBGMap = True Then
                        Log("MapDesc = " & map.MappingDesc)
                    End If
                    If map.MappingDesc <> "" Then
                        wComment(rc, map.MappingDesc)
                    End If
                End If
                objWriteSQD.Write(FORtgt)
                objWriteINL.Write(FORtgt)
                objWriteTMP.Write(FORtgt)
                objWriteSQD.WriteLine(FORsrc)
                objWriteINL.WriteLine(FORsrc)
                objWriteTMP.WriteLine(FORsrc)
                AddToLineNo(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wMapFunWTgt")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Mapping function with target field"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenMap
            rc.ErrorPath = pathSQD
            rc.ObjInode = map
            ErrorComment(rc)
        End Try

        wMapFunWTgt = Not rc.HasFatal

    End Function

    Function wVariable(ByRef rc As clsRcode, ByVal var As clsVariable) As Boolean

        Try
            Dim FORvar As String = String.Format("{0}{1}{2}", "DECLARE " & QuoteRes(var.VariableName), _
            " " & var.VarSize, " " & var.VarInitVal & semi)

            wBlankLine(rc)
            objWriteSQD.WriteLine(FORvar)
            objWriteINL.WriteLine(FORvar)
            objWriteTMP.WriteLine(FORvar)
            AddToLineNo(rc)
            wBlankLine(rc)
            'wSemiLine(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wVariable")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Variable"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenVar
            rc.ErrorPath = pathSQD
            rc.ObjInode = var
            ErrorComment(rc)
        End Try

        wVariable = Not rc.HasFatal

    End Function

    Function wMain(ByRef rc As clsRcode, ByVal Main As clsTask, ByVal SourceName As String, ByVal Last As Boolean, ByVal first As Boolean) As Boolean

        Try
            Dim Prefix As String = " "
            Dim TgtStr As String = ""
            Dim FORmainStr1 As String = String.Format("{0}", "PROCESS INTO")
            '/// change for CDCstore
            'If Main.Engine.Sources(1) IsNot Nothing Then
            '    If CType(Main.Engine.Sources(1), clsDatastore).DsAccessMethod = DS_ACCESSMETHOD_CDCSTORE Then
            '        FORmainStr1 = String.Format("{0}", "CHANGE INTO")
            '    End If
            'End If
            Dim FORmainStr3 As String = String.Format("{0}", "SELECT")
            Dim FORmainFunct As String = String.Format("{0}", CType(CType(Main.ObjMappings(0), clsMapping) _
            .MappingSource, clsSQFunction).SQFunctionWithInnerText)
            Dim FORmainStr4 As String = String.Format("{0}{1}", "FROM ", SourceName)

            wBlankLine(rc)
            If first = True Then
                objWriteSQD.WriteLine(FORmainStr1)
                objWriteINL.WriteLine(FORmainStr1)
                objWriteTMP.WriteLine(FORmainStr1)
                AddToLineNo(rc)
                For Each tgt As clsDatastore In ObjThis.Targets
                    TgtStr = tgt.DatastoreName
                    Dim FORmainStr2 As String = String.Format("{0,8}{1}", Prefix, QuoteRes(TgtStr))
                    objWriteSQD.WriteLine(FORmainStr2)
                    objWriteINL.WriteLine(FORmainStr2)
                    objWriteTMP.WriteLine(FORmainStr2)
                    AddToLineNo(rc)
                    Prefix = ","
                Next
            End If

            objWriteSQD.WriteLine(FORmainStr3)
            objWriteINL.WriteLine(FORmainStr3)
            objWriteTMP.WriteLine(FORmainStr3)
            AddToLineNo(rc)

            objWriteSQD.WriteLine("{")
            objWriteINL.WriteLine("{")
            objWriteTMP.WriteLine("{")
            AddToLineNo(rc)

            objWriteSQD.WriteLine(FORmainFunct)
            objWriteINL.WriteLine(FORmainFunct)
            objWriteTMP.WriteLine(FORmainFunct)
            AddToLineNo(rc)

            objWriteSQD.WriteLine("}")
            objWriteINL.WriteLine("}")
            objWriteTMP.WriteLine("}")
            AddToLineNo(rc)

            objWriteSQD.Write(FORmainStr4)
            objWriteINL.Write(FORmainStr4)
            objWriteTMP.Write(FORmainStr4)
            AddToLineNo(rc)
            If Last = True Then
                wSemiLine(rc)
                'Else
                '    wBlankLine(rc)
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wMain")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Main Procedure"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenMain
            rc.ErrorPath = pathSQD
            rc.ObjInode = Main
            ErrorComment(rc)
        End Try

        wMain = Not rc.HasFatal

    End Function

    Function wUnion(ByRef rc As clsRcode) As Boolean

        Try
            Dim FORunion As String = String.Format("{0}", "UNION")

            wBlankLine(rc)
            objWriteSQD.WriteLine(FORunion)
            objWriteINL.WriteLine(FORunion)
            objWriteTMP.WriteLine(FORunion)
            AddToLineNo(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wUnion")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Union"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenMain
            rc.ErrorPath = pathSQD
            rc.ObjInode = ObjThis
            ErrorComment(rc)
        End Try

        wUnion = Not rc.HasFatal

    End Function

    Function wInclude(ByRef rc As clsRcode, ByVal InStr As String) As Boolean
        Try
            Dim FORincl As String = String.Format("{0}{1}{2}", "#INCLUDE ", InStr, semi)

            wBlankLine(rc)
            objWriteSQD.WriteLine(FORincl)
            objWriteINL.WriteLine(FORincl)
            objWriteTMP.WriteLine(FORincl)
            AddToLineNo(rc)
            wBlankLine(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wInclude")
            rc.HasError = True
            rc.ErrorCount += 1
            rc.LocalErrorMsg = "Error while writing Include"
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenProc
            rc.ErrorPath = pathSQD
            rc.ObjInode = ObjThis
            ErrorComment(rc)
        End Try

        wInclude = Not rc.HasFatal

    End Function

#End Region

#Region " Helper Functions and subs "

    Function wBlankLine(ByRef rc As clsRcode) As Boolean

        Try
            objWriteSQD.WriteLine()
            objWriteINL.WriteLine()
            objWriteTMP.WriteLine()
            AddToLineNo(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wBlankLine")
            rc.HasError = True
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenWindows
            rc.ErrorPath = pathSQD
        End Try

        wBlankLine = Not rc.HasFatal

    End Function

    Function wSemiLine(ByRef rc As clsRcode) As Boolean

        Try
            objWriteSQD.WriteLine(semi)
            objWriteINL.WriteLine(semi)
            objWriteTMP.WriteLine(semi)
            AddToLineNo(rc)

        Catch ex As Exception
            LogError(ex, "modGenerate wSemiLine")
            rc.HasError = True
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenWindows
            rc.ErrorPath = pathSQD
        End Try

        wSemiLine = Not rc.HasFatal

    End Function

    Function wComment(ByRef rc As clsRcode, ByVal comment As String) As Boolean

        Dim overLenStr As String = ""
        Dim FirstPart As String = ""
        Dim index As Integer
        Dim count As Integer = 0

        Try
            If comment.Contains(Chr(9)) = False And _
            comment.Contains(Chr(10)) = False And _
            comment.Contains(Chr(12)) = False And _
            comment.Contains(Chr(13)) = False _
            Then        'comment.Length < 69 And _
                objWriteSQD.WriteLine("-- " & comment)
                objWriteINL.WriteLine("-- " & comment)
                objWriteTMP.WriteLine("-- " & comment)
                AddToLineNo(rc)
            Else
                For Each chara As Char In comment.ToCharArray
                    count += 1
                    If (chara = Chr(32) Or _
                    chara = Chr(9) Or _
                    chara = Chr(10) Or _
                    chara = Chr(12) Or _
                    chara = Chr(13)) Then       'And count < 69
                        index = count
                        If chara = Chr(9) Or _
                        chara = Chr(10) Or _
                        chara = Chr(12) Or _
                        chara = Chr(13) Then
                            Exit For
                        End If
                    End If
                Next
                FirstPart = Strings.Left(comment, index)

                objWriteSQD.WriteLine("-- " & FirstPart)
                objWriteINL.WriteLine("-- " & FirstPart)
                objWriteTMP.WriteLine("-- " & FirstPart)
                AddToLineNo(rc)

                overLenStr = Strings.Right(comment, (comment.Length - index) + 1)
                wComment(rc, overLenStr.Trim)
            End If


        Catch ex As Exception
            LogError(ex, "modGenerate wComment")
            rc.HasError = True
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenWindows
        End Try

        wComment = Not rc.HasFatal

    End Function

    Public Function ErrorComment(ByRef rc As clsRcode, Optional ByVal LocalOnly As Boolean = False) As Boolean

        Try
            objWriteSQD.WriteLine("--*** " & rc.LocalErrorMsg)
            objWriteINL.WriteLine("--*** " & rc.LocalErrorMsg)
            objWriteTMP.WriteLine("--*** " & rc.LocalErrorMsg)
            If LocalOnly = False Then
                objWriteSQD.WriteLine("--*** " & Strings.Left(rc.ReturnCode, 62))
                objWriteINL.WriteLine("--*** " & Strings.Left(rc.ReturnCode, 62))
                objWriteTMP.WriteLine("--*** " & Strings.Left(rc.ReturnCode, 62))
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate wErrorComment")
            rc.HasError = True
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenWindows
        End Try
        Return True

    End Function

    Public Sub AddToLineNo(ByRef rc As clsRcode, Optional ByVal AddToSQD As Boolean = True, Optional ByVal AddToINL As Boolean = True, Optional ByVal AddToTMP As Boolean = True)

        Try
            If AddToSQD = True Then
                rc.SQDline += 1
            End If
            If AddToINL = True Then
                rc.INLline += 1
            End If
            If AddToTMP = True Then
                rc.TMPline += 1
            End If

        Catch ex As Exception
            LogError(ex, "modGenerate AddToLineNo")
            rc.HasError = True
            rc.ReturnCode = ex.Message
            rc.ErrorLocation = enumErrorLocation.ModGenWindows
            rc.ErrorPath = pathSQD
        End Try

    End Sub

    Public Function GetStrType(ByVal strType As enumStructure) As String

        Dim retStr As String

        Select Case strType
            Case enumStructure.STRUCT_C
                retStr = "CSTRUCT "
            Case enumStructure.STRUCT_COBOL
                retStr = "COBOL "
            Case enumStructure.STRUCT_COBOL_IMS
                retStr = "COBOL "
            Case enumStructure.STRUCT_IMS
                retStr = "COBOL "
            Case enumStructure.STRUCT_REL_DDL
                retStr = "SQLDDL "
            Case enumStructure.STRUCT_REL_DML
                retStr = "SQLDML "
            Case enumStructure.STRUCT_REL_DML_FILE
                retStr = "SQLDML "
            Case enumStructure.STRUCT_XMLDTD
                retStr = "XMLDTD "
            Case Else
                retStr = ""
        End Select

        GetStrType = retStr

    End Function

    Public Function GetStrPath(ByVal str As clsStructure, ByVal eng As clsEngine, Optional ByVal IsDBD As Boolean = False) As String

        Dim DDstr As String = ""
        Dim objSys As clsSystem = eng.ObjSystem
        Dim cobLib As String = ""
        Dim ddlLib As String = ""
        Dim dtdLib As String = ""
        Dim incLib As String = ""
        Dim descLib As String = ""

        'Get CopyBook path on target
        If objSys.CopybookLib.Trim <> "" Then
            cobLib = objSys.CopybookLib
        End If
        If eng.CopybookLib.Trim <> "" Then
            cobLib = eng.CopybookLib
        End If

        'Get ddl path on target
        If objSys.DDLLib.Trim <> "" Then
            ddlLib = objSys.DDLLib
        End If
        If eng.DDLLib.Trim <> "" Then
            ddlLib = eng.DDLLib
        End If

        'Get dtd path on target
        If objSys.DTDLib.Trim <> "" Then
            dtdLib = objSys.DTDLib
        End If
        If eng.DTDLib.Trim <> "" Then
            dtdLib = eng.DTDLib
        End If

        Select Case str.StructureType
            Case enumStructure.STRUCT_COBOL, enumStructure.STRUCT_COBOL_IMS, enumStructure.STRUCT_IMS
                descLib = cobLib
            Case enumStructure.STRUCT_XMLDTD
                descLib = dtdLib
            Case enumStructure.STRUCT_REL_DDL
                descLib = ddlLib
            Case Else
                'Path is nothing already
        End Select

        If str.fPath1.StartsWith("%") = True Then
            DDstr = str.fPath1
        Else
            Select Case objSys.OSType
                Case "z/OS"
                    If IsDBD = False Then
                        If descLib.Trim <> "" And descLib.Contains("/") = False And descLib.Contains("\") = False Then
                            DDstr = "DD:" & QuoteRes(descLib) & "(" & QuoteRes(GetFileNameWithoutExtenstionFromPath(str.fPath1)) & ")"
                        Else
                            DDstr = "DD:" & QuoteRes(GetFileNameWithoutExtenstionFromPath(str.fPath1))
                        End If
                    Else
                        If descLib.Trim <> "" And descLib.Contains("/") = False And descLib.Contains("\") = False Then
                            DDstr = "DD:" & QuoteRes(descLib) & "(" & QuoteRes(GetFileNameWithoutExtenstionFromPath(str.fPath2)) & ")"
                        Else
                            DDstr = "DD:" & QuoteRes(GetFileNameWithoutExtenstionFromPath(str.fPath2))
                        End If
                    End If
                Case "UNIX"
                    If IsDBD = False Then
                        If descLib.Trim <> "" Then  'And descLib.Contains("/") = True
                            DDstr = descLib & "/" & GetFileNameOnly(str.fPath1)
                        Else
                            DDstr = str.fPath1.Replace("\", "/")
                        End If
                    Else
                        If descLib.Trim <> "" Then    'And descLib.Contains("/") = True
                            DDstr = descLib & "/" & GetFileNameOnly(str.fPath2)
                        Else
                            DDstr = str.fPath2.Replace("\", "/")
                        End If
                    End If
                    If DDstr.Contains("/") = True Then
                        Dim UnixScriptPath As String = ScriptPath.Replace("\", "/")
                        If DDstr.Contains(UnixScriptPath) = True Then
                            DDstr = DDstr.Replace(UnixScriptPath & "/", "")
                        End If
                    End If
                Case "Windows"
                    If IsDBD = False Then
                        If descLib.Trim <> "" And descLib.Contains("\") = True Then
                            DDstr = descLib & "\" & GetFileNameOnly(str.fPath1)
                        Else
                            DDstr = str.fPath1
                        End If
                    Else
                        If descLib.Trim <> "" And descLib.Contains("\") = True Then
                            DDstr = descLib & "\" & GetFileNameOnly(str.fPath2)
                        Else
                            DDstr = str.fPath2
                        End If
                    End If
                    If DDstr.Contains("/") = False Then
                        If DDstr.Contains(ScriptPath) = True Then
                            DDstr = DDstr.Replace(ScriptPath & "\", "")
                        End If
                    End If
            End Select
        End If
        GetStrPath = DDstr

    End Function

    Public Function GetIncPath(ByVal InStr As String, ByVal eng As clsEngine) As String

        Dim DDstr As String = ""
        Dim objSys As clsSystem = eng.ObjSystem
        Dim incLib As String = ""

        If objSys.IncludeLib.Trim <> "" Then
            incLib = objSys.IncludeLib
        End If
        If eng.IncludeLib.Trim <> "" Then
            incLib = eng.IncludeLib
        End If

        Select Case objSys.OSType
            Case "z/OS"
                If incLib.Trim <> "" Then
                    DDstr = "DD:" & QuoteRes(incLib) & "(" & GetFileNameWithoutExtenstionFromPath(InStr) & ")"
                Else
                    DDstr = "DD:" & QuoteRes(GetFileNameWithoutExtenstionFromPath(InStr))
                End If
            Case "UNIX"
                If incLib.Trim <> "" Then
                    DDstr = incLib & "/" & QuoteRes(GetFileNameOnly(InStr))
                Else
                    DDstr = QuoteRes(InStr)
                End If
            Case "Windows"
                If incLib.Trim <> "" Then
                    DDstr = incLib & "\" & QuoteRes(GetFileNameOnly(InStr))
                Else
                    DDstr = QuoteRes(InStr)
                End If
        End Select

        GetIncPath = DDstr

    End Function

    Public Function GetDStype(ByVal DStype As enumDatastore) As String

        Dim retStr As String

        Select Case DStype
            Case enumDatastore.DS_UNKNOWN
                retStr = "UNKNOWN"
            Case enumDatastore.DS_BINARY
                retStr = "BINARY"
            Case enumDatastore.DS_TEXT
                retStr = "TEXT"
            Case enumDatastore.DS_DELIMITED
                retStr = "DELIMITED"
            Case enumDatastore.DS_XML
                retStr = "XMLCDC"
            Case enumDatastore.DS_RELATIONAL
                retStr = "RELATIONAL"
            Case enumDatastore.DS_DB2LOAD
                retStr = "DB2LOAD"
            Case enumDatastore.DS_HSSUNLOAD
                retStr = "HSSUNLOAD"
            Case enumDatastore.DS_IMSDB
                retStr = "IMSDB"
            Case enumDatastore.DS_VSAM
                retStr = "VSAM"
            Case enumDatastore.DS_IMSCDC
                retStr = "IMSCDC"
            Case enumDatastore.DS_IMSCDCLE
                retStr = "IMSLE"
            Case enumDatastore.DS_DB2CDC
                retStr = "DB2CDC"
            Case enumDatastore.DS_VSAMCDC
                retStr = "VSAMCDC"
                'Case enumDatastore.DS_XMLCDC
                '    retStr = "XMLCDC"
            Case enumDatastore.DS_IBMEVENT
                retStr = "IBMEVENT"
                'Case enumDatastore.DS_TRBCDC
                '    retStr = "TRBCDC"
            Case enumDatastore.DS_UTSCDC
                retStr = "UTSCDC"
            Case enumDatastore.DS_ORACLECDC
                retStr = "ORACLECDC"
                'Case enumDatastore.DS_IMSLEBATCH
                '    retStr = "IMSLEBATCH"
            Case enumDatastore.DS_SUBVAR
                retStr = GetVarNum("SubVar")
            Case Else
                retStr = ""
        End Select

        GetDStype = retStr

    End Function

    Public Function GetOperType(ByVal OpType As String) As String

        Dim retStr As String

        Select Case OpType
            Case "I"
                retStr = "INSERT"
            Case "R"
                retStr = "REPLACE"
            Case "U"
                retStr = "UPDATE"
            Case "M"
                retStr = "MODIFY"
            Case "C"
                retStr = "CHANGE"
            Case "D"
                retStr = "DELETE"
            Case Else
                retStr = ""
        End Select

        GetOperType = retStr

    End Function

    Public Function PopulateResWords() As Boolean

        Try
            Dim Word As String = "dummy"
            ResWords.Clear()
            If System.IO.File.Exists(GetAppPath() & "Resword.dfn") = True Then
                ResWordStream = System.IO.File.Open(GetAppPath() & "Resword.dfn", IO.FileMode.Open)
                objReadResWord = New System.IO.StreamReader(ResWordStream)

                While objReadResWord.Peek > -1
                    Word = objReadResWord.ReadLine
                    If ResWords.Contains(Word) = False Then
                        ResWords.Add(Word, Word)
                    End If
                End While
                objReadResWord.Close()
                ResWordStream.Close()
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modGenerateV3 PopulateResWords")
            Return False
        End Try

    End Function

    Public Function QuoteRes(ByVal InStr As String) As String

        Try
            If ResWords.Contains(InStr) = True Then
                InStr = "`" & InStr & "`"
            End If

            Return InStr

        Catch ex As Exception
            LogError(ex, "modGenerateV3 QuoteRes")
            Return InStr
        End Try

    End Function

    Public Function QuoteAttr(ByVal InStr As String) As String

        Try
            Dim StrArry As Char() = InStr.ToCharArray

            For Each chr As Char In StrArry
                If Char.IsLetter(chr) = True Then
                    InStr = "'" & InStr & "'"
                    Exit For
                End If
            Next

            Return InStr

        Catch ex As Exception
            LogError(ex, "modGenerateV3 QuoteRes")
            Return InStr
        End Try

    End Function

#End Region

#Region " Run SQData "

    Function RunSQData(ByVal pathPRC As String, Optional ByVal InEng As clsEngine = Nothing) As String

        '//run out exe with command line args so Runs the Engine in SQData
        Try
            'Dim TempPath As String = GetAppTemp()
            Dim fsERR As System.IO.FileStream
            Dim objWriteERR As System.IO.StreamWriter
            Dim PathErr As String = IO.Path.Combine(GetDirFromPath(pathPRC), "sqdata.ERR")
            'Dim fsOUT As System.IO.FileStream
            'Dim objWriteOUT As System.IO.StreamWriter
            'Dim PathOUT As String = IO.Path.Combine(GetDirFromPath(pathPRC), GetFileNameWithoutExtenstionFromPath(pathPRC) & ".OUT")

            Dim FORargs As String = String.Format("{0}", Quote(GetFileNameFromPath(pathPRC), """"))
            Dim args As String = FORargs
            Dim si As New ProcessStartInfo()

            Log("********* sqdata engine start *********")
            Log("*** SQData Engine Started : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

            '//delete previous log file
            If System.IO.File.Exists(PathErr) Then
                System.IO.File.Delete(PathErr)
            End If
            '// create new error log stream
            fsERR = System.IO.File.Create(PathErr)
            PathErr = fsERR.Name
            objWriteERR = New System.IO.StreamWriter(fsERR)

            ''//delete previous output file
            'If System.IO.File.Exists(PathOUT) Then
            '    System.IO.File.Delete(PathOUT)
            'End If
            ''// create new output file stream
            'fsOUT = System.IO.File.Create(PathOUT)
            'PathOUT = fsOUT.Name
            'objWriteOUT = New System.IO.StreamWriter(fsOUT)  

            If InEng IsNot Nothing Then
                If InEng.EngVersion <> "" Then
                    EnginePath = GetAppPath() & "sqdata.exe"  'InEng.EngVersion & "\" & 
                Else
                    EnginePath = GetAppPath() & "sqdata.exe"
                End If
            End If

            si.WorkingDirectory = GetDirFromPath(pathPRC)
            si.FileName = EnginePath
            si.Arguments = args
            si.UseShellExecute = False
            si.CreateNoWindow = True

            '// Redirect Standard Error. Let standard Output go  
            si.RedirectStandardOutput = False
            si.RedirectStandardError = True

            '// Create a new process to Run SQData
            Using myProcess As New System.Diagnostics.Process()
                myProcess.StartInfo = si

                Log("*** " & si.FileName & args)

                myProcess.Start()

                Dim OutStr As String = ""
                Dim ErrStr As String = ""

                '/// split output into multiple threads to capture each stream to a string
                '/// OutStr stays as "" because StdOut is NOT redirected
                OutputToEnd(myProcess, OutStr, ErrStr)

                '//wait until task is done
                myProcess.WaitForExit()

                objWriteERR.Write(ErrStr)
                objWriteERR.Close()
                fsERR.Close()

                'objWriteOUT.Write(OutStr) 
                'objWriteOUT.Close()
                'fsOUT.Close()

                If myProcess.ExitCode <> 0 Then
                    If MsgBox("Error occurred while Engine [" & _
                              GetFileNameWithoutExtenstionFromPath(pathPRC) & "] Executed." & _
                              vbCrLf & "Do you want to see the log?", _
                              MsgBoxStyle.Critical Or MsgBoxStyle.YesNoCancel, _
                              MsgTitle) = MsgBoxResult.Yes Then
                        If IO.File.Exists(PathErr) Then
                            Process.Start(PathErr)
                        End If
                    End If
                    RunSQData = ""
                Else
                    RunSQData = PathErr
                End If

                Log("*** SQData Engine Return Code = " & myProcess.ExitCode & " *********")
                Log("*** SQData Engine Report file saved at : " & PathErr)
                Log("*** SQData Engine Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

                myProcess.Close()

            End Using

        Catch ex As Exception
            LogError(ex, "modGeneral RunSQData")
            Return ""
        End Try

    End Function

#End Region

End Module