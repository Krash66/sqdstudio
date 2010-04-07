Module modRename
    'This module is intended to enable renaming of all objects with Names stored in Metadata
    'Added January 2007 by Tom Karasch  
    'Modified May 2007 For Structure File Replacement
    'Modified July 2007 for Moving of connections to Environment
    'Modified August 2008
    'Modified Jan thru May 2009 for V3 and new Object Model changes

#Region "Main Rename Functions"

    Function RenameProject(ByRef Proj As clsProject, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameProject = False
            Exit Function
        End If
        If Proj.ValidateNewObject(NewName) = False Then
            RenameProject = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(Proj.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran
                If Proj.ProjectMetaVersion = enumMetaVersion.V2 Then
                    success = EditProjects(cmd, Proj.ProjectName, NewName, Proj)
                    If success = True Then
                        success = EditConnections(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDatastores(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDSSELFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDSSELECTIONS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditEngines(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditEnvironments(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditStructures(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditSTRUCTFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditSTRSELFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditStructureSelections(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditSystems(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTaskDatastores(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTaskMappings(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTasks(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditVariables(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                Else
                    success = EditProjects(cmd, Proj.ProjectName, NewName, Proj)
                    If success = True Then
                        success = EditConnections(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDatastores(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDSSELFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDSSELECTIONS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditEngines(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditEnvironments(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDescriptions(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDESCFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDESCSELFIELDS(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditDESCSelections(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditSystems(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTaskDatastores(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTaskMappings(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditTasks(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                    If success = True Then
                        success = EditVariables(cmd, Proj.ProjectName, NewName, Proj)
                    End If
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameProject = True
                Else
                    tran.Rollback()
                    RenameProject = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameProject")
                RenameProject = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

    Function RenameEnvironment(ByRef Env As clsEnvironment, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameEnvironment = False
            Exit Function
        End If
        If Env.ValidateNewObject(NewName) = False Then
            RenameEnvironment = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(Env.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditConnections(cmd, Env.EnvironmentName, NewName, Env)
                If success = True Then
                    success = EditDatastores(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditDSSELFIELDS(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditDSSELECTIONS(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditEngines(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditEnvironments(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If Env.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    If success = True Then
                        success = EditStructures(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditSTRUCTFIELDS(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditSTRSELFIELDS(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditStructureSelections(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                Else
                    If success = True Then
                        success = EditDescriptions(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditDESCFIELDS(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditDESCSELFIELDS(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                    If success = True Then
                        success = EditDESCSelections(cmd, Env.EnvironmentName, NewName, Env)
                    End If
                End If
                If success = True Then
                    success = EditSystems(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditTaskDatastores(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditTasks(cmd, Env.EnvironmentName, NewName, Env)
                End If
                If success = True Then
                    success = EditVariables(cmd, Env.EnvironmentName, NewName, Env)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameEnvironment = True
                Else
                    tran.Rollback()
                    RenameEnvironment = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameEnvironment")
                RenameEnvironment = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

    Function RenameSystem(ByRef sys As clsSystem, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameSystem = False
            Exit Function
        End If
        If sys.ValidateNewObject(NewName) = False Then
            RenameSystem = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(sys.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                'Call EditConnections(cmd, sys.SystemName, NewName, sys)
                success = EditDatastores(cmd, sys.SystemName, NewName, sys)
                If success = True Then
                    success = EditDSSELFIELDS(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditDSSELECTIONS(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditEngines(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditSystems(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditTaskDatastores(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditTasks(cmd, sys.SystemName, NewName, sys)
                End If
                If success = True Then
                    success = EditVariables(cmd, sys.SystemName, NewName, sys)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameSystem = True
                Else
                    tran.Rollback()
                    RenameSystem = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameSystem")
                RenameSystem = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

    Function RenameEngine(ByRef eng As clsEngine, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName, True) = False Then
            RenameEngine = False
            Exit Function
        End If
        If eng.ValidateNewObject(NewName) = False Then
            RenameEngine = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(eng.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditDatastores(cmd, eng.EngineName, NewName, eng)
                If success = True Then
                    success = EditDSSELFIELDS(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditDSSELECTIONS(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditEngines(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditTaskDatastores(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditTasks(cmd, eng.EngineName, NewName, eng)
                End If
                If success = True Then
                    success = EditVariables(cmd, eng.EngineName, NewName, eng)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameEngine = True
                Else
                    tran.Rollback()
                    RenameEngine = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameEngine")
                RenameEngine = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

    Function RenameStructure(ByRef struct As clsStructure, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameStructure = False
            Exit Function
        End If
        If struct.ValidateNewObject(NewName) = False Then
            RenameStructure = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(struct.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran


                success = EditDSSELFIELDS(cmd, struct.StructureName, NewName, struct)
                If success = True Then
                    success = EditDSSELECTIONS(cmd, struct.StructureName, NewName, struct)
                End If
                If struct.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    If success = True Then
                        success = EditStructures(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditSTRUCTFIELDS(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditSTRSELFIELDS(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditStructureSelections(cmd, struct.StructureName, NewName, struct)
                    End If
                Else
                    If success = True Then
                        success = EditDescriptions(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditDESCFIELDS(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditDESCSELFIELDS(cmd, struct.StructureName, NewName, struct)
                    End If
                    If success = True Then
                        success = EditDESCSelections(cmd, struct.StructureName, NewName, struct)
                    End If
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, struct.StructureName, NewName, struct)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameStructure = True
                Else
                    tran.Rollback()
                    RenameStructure = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameStructure")
                RenameStructure = False
            Finally
                'cnn.Close()
                struct.LoadItems(True)
            End Try
        End If

    End Function

    Function RenameStructureSelection(ByRef StrSel As clsStructureSelection, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameStructureSelection = False
            Exit Function
        End If
        If StrSel.ValidateNewObject(NewName) = False Then
            RenameStructureSelection = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(StrSel.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditDSSELFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                If success = True Then
                    success = EditDSSELECTIONS(cmd, StrSel.SelectionName, NewName, StrSel)
                End If
                If StrSel.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    If success = True Then
                        success = EditSTRUCTFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                    If success = True Then
                        success = EditSTRSELFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                    If success = True Then
                        success = EditStructureSelections(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                Else
                    If success = True Then
                        success = EditDESCFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                    If success = True Then
                        success = EditDESCSELFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                    If success = True Then
                        success = EditDESCSelections(cmd, StrSel.SelectionName, NewName, StrSel)
                    End If
                End If
                If success = True Then
                    success = EditTaskDatastores(cmd, StrSel.SelectionName, NewName, StrSel)
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, StrSel.SelectionName, NewName, StrSel)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameStructureSelection = True
                Else
                    tran.Rollback()
                    RenameStructureSelection = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameStructureSelection")
                RenameStructureSelection = False
            Finally
                'cnn.Close()
                StrSel.LoadItems()
            End Try
        End If

    End Function

    Function RenameConnection(ByRef Conn As clsConnection, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameConnection = False
            Exit Function
        End If
        If Conn.ValidateNewObject(NewName) = False Then
            RenameConnection = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(Conn.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditConnections(cmd, Conn.ConnectionName, NewName, Conn)
                If success = True Then
                    If Conn.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        success = EditStructures(cmd, Conn.ConnectionName, NewName, Conn)
                    Else
                        success = EditDescriptions(cmd, Conn.ConnectionName, NewName, Conn)
                    End If
                End If
                If success = True Then
                    success = EditEngines(cmd, Conn.ConnectionName, NewName, Conn)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameConnection = True
                Else
                    tran.Rollback()
                    RenameConnection = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameConnection")
                RenameConnection = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

    Function RenameDatastore(ByRef DS As clsDatastore, ByVal NewName As String) As Boolean

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim success As Boolean = False

        If ValidateNewName(NewName) = False Then
            RenameDatastore = False
            Exit Function
        End If
        If DS.ValidateNewObject(NewName) = False Then
            RenameDatastore = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(DS.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditDatastores(cmd, DS.DatastoreName, NewName, DS)
                If success = True Then
                    success = EditDSSELECTIONS(cmd, DS.DatastoreName, NewName, DS)
                End If
                If success = True Then
                    success = EditDSSELFIELDS(cmd, DS.DatastoreName, NewName, DS)
                End If
                If success = True Then
                    success = EditTaskDatastores(cmd, DS.DatastoreName, NewName, DS)
                End If
                If success = True Then
                    success = EditTaskMappings(cmd, DS.DatastoreName, NewName, DS)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameDatastore = True
                Else
                    tran.Rollback()
                    RenameDatastore = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameDatastore")
                RenameDatastore = False
            Finally
                'cnn.Close()
                DS.LoadItems()
            End Try
        End If

    End Function

    Function RenameTask(ByRef Task As clsTask, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameTask = False
            Exit Function
        End If
        If Task.ValidateNewObject(NewName) = False Then
            RenameTask = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(Task.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditTaskDatastores(cmd, Task.TaskName, NewName, Task)
                If success = True Then
                    success = EditTaskMappings(cmd, Task.TaskName, NewName, Task)
                End If
                If success = True Then
                    success = EditTasks(cmd, Task.TaskName, NewName, Task)
                End If
                'If success = True Then
                '    success = EditVariables(cmd, Task.TaskName, NewName, Task)
                'End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameTask = True
                Else
                    tran.Rollback()
                    RenameTask = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameTask")
                RenameTask = False
            Finally
                'cnn.Close()
                Task.LoadDatastores()
                Task.LoadMappings()
            End Try
        End If

    End Function

    Function RenameVariable(ByRef Var As clsVariable, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        If ValidateNewName(NewName) = False Then
            RenameVariable = False
            Exit Function
        End If
        If Var.ValidateNewObject(NewName) = False Then
            RenameVariable = False
            Exit Function
        Else
            Try
                'cnn = New Odbc.OdbcConnection(Var.Project.MetaConnectionString)
                'cnn.Open()
                cmd.Connection = cnn
                tran = cnn.BeginTransaction()
                cmd.Transaction = tran

                success = EditVariables(cmd, Var.VariableName, NewName, Var)

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameVariable = True
                Else
                    tran.Rollback()
                    RenameVariable = False
                End If

            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameVariable")
                RenameVariable = False
            Finally
                'cnn.Close()
            End Try
        End If

    End Function

#End Region

#Region "Metadata Table Updates"

    'done
    Function EditConnections(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Try
            Dim objProj As clsProject
            Dim objEnv As clsEnvironment
            Dim objconn As clsConnection

            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblConnections & " Set ProjectName= " & Quote(NewValue) & _
                    " WHERE ProjectName= " & Quote(OldValue)

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblConnections & " Set EnvironmentName= " & Quote(NewValue) & _
                    " WHERE ProjectName=" & Quote(objEnv.Project.ProjectName) & _
                    " AND EnvironmentName=" & Quote(OldValue)
                   
                Case NODE_CONNECTION
                    objconn = CType(obj, clsConnection)
                    sql = "Update " & objconn.Project.tblConnections & " Set ConnectionName= " & Quote(NewValue) & _
                    " WHERE  ProjectName=" & Quote(objconn.Project.ProjectName) & _
                    " AND EnvironmentName=" & Quote(objconn.Environment.EnvironmentName) & _
                    " AND ConnectionName=" & Quote(OldValue)

                Case Else
                    EditConnections = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditConnections = True

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                EditConnections = EditConnectionATTR(cmd, OldValue, NewValue, obj)
            End If

        Catch ex As Exception
            LogError(ex, "modRename EditConnections", sql)
            EditConnections = False
        End Try

    End Function

    'done
    Function EditConnectionATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Try
            Dim objProj As clsProject
            Dim objEnv As clsEnvironment
            Dim objconn As clsConnection

            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblConnectionsATTR & " Set ProjectName= " & Quote(NewValue) & _
                    " WHERE ProjectName= " & Quote(OldValue)

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblConnectionsATTR & " Set EnvironmentName= " & Quote(NewValue) & _
                    " WHERE ProjectName=" & Quote(objEnv.Project.ProjectName) & _
                    " AND EnvironmentName=" & Quote(OldValue)

                Case NODE_CONNECTION
                    objconn = CType(obj, clsConnection)
                    sql = "Update " & objconn.Project.tblConnectionsATTR & " Set ConnectionName= " & Quote(NewValue) & _
                    " WHERE  ProjectName=" & Quote(objconn.Project.ProjectName) & _
                    " AND EnvironmentName=" & Quote(objconn.Environment.EnvironmentName) & _
                    " AND ConnectionName=" & Quote(OldValue)

                Case Else
                    EditConnectionATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditConnectionATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditConnectionsATTR", sql)
            EditConnectionATTR = False
        End Try

    End Function

    'done
    Function EditDatastores(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objDS As clsDatastore

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDatastores & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "';"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDatastores & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblDatastores & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblDatastores & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    objDS = CType(obj, clsDatastore)
                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblDatastores & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblDatastores & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If
                   

                Case Else
                    EditDatastores = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Call EditDatastoresATTR(cmd, OldValue, NewValue, obj)

            EditDatastores = True

        Catch ex As Exception
            LogError(ex, "modRename EditDatastores", sql)
            EditDatastores = False
        End Try

    End Function

    'done
    Function EditDatastoresATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objDS As clsDatastore

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDatastoresATTR & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "';"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDatastoresATTR & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblDatastoresATTR & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblDatastoresATTR & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE  ', node_datastore
                    objDS = CType(obj, clsDatastore)
                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblDatastoresATTR & " Set DatastoreName= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblDatastoresATTR & _
                        " Set DatastoreName= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If
                    

                Case Else
                    EditDatastoresATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDatastoresATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditDatastores", sql)
            EditDatastoresATTR = False
        End Try

    End Function

    'done
    Function EditDSSELFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objDS As clsDatastore
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDSselFields & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDSselFields & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblDSselFields & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblDSselFields & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    objDS = CType(obj, clsDatastore)
                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblDSselFields & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblDSselFields & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If
                   

                    If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        If objDS.Engine IsNot Nothing Then
                            sql = "Update " & objDS.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                                                    "' WHERE Parent= '" & OldValue & _
                                                    "' AND DatastoreName='" & NewValue & _
                                                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                                                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                                    "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                                    "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                        Else
                            sql = "Update " & objDS.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                                                    "' WHERE Parent= '" & OldValue & _
                                                    "' AND DatastoreName='" & NewValue & _
                                                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                                                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                        End If
                    Else
                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        If objDS.Engine IsNot Nothing Then
                            sql = "Update " & objDS.Project.tblDSselFields & " Set ParentName= '" & NewValue & _
                                                    "' WHERE ParentName= '" & OldValue & _
                                                    "' AND DatastoreName='" & NewValue & _
                                                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                                                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                                    "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                                    "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                        Else
                            sql = "Update " & objDS.Project.tblDSselFields & " Set ParentName= '" & NewValue & _
                                                    "' WHERE ParentName= '" & OldValue & _
                                                    "' AND DatastoreName='" & NewValue & _
                                                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                                                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                        End If
                    End If

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    If objStr.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "Update " & objStr.Project.tblDSselFields & _
                        " Set StructureName= '" & NewValue & _
                        "',SelectionName= '" & NewValue & _
                        "' WHERE StructureName= '" & OldValue & _
                        "' AND SelectionName='" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselFields & " Set StructureName= '" & NewValue & _
                        "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                        "' WHERE Parent= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"
                    Else
                        sql = "Update " & objStr.Project.tblDSselFields & _
                        " Set DescriptionName= '" & NewValue & _
                        "',SelectionName= '" & NewValue & _
                        "' WHERE DescriptionName= '" & OldValue & _
                        "' AND SelectionName='" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselFields & _
                        " Set DescriptionName= '" & NewValue & _
                        "' WHERE DescriptionName= '" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselFields & _
                        " Set ParentName= '" & NewValue & _
                        "' WHERE ParentName= '" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"
                    End If
                   

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    If objSel.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "Update " & objSel.Project.tblDSselFields & " Set SelectionName= '" & NewValue & _
                        "' WHERE SelectionName= '" & OldValue & "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                        "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objSel.Project.tblDSselFields & _
                        " Set Parent= '" & NewValue & _
                        "' WHERE Parent= '" & OldValue & _
                        "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"
                    Else
                        sql = "Update " & objSel.Project.tblDSselFields & _
                        " Set SelectionName= '" & NewValue & _
                        "' WHERE SelectionName= '" & OldValue & _
                        "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                        "' AND DescriptionName= '" & objSel.ObjStructure.StructureName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objSel.Project.tblDSselFields & _
                        " Set ParentName= '" & NewValue & _
                        "' WHERE ParentName= '" & OldValue & _
                        "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"
                    End If


                Case Else
                    EditDSSELFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Call EditFkey(cmd, OldValue, NewValue, obj)

            
            EditDSSELFIELDS = True

        Catch ex As Exception
            LogError(ex, "modRename EditDSselFields", sql)
            EditDSSELFIELDS = False
        End Try

    End Function

    'done
    Function EditDSSELECTIONS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objDS As clsDatastore
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDSselections & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDSselections & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblDSselections & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblDSselections & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    objDS = CType(obj, clsDatastore)

                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblDSselections & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblDSselections & " Set DatastoreName= '" & NewValue & _
                                           "' WHERE DatastoreName= '" & OldValue & _
                                           "' AND ProjectName='" & objDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If
                   
                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblDSselections & _
                                            " Set Parent= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & NewValue & _
                                            "' AND Parent= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblDSselections & _
                                            " Set Parent= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & NewValue & _
                                            "' AND Parent= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If
                    

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    If objStr.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "Update " & objStr.Project.tblDSselections & " Set StructureName= '" & NewValue & _
                                            "',  SelectionName= '" & NewValue & "' WHERE StructureName= '" & OldValue & _
                                            "' AND SelectionName='" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselections & " Set StructureName= '" & NewValue & _
                        "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    Else
                        sql = "Update " & objStr.Project.tblDSselections & _
                        " Set DescriptionName= '" & NewValue & _
                        "', SelectionName= '" & NewValue & _
                        "' WHERE DescriptionName= '" & OldValue & _
                        "' AND SelectionName='" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                        cmd.CommandText = sql
                        Log(sql)
                        cmd.ExecuteNonQuery()

                        sql = "Update " & objStr.Project.tblDSselections & _
                        " Set DescriptionName= '" & NewValue & _
                        "' WHERE DescriptionName= '" & OldValue & _
                        "' AND ProjectName='" & objStr.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    End If
                    
                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objStr.Project.tblDSselections & " Set Parent= '" & NewValue & _
                    "' WHERE Parent= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    If objSel.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "Update " & objSel.Project.tblDSselections & _
                        " Set SelectionName= '" & NewValue & _
                        "' WHERE SelectionName= '" & OldValue & _
                        "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                        "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"
                    Else
                        sql = "Update " & objSel.Project.tblDSselections & _
                        " Set SelectionName= '" & NewValue & _
                        "' WHERE SelectionName= '" & OldValue & _
                        "' AND ProjectName='" & objSel.Project.ProjectName & _
                        "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                        "' AND DescriptionName= '" & objSel.ObjStructure.StructureName & "'"
                    End If

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objSel.Project.tblDSselections & " Set Parent= '" & NewValue & _
                    "' WHERE Parent= '" & OldValue & "' AND ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"

                Case Else
                    EditDSSELECTIONS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDSSELECTIONS = True

        Catch ex As Exception
            LogError(ex, "modRename EditDSselections", sql)
            EditDSSELECTIONS = False
        End Try

    End Function

    'done
    Function EditEngines(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim ObjConn As clsConnection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblEngines & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblEngines & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblEngines & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblEngines & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_CONNECTION
                    ObjConn = CType(obj, clsConnection)
                    If ObjConn.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        sql = "Update " & ObjConn.Project.tblEngines & _
                        " Set ConnectionName= '" & NewValue & _
                        "' WHERE ConnectionName= '" & OldValue & _
                        "' AND ProjectName='" & ObjConn.Project.ProjectName & _
                        "' AND EnvironmentName= '" & ObjConn.Environment.EnvironmentName & "'"
                    Else
                        GoTo goto1
                    End If
                    
                Case Else
                    EditEngines = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

goto1:      If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                Call EditEnginesATTR(cmd, OldValue, NewValue, obj)
            End If

            EditEngines = True

        Catch ex As Exception
            LogError(ex, "modRename EditEngines", sql)
            EditEngines = False
        End Try

    End Function

    'done
    Function EditEnginesATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim ObjConn As clsConnection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblEnginesATTR & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblEnginesATTR & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblEnginesATTR & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblEnginesATTR & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_CONNECTION
                    ObjConn = CType(obj, clsConnection)

                    sql = "Update " & ObjConn.Project.tblEnginesATTR & " Set EngineAttrbValue= " & Quote(NewValue) & _
                    " WHERE ProjectName=" & Quote(ObjConn.Project.ProjectName) & _
                    " AND EnvironmentName= " & Quote(ObjConn.Environment.EnvironmentName) & _
                    " AND EngineAttrb='CONNECTIONNAME'" & _
                    " AND EngineAttrbValue=" & Quote(OldValue)


                Case Else
                    EditEnginesATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditEnginesATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditEnginesATTR", sql)
            EditEnginesATTR = False
        End Try

    End Function

    'done
    Function EditEnvironments(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblEnvironments & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblEnvironments & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case Else
                    EditEnvironments = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditEnvironments = True

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                EditEnvironments = EditEnvironmentsATTR(cmd, OldValue, NewValue, obj)
            End If

        Catch ex As Exception
            LogError(ex, "modRename EditEnvironments", sql)
            EditEnvironments = False
        End Try

    End Function

    'done
    Function EditEnvironmentsATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblEnvironmentsATTR & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblEnvironmentsATTR & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case Else
                    EditEnvironmentsATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditEnvironmentsATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditEnvironmentsATTR", sql)
            EditEnvironmentsATTR = False
        End Try

    End Function

    'done
    Function EditProjects(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean
        Dim sql As String = ""

        Try
            sql = "Update " & obj.Project.tblProjects & _
            " Set ProjectName= '" & NewValue & _
            "' WHERE ProjectName= '" & OldValue & "'"

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditProjects = True

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                EditProjects = EditProjectsATTR(cmd, OldValue, NewValue, obj)
            End If

        Catch ex As Exception
            LogError(ex, "modRename EditProjects", sql)
            EditProjects = False
        End Try

    End Function

    'done
    Function EditProjectsATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean
        Dim sql As String = ""

        Try
            sql = "Update " & obj.Project.tblProjectsATTR & _
            " Set ProjectName= '" & NewValue & _
            "' WHERE ProjectName= '" & OldValue & "'"

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditProjectsATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditProjects", sql)
            EditProjectsATTR = False
        End Try

    End Function

    'done
    Function EditSTRUCTFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblStructFields & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblStructFields & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    '/// sets structure name and parent name
                    objStr = CType(obj, clsStructure)

                    sql = "Update " & objStr.Project.tblStructFields & _
                    " Set StructureName= '" & NewValue & _
                    "',ParentName= '" & NewValue & _
                    "' WHERE StructureName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case Else
                    EditSTRUCTFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditSTRUCTFIELDS = True

        Catch ex As Exception
            LogError(ex, "modRename EditStructFields", sql)
            EditSTRUCTFIELDS = False
        End Try

    End Function

    'done
    Function EditStructures(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim ObjConn As clsConnection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblStructures & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblStructures & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblStructures & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_CONNECTION
                    objConn = CType(obj, clsConnection)
                    sql = "Update " & ObjConn.Project.tblStructures & " Set ConnectionName= '" & NewValue & "' WHERE ConnectionName= '" & OldValue & "' AND ProjectName='" & ObjConn.Project.ProjectName & "' AND EnvironmentName= '" & ObjConn.Environment.EnvironmentName & "'"

                Case Else
                    EditStructures = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditStructures = True

        Catch ex As Exception
            LogError(ex, "modRename EditStructures", sql)
            EditStructures = False
        End Try

    End Function

    'done
    Function EditSTRSELFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblStrSelFields & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblStrSelFields & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblStrSelFields & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "Update " & objSel.Project.tblStrSelFields & " Set SelectionName= '" & NewValue & "' WHERE SelectionName= '" & OldValue & "' AND ProjectName='" & objSel.Project.ProjectName & "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

                Case Else
                    EditSTRSELFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditSTRSELFIELDS = True

        Catch ex As Exception
            LogError(ex, "modRename EditStrSelFields", sql)
            EditSTRSELFIELDS = False
        End Try

    End Function

    'done
    Function EditStructureSelections(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblStructSel & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblStructSel & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblStructSel & " Set StructureName= '" & NewValue & "',SelectionName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "' AND ISSYSTEMSELECT= '1'"

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objStr.Project.tblStructSel & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "' AND ISSYSTEMSELECT= '0'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "Update " & objSel.Project.tblStructSel & " Set SelectionName= '" & NewValue & "' WHERE SelectionName= '" & OldValue & "' AND ProjectName= '" & objSel.Project.ProjectName & "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

                Case Else
                    EditStructureSelections = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditStructureSelections = True

        Catch ex As Exception
            LogError(ex, "modRename EditStructureSelections", sql)
            EditStructureSelections = False
        End Try

    End Function

    'done
    Function EditDESCFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDescriptionFields & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDescriptionFields & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    '/// sets structure name and parent name
                    objStr = CType(obj, clsStructure)
                    
                    sql = "Update " & objStr.Project.tblDescriptionFields & _
                    " set descriptionname=" & Quote(NewValue) & _
                    ", ParentName=" & Quote(NewValue) & _
                    " where ProjectName= " & Quote(objStr.Project.ProjectName) & _
                    " and EnvironmentName= " & Quote(objStr.Environment.EnvironmentName) & _
                    " and DescriptionName= " & Quote(OldValue)

                Case Else
                    EditDESCFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDESCFIELDS = True

        Catch ex As Exception
            LogError(ex, "modRename EditDescFields", sql)
            EditDESCFIELDS = False
        End Try

    End Function

    'done
    Function EditDescriptions(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
       
        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDescriptions & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDescriptions & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblDescriptions & _
                    " Set DescriptionName= '" & NewValue & _
                    "' WHERE DescriptionName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_CONNECTION
                    GoTo editGoTo

                Case Else
                    EditDescriptions = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

editGoTo:   EditDescriptions = EditDescriptionsATTR(cmd, OldValue, NewValue, obj)

        Catch ex As Exception
            LogError(ex, "modRename EditDescriptions", sql)
            EditDescriptions = False
        End Try

    End Function

    'done
    Function EditDescriptionsATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim ObjConn As clsConnection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDescriptionsATTR & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDescriptionsATTR & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblDescriptionsATTR & _
                    " Set DescriptionName= '" & NewValue & _
                    "' WHERE DescriptionName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_CONNECTION
                    ObjConn = CType(obj, clsConnection)
                    sql = "Update " & ObjConn.Project.tblDescriptionsATTR & _
                    " Set DESCRIPTIONATTRBVALUE= '" & NewValue & _
                    "' WHERE DESCRIPTIONATTRB='CONNECTIONNAME'" & _
                    " AND DESCRIPTIONATTRBVALUE= '" & OldValue & _
                    "' AND ProjectName='" & ObjConn.Project.ProjectName & _
                    "' AND EnvironmentName= '" & ObjConn.Environment.EnvironmentName & "'"

                Case Else
                    EditDescriptionsATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDescriptionsATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditDescriptionsATTR", sql)
            EditDescriptionsATTR = False
        End Try

    End Function

    'done
    Function EditDESCSELFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDescriptionSelFields & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDescriptionSelFields & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblDescriptionSelFields & _
                    " Set DescriptionName= '" & NewValue & _
                    "' WHERE DescriptionName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "Update " & objSel.Project.tblDescriptionSelFields & _
                    " Set SelectionName= '" & NewValue & _
                    "' WHERE SelectionName= '" & OldValue & _
                    "' AND ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                    "' AND DescriptionName= '" & objSel.ObjStructure.StructureName & "'"

                Case Else
                    EditDESCSELFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDESCSELFIELDS = True

        Catch ex As Exception
            LogError(ex, "modRename EditDESCSELFIELDS", sql)
            EditDESCSELFIELDS = False
        End Try

    End Function

    'done
    Function EditDESCSelections(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblDescriptionSelect & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblDescriptionSelect & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Update " & objStr.Project.tblDescriptionSelect & _
                    " Set DescriptionName= '" & NewValue & _
                    "',SelectionName= '" & NewValue & _
                    "' WHERE DescriptionName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & _
                    "' AND ISSYSTEMSEL= 1"

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objStr.Project.tblDescriptionSelect & _
                    " Set DescriptionName= '" & NewValue & _
                    "' WHERE DescriptionName= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & _
                    "' AND ISSYSTEMSEL= 0"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "Update " & objSel.Project.tblStructSel & _
                    " Set SelectionName= '" & NewValue & _
                    "' WHERE SelectionName= '" & OldValue & _
                    "' AND ProjectName= '" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                    "' AND DescriptionName= '" & objSel.ObjStructure.StructureName & "'"

                Case Else
                    EditDESCSelections = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDESCSelections = True

        Catch ex As Exception
            LogError(ex, "modRename EditDESCSelections", sql)
            EditDESCSelections = False
        End Try

    End Function

    'done
    Function EditSystems(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblSystems & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblSystems & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblSystems & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case Else
                    EditSystems = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditSystems = True

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                EditSystems = EditSystemsATTR(cmd, OldValue, NewValue, obj)
            End If

        Catch ex As Exception
            LogError(ex, "modRename EditSystems", sql)
            EditSystems = False
        End Try

    End Function

    'done
    Function EditSystemsATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblSystemsATTR & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblSystemsATTR & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblSystemsATTR & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case Else
                    EditSystemsATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditSystemsATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditSystemsATTR", sql)
            EditSystemsATTR = False
        End Try

    End Function

    'done
    Function EditTaskDatastores(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objDS As clsDatastore
        Dim objTask As clsTask

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblTaskDS & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblTaskDS & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblTaskDS & _
                    " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblTaskDS & _
                    " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & _
                    "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    objDS = CType(obj, clsDatastore)
                    If objDS.Engine IsNot Nothing Then
                        sql = "Update " & objDS.Project.tblTaskDS & _
                                            " Set DatastoreName= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objDS.Project.tblTaskDS & _
                                            " Set DatastoreName= '" & NewValue & _
                                            "' WHERE DatastoreName= '" & OldValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    End If

                Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                    objTask = CType(obj, clsTask)
                    If objTask.Engine IsNot Nothing Then
                        sql = "Update " & objTask.Project.tblTaskDS & _
                                            " Set TaskName= '" & NewValue & _
                                            "' WHERE TaskName= '" & OldValue & _
                                            "' AND ProjectName='" & objTask.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objTask.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objTask.Project.tblTaskDS & _
                                            " Set TaskName= '" & NewValue & _
                                            "' WHERE TaskName= '" & OldValue & _
                                            "' AND ProjectName='" & objTask.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    End If

                Case Else
                    EditTaskDatastores = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTaskDatastores = True

        Catch ex As Exception
            LogError(ex, "modRename EditTaskDatastores", sql)
            EditTaskDatastores = False
        End Try

    End Function

    'done
    Function EditTaskMappings(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objSDS As clsDatastore
        Dim objTDS As clsDatastore
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection
        Dim objTask As clsTask
        Dim objVar As clsVariable

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblTaskMap & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblTaskMap & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblTaskMap & _
                    " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblTaskMap & _
                    " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & _
                    "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                    objTask = CType(obj, clsTask)

                    Call EditMappingSrc(cmd, OldValue, NewValue, objTask)

                    If objTask.Engine IsNot Nothing Then
                        sql = "Update " & objTask.Project.tblTaskMap & _
                                           " Set TaskName= '" & NewValue & _
                                           "' WHERE TaskName= '" & OldValue & _
                                           "' AND ProjectName='" & objTask.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objTask.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objTask.Project.tblTaskMap & _
                                           " Set TaskName= '" & NewValue & _
                                           "' WHERE TaskName= '" & OldValue & _
                                           "' AND ProjectName='" & objTask.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    End If

                Case NODE_SOURCEDATASTORE
                    objSDS = CType(obj, clsDatastore)

                    Call EditMappingSrc(cmd, OldValue, NewValue, objSDS)

                    If objSDS.Engine IsNot Nothing Then
                        sql = "Update " & objSDS.Project.tblTaskMap & _
                                            " Set SourceDatastore= '" & NewValue & _
                                            "' WHERE SourceDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objSDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objSDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objSDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objSDS.Project.tblTaskMap & _
                                            " Set SourceDatastore= '" & NewValue & _
                                            "' WHERE SourceDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objSDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & "'"
                    End If

                Case NODE_TARGETDATASTORE
                    objTDS = CType(obj, clsDatastore)

                    Call EditMappingSrc(cmd, OldValue, NewValue, objTDS)

                    If objTDS.Engine IsNot Nothing Then
                        sql = "Update " & objTDS.Project.tblTaskMap & _
                                            " Set TargetDatastore= '" & NewValue & _
                                            "' WHERE TargetDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objTDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objTDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objTDS.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objTDS.Project.tblTaskMap & _
                                            " Set TargetDatastore= '" & NewValue & _
                                            "' WHERE TargetDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objTDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & "'"
                    End If

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    Call EditMappingSrc(cmd, OldValue, NewValue, objStr)

                    sql = "Update " & objStr.Project.tblTaskMap & _
                    " Set SourceParent= '" & NewValue & _
                    "' WHERE SourceParent= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objStr.Project.tblTaskMap & _
                    " Set TargetParent= '" & NewValue & _
                    "' WHERE TargetParent= '" & OldValue & _
                    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    EditMappingSrc(cmd, OldValue, NewValue, objSel)

                    EditTaskMappings = True
                    Exit Try

                Case NODE_VARIABLE
                    objVar = CType(obj, clsVariable)

                    EditMappingSrc(cmd, OldValue, NewValue, objVar)

                    EditTaskMappings = True
                    Exit Try

                Case Else
                    EditTaskMappings = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTaskMappings = True

        Catch ex As Exception
            LogError(ex, "modRename EditTaskMappings", sql)
            EditTaskMappings = False
        End Try

    End Function

    'done
    Function EditTasks(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objTask As clsTask

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblTasks & _
                    " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblTasks & _
                    " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblTasks & _
                    " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblTasks & _
                    " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & _
                    "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                    objTask = CType(obj, clsTask)
                    If objTask.Engine IsNot Nothing Then
                        sql = "Update " & objTask.Project.tblTasks & _
                                            " Set TaskName= '" & NewValue & _
                                            "' WHERE TaskName= '" & OldValue & _
                                            "' AND ProjectName='" & objTask.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objTask.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objTask.Project.tblTasks & _
                                            " Set TaskName= '" & NewValue & _
                                            "' WHERE TaskName= '" & OldValue & _
                                            "' AND ProjectName='" & objTask.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    End If

                Case Else
                    EditTasks = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTasks = True

        Catch ex As Exception
            LogError(ex, "modRename EditTasks", sql)
            EditTasks = False
        End Try

    End Function

    'done
    Function EditVariables(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objVar As clsVariable

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblVariables & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblVariables & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblVariables & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblVariables & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & _
                    "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_VARIABLE
                    objVar = CType(obj, clsVariable)
                    If objVar.Engine IsNot Nothing Then
                        sql = "Update " & objVar.Project.tblVariables & " Set VariableName= '" & NewValue & _
                                           "' WHERE VariableName= '" & OldValue & _
                                           "' AND ProjectName='" & objVar.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objVar.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objVar.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objVar.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objVar.Project.tblVariables & " Set VariableName= '" & NewValue & _
                                           "' WHERE VariableName= '" & OldValue & _
                                           "' AND ProjectName='" & objVar.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objVar.Environment.EnvironmentName & "'"
                    End If


                Case Else
                    EditVariables = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditVariables = True

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                EditVariables = EditVariablesATTR(cmd, OldValue, NewValue, obj)
            End If

        Catch ex As Exception
            LogError(ex, "modRename EditVariables", sql)
            EditVariables = False
        End Try

    End Function

    'done
    Function EditVariablesATTR(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Dim objProj As clsProject
        Dim objEnv As clsEnvironment
        Dim objSys As clsSystem
        Dim objEng As clsEngine
        Dim objVar As clsVariable

        Try
            Select Case obj.Type
                Case NODE_PROJECT
                    objProj = CType(obj, clsProject)
                    sql = "Update " & objProj.Project.tblVariablesATTR & " Set ProjectName= '" & NewValue & _
                    "' WHERE ProjectName= '" & OldValue & "'"

                Case NODE_ENVIRONMENT
                    objEnv = CType(obj, clsEnvironment)
                    sql = "Update " & objEnv.Project.tblVariablesATTR & " Set EnvironmentName= '" & NewValue & _
                    "' WHERE EnvironmentName= '" & OldValue & _
                    "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

                Case NODE_SYSTEM
                    objSys = CType(obj, clsSystem)
                    sql = "Update " & objSys.Project.tblVariablesATTR & " Set SystemName= '" & NewValue & _
                    "' WHERE SystemName= '" & OldValue & _
                    "' AND ProjectName='" & objSys.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSys.Environment.EnvironmentName & "'"

                Case NODE_ENGINE
                    objEng = CType(obj, clsEngine)
                    sql = "Update " & objEng.Project.tblVariablesATTR & " Set EngineName= '" & NewValue & _
                    "' WHERE EngineName= '" & OldValue & _
                    "' AND ProjectName='" & objEng.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objEng.ObjSystem.Environment.EnvironmentName & _
                    "' AND SystemName= '" & objEng.ObjSystem.SystemName & "'"

                Case NODE_VARIABLE
                    objVar = CType(obj, clsVariable)
                    If objVar.Engine IsNot Nothing Then
                        sql = "Update " & objVar.Project.tblVariablesATTR & " Set VariableName= '" & NewValue & _
                                           "' WHERE VariableName= '" & OldValue & _
                                           "' AND ProjectName='" & objVar.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objVar.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objVar.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objVar.Engine.EngineName & "'"
                    Else
                        sql = "Update " & objVar.Project.tblVariablesATTR & " Set VariableName= '" & NewValue & _
                                           "' WHERE VariableName= '" & OldValue & _
                                           "' AND ProjectName='" & objVar.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objVar.Environment.EnvironmentName & "'"
                    End If


                Case Else
                    EditVariablesATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditVariablesATTR = True

        Catch ex As Exception
            LogError(ex, "modRename EditVariablesATTR", sql)
            EditVariablesATTR = False
        End Try

    End Function

#End Region

#Region "Helper Functions"

    '/// checks to see that names are only alphanumeric plus a few allowed characters and start with a letter
    Function ValidateNewName(ByVal NewName As String, Optional ByVal IsEng As Boolean = False) As Boolean

        Dim Letter As Char  '// character of new name presently being analized
        Dim i As Integer   '// loop counter
        Dim loopRange As Integer = Len(NewName.Trim)  '// length of new name
        Dim asciiInt As Integer     '// the ascii code for the character being analized
        Dim FirstLetterFlag As Boolean = True  '// determines if the character is the first letter in the new name

        Try
            NewName = NewName.Trim
            If loopRange > 20 Then
                ValidateNewName = False
                MsgBox("Name cannot be more than 20 characters," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                Exit Function
            End If

            If NewName = "" Then
                ValidateNewName = False
                MsgBox("Name cannot be left blank," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                Exit Function
            End If

            If NewName.StartsWith("%") = True And IsEng = True Then
                If NewName.Length <> 3 Then
                    ValidateNewName = False
                    MsgBox("Substitution Variables Must be in the format '%XX'," & Chr(13) & _
                    "where each X is a digit from 0 thru 9." & Chr(13) & _
                    "New name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                    Exit Function
                End If
            Else
                'loopRange = Len(NewName)
                For i = 1 To loopRange
                    Letter = Left(NewName, 1)
                    asciiInt = Asc(Letter)
                    If FirstLetterFlag = True Then
                        If Not ((asciiInt >= 97 And asciiInt <= 122) Or (asciiInt >= 64 And asciiInt <= 90) Or (asciiInt = 38)) Then
                            ValidateNewName = False
                            MsgBox("Name must start with an alphabetic character, &, or @." & Chr(13) & "New name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                            Exit Function
                        End If
                    Else
                        If Not ((asciiInt >= 97 And asciiInt <= 122) Or (asciiInt >= 64 And asciiInt <= 90) Or (asciiInt >= 48 And asciiInt <= 57) Or (asciiInt = 38) Or (asciiInt = 35) Or (asciiInt = 45) Or (asciiInt = 95)) Then
                            ValidateNewName = False
                            MsgBox("Name must only consist of alpha-numeric characters, &, @, -, _, or #." & Chr(13) & "New Name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                            Exit Function
                        End If
                    End If
                    NewName = Right(NewName, Len(NewName) - 1)
                    FirstLetterFlag = False
                Next
            End If
            ValidateNewName = True

        Catch ex As Exception
            LogError(ex, "modRename ValidateNewName")
            Return False
        End Try

    End Function

    'done
    '/// updates 'Foreign Key' field in DSSELFIELDS Table
    Function EditFkey(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim sql As String = ""
        Try
            '/// initialize all the variables
            Dim dr As System.Data.Odbc.OdbcDataReader
            Dim i As Integer   '// item number for all arraylists to keep them in SYNC
            Dim objStr As clsStructure
            Dim objSel As clsStructureSelection
            Dim OldFKey As String
            Dim NewFKey As String
            Dim OldKeyFld As String
            Dim OldKeyStr As String
            '// these variables are just for keeping track of what exact field had it's fKey replaced
            '// so we can update it accordingly
            Dim DSName As String
            Dim EngName As String
            Dim SysName As String
            Dim SelName As String
            Dim StrName As String
            Dim EnvName As String
            Dim ProjName As String
            Dim FldName As String
            Dim ParName As String
            Dim DSDir As String = ""

            '// initialize all arraylists for all of the fields being read in
            Dim FldNameArray As New ArrayList
            Dim FkeyArray As New ArrayList
            Dim DSNameArray As New ArrayList
            Dim EngNameArray As New ArrayList
            Dim SysNameArray As New ArrayList
            Dim SelNameArray As New ArrayList
            Dim StrNameArray As New ArrayList
            Dim EnvNameArray As New ArrayList
            Dim ProjNameArray As New ArrayList
            Dim ParNameArray As New ArrayList
            Dim DSDirArray As New ArrayList

            '/// clear all the arraylists
            FldNameArray.Clear()
            FkeyArray.Clear()
            DSNameArray.Clear()
            EngNameArray.Clear()
            SysNameArray.Clear()
            SelNameArray.Clear()
            StrNameArray.Clear()
            EnvNameArray.Clear()
            ProjNameArray.Clear()
            ParNameArray.Clear()
            DSDirArray.Clear()

            Select Case obj.Type
                Case NODE_PROJECT, NODE_ENVIRONMENT, NODE_SYSTEM, NODE_ENGINE, NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                    EditFkey = True
                    Exit Function

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "Select * FROM " & objStr.Project.tblDSselFields & _
                    " WHERE ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "Select * FROM " & objSel.Project.tblDSselFields & _
                    " WHERE ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"

                Case Else
                    EditFkey = False
                    Exit Function
            End Select

            '// if obj is a structure or structure selection then by default the DSSelection will be named the same
            '// so read in all foreign keys for the environment and look for matches
            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                OldFKey = dr.Item("ForeignKey")
                DSName = dr.Item("DatastoreName")
                EngName = dr.Item("EngineName")
                SysName = dr.Item("SystemName")
                SelName = dr.Item("SelectionName")
                EnvName = dr.Item("EnvironmentName")
                ProjName = dr.Item("ProjectName")
                FldName = dr.Item("FieldName")
                If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    ParName = dr.Item("Parent")
                    DSDir = dr.Item("DsDirection")
                    StrName = dr.Item("StructureName")
                Else
                    ParName = dr.Item("ParentName")
                    StrName = dr.Item("DescriptionName")
                End If


                '// split the foreign key into it's components (structure.Field) and compare DSSelection part of the 
                '// name with the old Value and determine if it should be replaced with the new name
                If OldFKey <> "" Then
                    OldKeyStr = Split(OldFKey, ".")(0)
                    OldKeyFld = Split(OldFKey, ".")(1)
                Else
                    OldKeyStr = ""
                    OldKeyFld = ""
                End If

                '// If the fKey needs replacing then add all other field values to their respective arraylists
                '// so that when the update occurs, the new fKey is put back exactly where it needs to go
                If (OldKeyStr = OldValue) And Not (OldKeyStr = "" Or OldKeyFld = "") Then
                    NewFKey = NewValue & "." & OldKeyFld
                    FkeyArray.Add(NewFKey)
                    FldNameArray.Add(FldName)
                    DSNameArray.Add(DSName)
                    EngNameArray.Add(EngName)
                    SysNameArray.Add(SysName)
                    SelNameArray.Add(SelName)
                    StrNameArray.Add(StrName)
                    EnvNameArray.Add(EnvName)
                    ProjNameArray.Add(ProjName)
                    ParNameArray.Add(ParName)
                    If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                        DSDirArray.Add(DSDir)
                    End If
                End If

            End While
            dr.Close()  '// must be closed when done so that the rest of the "non-query" part of the 
            '// transaction can continue

            '// Now that we have made new foreign keys using the new name
            '// we will update the DSselFields Table with the new Fkeys where all the other fields read in
            '// match, so that we put the right fKey in the right selection field
            For i = 0 To FkeyArray.Count - 1
                If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    sql = "Update " & obj.Project.tblDSselFields & _
                    " Set ForeignKey= '" & FkeyArray(i) & _
                    "' WHERE FieldName= '" & FldNameArray(i) & _
                    "' AND DatastoreName= '" & DSNameArray(i) & _
                    "' AND EngineName= '" & EngNameArray(i) & _
                    "' AND SystemName= '" & SysNameArray(i) & _
                    "' AND SelectionName= '" & SelNameArray(i) & _
                    "' AND StructureName= '" & StrNameArray(i) & _
                    "' AND EnvironmentName= '" & EnvNameArray(i) & _
                    "' AND ProjectName= '" & ProjNameArray(i) & _
                    "' AND Parent= '" & ParNameArray(i) & _
                    "' AND DsDirection= '" & DSDirArray(i) & "'"
                Else
                    sql = "Update " & obj.Project.tblDSselFields & _
                    " Set ForeignKey= '" & FkeyArray(i) & _
                    "' WHERE FieldName= '" & FldNameArray(i) & _
                    "' AND DatastoreName= '" & DSNameArray(i) & _
                    "' AND EngineName= '" & EngNameArray(i) & _
                    "' AND SystemName= '" & SysNameArray(i) & _
                    "' AND SelectionName= '" & SelNameArray(i) & _
                    "' AND DescriptionName= '" & StrNameArray(i) & _
                    "' AND EnvironmentName= '" & EnvNameArray(i) & _
                    "' AND ProjectName= '" & ProjNameArray(i) & _
                    "' AND ParentName= '" & ParNameArray(i) & "'"
                End If


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            EditFkey = True

        Catch ex As Exception
            LogError(ex, "modRename EditFKey", sql)
            EditFkey = False
        End Try

    End Function

    'done
    '// called to update the MappingSource "memo" field of TASKMAP Table
    '// this is essentially the CASE statement in the main procedure.
    Function EditMappingSrc(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

        Dim dr As System.Data.Odbc.OdbcDataReader = Nothing
        Dim objSDS As clsDatastore
        Dim objTDS As clsDatastore
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection
        Dim objTask As clsTask
        Dim objVar As clsVariable
        Dim sql As String
        Dim MapString As String
        Dim SeqNo As String
        Dim TaskName As String
        Dim EngineName As String
        Dim SystemName As String
        Dim EnvName As String
        Dim ProjectName As String
        Dim i As Integer

        Dim SeqNoArray As New ArrayList
        Dim MapSourceArray As New ArrayList
        Dim TaskNameArray As New ArrayList
        Dim EngineNameArray As New ArrayList
        Dim SystemNameArray As New ArrayList
        Dim EnvironmentNameArray As New ArrayList
        Dim ProjectNameArray As New ArrayList

        Try
            SeqNoArray.Clear()
            MapSourceArray.Clear()
            TaskNameArray.Clear()
            EngineNameArray.Clear()
            SystemNameArray.Clear()
            EnvironmentNameArray.Clear()
            ProjectNameArray.Clear()

            Select Case obj.Type

                Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                    objTask = CType(obj, clsTask)
                    If objTask.Engine IsNot Nothing Then
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " _
                                           & objTask.Project.tblTaskMap & _
                                           " WHERE ProjectName='" & objTask.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objTask.Engine.EngineName & "'"
                    Else
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " _
                                           & objTask.Project.tblTaskMap & _
                                           " WHERE ProjectName='" & objTask.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    End If


                Case NODE_SOURCEDATASTORE
                    objSDS = CType(obj, clsDatastore)
                    If objSDS.Engine IsNot Nothing Then
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                            objSDS.Project.tblTaskMap & _
                                            " WHERE ProjectName='" & objSDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objSDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objSDS.Engine.EngineName & "'"
                    Else
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                            objSDS.Project.tblTaskMap & _
                                            " WHERE ProjectName='" & objSDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & "'"
                    End If


                Case NODE_TARGETDATASTORE
                    objTDS = CType(obj, clsDatastore)
                    If objTDS.Engine IsNot Nothing Then
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                           objTDS.Project.tblTaskMap & _
                                           " WHERE ProjectName='" & objTDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objTDS.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objTDS.Engine.EngineName & "'"
                    Else
                        sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                           objTDS.Project.tblTaskMap & _
                                           " WHERE ProjectName='" & objTDS.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & "'"
                    End If


                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    objStr.Project.tblTaskMap & _
                    " WHERE ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    objSel.Project.tblTaskMap & _
                    " WHERE ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"

                Case NODE_VARIABLE
                    objvar = CType(obj, clsVariable)
                    sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    objVar.Project.tblTaskMap & _
                    " WHERE ProjectName='" & objVar.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objVar.Environment.EnvironmentName & "'"

                Case Else
                    EditMappingSrc = False
                    Exit Function
            End Select


            '/// Not Very Elegant, but it works
            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            '// read in all the appropriate fields so that we know where the MappingSource field came from
            While dr.Read
                MapString = GetVal(dr.Item("MappingSource"))
                SeqNo = GetVal(dr.Item("SEQNO"))
                TaskName = GetVal(dr.Item("TASKNAME"))
                EngineName = GetVal(dr.Item("ENGINENAME"))
                SystemName = GetVal(dr.Item("SYSTEMNAME"))
                EnvName = GetVal(dr.Item("ENVIRONMENTNAME"))
                ProjectName = GetVal(dr.Item("PROJECTNAME"))

                Dim NewMapSrc As New System.Text.StringBuilder(MapString)
                Dim MapSource As New System.Text.StringBuilder(MapString)

                Select Case obj.Type
                    Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                        NewMapSrc.Replace(OldValue, NewValue)

                        '// now do UGLY replacements 
                        '// could not get regular expressions to work correctly here
                    Case NODE_SOURCEDATASTORE, NODE_TARGETDATASTORE
                        NewMapSrc.Replace("(" & OldValue & ")", "(" & NewValue & ")")
                        NewMapSrc.Replace("(" & OldValue & ".", "(" & NewValue & ".")
                        NewMapSrc.Replace("(" & OldValue & " ", "(" & NewValue & " ")
                        NewMapSrc.Replace("(" & OldValue & ",", "(" & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & ")", " " & NewValue & ")")
                        NewMapSrc.Replace(" " & OldValue & ",", " " & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & ".", " " & NewValue & ".")
                        NewMapSrc.Replace("," & OldValue & ")", "," & NewValue & ")")
                        NewMapSrc.Replace("," & OldValue & " ", "," & NewValue & " ")

                    Case NODE_STRUCT, NODE_STRUCT_SEL, NODE_VARIABLE
                        NewMapSrc.Replace("'" & OldValue & "'", "'" & NewValue & "'")
                        NewMapSrc.Replace("." & OldValue & ".", "." & NewValue & ".")
                        NewMapSrc.Replace("(" & OldValue & " ", "(" & NewValue & " ")
                        NewMapSrc.Replace("(" & OldValue & ",", "(" & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & ")", " " & NewValue & ")")
                        NewMapSrc.Replace(" " & OldValue & ",", " " & NewValue & ",")
                        NewMapSrc.Replace("," & OldValue & ")", "," & NewValue & ")")
                        NewMapSrc.Replace("," & OldValue & " ", "," & NewValue & " ")

                End Select

                '// if nothing was replaced it means that old value didn't exist in that 
                '// mappingSource field, so ignore it and go to the next match
                '// if the old Mapping Source Field doesn't match the New "replaced" Mapping Source
                '// field, then we will need to update it, so save all of the row data to 
                '// arraylists so they can be replaced one at a time in the loop below
                If Not (NewMapSrc.ToString = MapSource.ToString) Then
                    TaskNameArray.Add(TaskName)
                    SeqNoArray.Add(SeqNo)
                    MapSourceArray.Add(NewMapSrc.ToString)
                    EngineNameArray.Add(EngineName)
                    SystemNameArray.Add(SystemName)
                    EnvironmentNameArray.Add(EnvName)
                    ProjectNameArray.Add(ProjectName)
                End If
            End While

        Catch ex As Exception
            LogError(ex, "editmappingsrc modrename")
            Return False
            Exit Function
        Finally
            dr.Close()
        End Try

        '// replacement of Old names with new ones id complete
        '// so now Update all of the affected records.
        Try
            For i = 0 To SeqNoArray.Count - 1
                sql = "UPDATE " & obj.Project.tblTaskMap & _
                " SET MappingSource= '" & FixStr(MapSourceArray(i)) & _
                "' WHERE SEQNO= " & CInt(SeqNoArray(i)) & _
                " AND TASKNAME= '" & FixStr(TaskNameArray(i)) & _
                "' AND ENGINENAME='" & FixStr(EngineNameArray(i)) & _
                "' AND SYSTEMNAME='" & FixStr(SystemNameArray(i)) & _
                "' AND ENVIRONMENTNAME='" & FixStr(EnvironmentNameArray(i)) & _
                "' AND PROJECTNAME='" & FixStr(ProjectNameArray(i)) & "'"

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            EditMappingSrc = True

        Catch ex As Exception
            LogError(ex, "modRename EditMappingSource")
            EditMappingSrc = False
        End Try

        '/// sample of using Regular expressions to solve this problem
        'Dim RegExp1 As New System.Text.RegularExpressions.Regex("(^|[^-.0-9A-Za-z_])" & OldValue & "([^-_0-9A-Za-z]|$)")
        'Dim RegExp2 As New System.Text.RegularExpressions.Regex(MapString)

    End Function

#End Region

    '// functions in this region are called from the frmStructure form 
    '// when replacing structure files ... by Tom Karasch April and June 2007
#Region "Replace structure file"

    '/// added by TKarasch April 07 to replace Structure file ... 
    '// Main replacement function ... Needs refining
    '**********************************************************
    '/// Refined 1/09 to 4/09 by TKarasch
    Function ReplaceStrFile(ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure) As Boolean

        '// return value from "compareOldNewFldNames" function to determine status for replacement
        Dim IntRet As Integer
        '// are the new field names different form the old field names
        Dim HasBadNames As Boolean = False
        '// Arraylists to get populated by CompareOldNewFldNames function
        '***** Array of New Fields in Incoming description that don't exist in the existing description
        Dim AddedFieldList As New ArrayList
        '***** Array of Old Fields in current description that don't exist in the existing description
        Dim DeletedFieldList As New ArrayList

        AddedFieldList.Clear()
        DeletedFieldList.Clear()

        Try
            IntRet = CompareOldNewFldNames(NewStrObj.ObjFields, OldStrObj.ObjFields, AddedFieldList, DeletedFieldList)

            If IntRet = -2 Then
                MsgBox("An Error Occured During replacement, Please Check your replacement file" & Chr(13) & "Replacement will not take place", MsgBoxStyle.Information, MsgTitle)
                Return False
                Exit Function
            End If



            Dim frm As frmRplcDescRet
            Dim Replace As Boolean

            frm = New frmRplcDescRet
            Replace = frm.ShowFldDiff(NewStrObj, OldStrObj, AddedFieldList, DeletedFieldList)

            '//// new 1/15/10
            If AddedFieldList.Count > 0 Or DeletedFieldList.Count > 0 Then
                HasBadNames = True
            End If

            If Replace = True Then
                ReplaceStrFile = SQLReplaceFile(NewStrObj, OldStrObj, HasBadNames)
            Else
                ReplaceStrFile = False
            End If

        Catch ex As Exception
            LogError(ex, "modRename ReplaceStrFile")
            ReplaceStrFile = False
        End Try

    End Function

    '/// added by TKarasch April 07 to replace Structure file ... to compare field names
    Function CompareOldNewFldNames(ByVal NewArr As ArrayList, ByVal OldArr As ArrayList, ByRef AddFldList As ArrayList, ByRef DelFldList As ArrayList) As Integer

        Dim i, j As Integer
        Dim NewFld, Oldfld As clsField
        Dim BadFldFlag As Boolean = True


        Try
            '// Init Delfldlist
            For Each fld As clsField In OldArr
                DelFldList.Add(fld)
            Next

            '// loop through the field arraylist of the structure with fewer fields and look
            '// for matching field names in the larger structure field arraylist.
            '// if no match is found for a field in the smaller array, then we will need to
            '// warn the user that if this structure is replaced, all the associated 
            '// Datastore uses and mappings will be deleted
            For i = 0 To NewArr.Count - 1
                NewFld = CType(NewArr(i), clsField)
                For j = 0 To OldArr.Count - 1
                    Oldfld = CType(OldArr(j), clsField)

                    If NewFld.FieldName = Oldfld.FieldName Then
                        DelFldList.Remove(Oldfld)
                        BadFldFlag = False
                        Exit For
                    End If
                Next
                If BadFldFlag = True Then
                    '// Add NewFld to the Added Field List
                    AddFldList.Add(NewFld)
                End If
                BadFldFlag = True
            Next

            '/// Means OK to replace
            CompareOldNewFldNames = 1

        Catch ex As Exception
            LogError(ex, "modRename compareOldNewFieldNames")
            CompareOldNewFldNames = -2
        End Try

    End Function

    '/// added by TKarasch April 07 to replace Structure file ... to update MetaData
    Function SQLReplaceFile(ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure, ByVal HasBadNames As Boolean) As Boolean

        '// All Comparison testing of fields has been done before we open Metadata
        '// so we are sure that Updating of metadata will work without errors
        '// and all deletes and adds will happen in one Transaction, so rollback is possible

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim cascade As Boolean = False
        Dim Step1 As Boolean
        Dim success As Boolean = True

        Try
            '// start SQL Tran
            'cnn = New Odbc.OdbcConnection(NewStrObj.Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            '// First set cascade based on whether fields are to be replaced one to one or not
            If HasBadNames = False Then
                '// Only delete OldStrObj from two tables and add NewStrObj back in it's place
                cascade = False
            Else
                '// Delete OldStrObj from Structures, Datastores and Mappings
                '// Delete all selections, structure, coresponding DSselections, and mappings
                '// Add back in to Structures and StructFields Table ONLY, because number of 
                '// fields Has Changed, so we will not guess on Selections, selected fields,
                '/// mappings, etc...
                cascade = True
            End If

            '// Now delete Structure and fields from metadata and parent collections based on cascade
            '// If either delete or add operations fail, then ROLLBACK transaction
            If OldStrObj.Delete(cmd, cnn, cascade, cascade) = True Then
                If cascade = False Then
                    RemoveFromCollection(NewStrObj.Environment.Structures, OldStrObj.GUID)
                    '///**** Very important ****
                    '/// now fields pointers must be swapped in all of the correponding mapping 
                    '// objects so that tasks and mappings will recognize the new fields
                    If UpdateMappings(NewStrObj) = False Then
                        tran.Rollback()
                        Step1 = False
                    End If
                End If
                '// Now add the new Structure properties back in to metadata
                If NewStrObj.AddNew(cmd) = True Then
                    '// Now add any additional fields to the DSSelectionFields Table
                    If UpdateDSSelFldsForStrChng(cnn, cmd, NewStrObj, OldStrObj) = True Then
                        '// All went well, then commit
                        tran.Commit()
                        Step1 = True
                    Else
                        tran.Rollback()
                        Step1 = False
                    End If
                Else
                    tran.Rollback()
                    Step1 = False
                End If
            Else
                tran.Rollback()
                Step1 = False
            End If

            If Step1 = True Then
                For Each DSsel As clsDSSelection In NewStrObj.SysAllSelection.ObjDSselections
                    If DSsel.ObjDatastore.Save() = False Then
                        success = False
                    End If
                Next
                If success = False Then
                    MsgBox("There were problems changing the description in your datastores." & Chr(13) & _
                    "Please check your Datastores and Procedures for possible" & Chr(13) & _
                    "problems related to changing your description file.", MsgBoxStyle.Exclamation, "Problems changing Description file")
                    Return False
                    Exit Try
                Else
                    Return True
                End If
            Else
                Return False
                Exit Try
            End If

        Catch ex As Exception
            LogError(ex, "modRename SQLReplaceFile")
            tran.Rollback()
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '/// function used when replacing structure files.
    '// this function updates the mapping source and target fields
    '// that are in mappings that use the replaced structure 
    Function UpdateMappings(ByRef NewStructObj As clsStructure) As Boolean

        Try
            Dim Obj As Object

            For Each fld As clsField In NewStructObj.ObjFields
                For Each sys As clsSystem In NewStructObj.Environment.Systems
                    For Each eng As clsEngine In sys.Engines
                        'For Each task As clsTask In eng.Mains
                        '    For Each map As clsMapping In task.ObjMappings
                        '        Obj = map.MappingSource
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingSource = fld
                        '            End If
                        '        End If
                        '        Obj = map.MappingTarget
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingTarget = fld
                        '            End If
                        '        End If
                        '    Next
                        'Next
                        'For Each task As clsTask In eng.Lookups
                        '    For Each map As clsMapping In task.ObjMappings
                        '        Obj = map.MappingSource
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingSource = fld
                        '            End If
                        '        End If
                        '        Obj = map.MappingTarget
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingTarget = fld
                        '            End If
                        '        End If
                        '    Next
                        'Next
                        'For Each task As clsTask In eng.Joins
                        '    For Each map As clsMapping In task.ObjMappings
                        '        Obj = map.MappingSource
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingSource = fld
                        '            End If
                        '        End If
                        '        Obj = map.MappingTarget
                        '        If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                        '            If CType(Obj, clsField).FieldName = fld.FieldName Then
                        '                map.MappingTarget = fld
                        '            End If
                        '        End If
                        '    Next
                        'Next
                        For Each task As clsTask In eng.Procs
                            For Each map As clsMapping In task.ObjMappings
                                Obj = map.MappingSource
                                If Obj IsNot Nothing Then
                                    If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                        If CType(Obj, clsField).FieldName = fld.FieldName Then
                                            map.MappingSource = fld
                                        End If
                                    End If
                                End If
                                Obj = map.MappingTarget
                                If Obj IsNot Nothing Then
                                    If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                        If CType(Obj, clsField).FieldName = fld.FieldName Then
                                            map.MappingTarget = fld
                                        End If
                                    End If
                                End If
                            Next
                        Next
                    Next
                Next
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "UpdateMappings modRename")
            Return False
        End Try

    End Function

    Function UpdateDSSelFldsForStrChng(ByRef cnn As Odbc.OdbcConnection, ByRef cmd As Odbc.OdbcCommand, ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure) As Boolean

        Try
            If NewStrObj.ObjFields.Count = OldStrObj.ObjFields.Count Then
                Return True
                Exit Function
            End If

            For Each DSSel As clsDSSelection In OldStrObj.SysAllSelection.ObjDSselections
                Dim DS As clsDatastore = DSSel.ObjDatastore
                If DSSel.Delete(cmd, cnn, True, False) = True Then
                    Dim NewDSsel As clsDSSelection
                    cmd.Connection = cnn
                    NewDSsel = DS.CloneSSeltoDSSel(NewStrObj.SysAllSelection, , True, cmd)
                    DS.ObjSelections.Add(NewDSsel)
                Else
                    Return False
                    Exit Try
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "modRename UpdateDSSelFldsForStrChng")
            Return False
        End Try

    End Function

#End Region

End Module
