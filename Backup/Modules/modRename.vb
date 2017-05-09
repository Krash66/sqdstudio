Module modRename
    'This module is intended to enable renaming of all objects with Names stored in Metadata
    'Added January 2007 by Tom Karasch  
    'Modified May 2007 For Structure File Replacement
    'Modified July 2007 for Moving of connections to Environment
    'Modified August 2008
    'Modified Jan thru May 2009 for V3 and new Object Model changes
    'Modified Nov. 2011 to correct problems with Task renaming after "General" Proc type added

#Region "Main Rename Functions"

    Function RenameProject(ByRef Proj As clsProject, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        'PrevObjTreeNode = Proj.ObjTreeNode

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


                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameProject = True
                Else
                    tran.Rollback()
                    RenameProject = False
                End If

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameProject")
                RenameProject = False
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

        'PrevObjTreeNode = Env.ObjTreeNode

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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameEnvironment")
                RenameEnvironment = False
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

        'PrevObjTreeNode = sys.ObjTreeNode

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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameSystem")
                RenameSystem = False
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

        'PrevObjTreeNode = eng.ObjTreeNode

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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameEngine")
                RenameEngine = False
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

        'PrevObjTreeNode = struct.ObjTreeNode

        If ValidateNewName128(NewName) = False Then
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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameStructure")
                RenameStructure = False
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

        'PrevObjTreeNode = StrSel.ObjTreeNode

        If ValidateNewName128(NewName) = False Then
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
                'If success = True Then
                '    success = EditDESCFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                'End If
                If success = True Then
                    success = EditDESCSELFIELDS(cmd, StrSel.SelectionName, NewName, StrSel)
                End If
                If success = True Then
                    success = EditDESCSelections(cmd, StrSel.SelectionName, NewName, StrSel)
                End If
                'If success = True Then
                '    success = EditTaskDatastores(cmd, StrSel.SelectionName, NewName, StrSel)
                'End If
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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameStructureSelection")
                RenameStructureSelection = False
            Catch ex As Exception
                tran.Rollback()
                LogError(ex, "modRename RenameStructureSelection")
                RenameStructureSelection = False
            Finally
                'cnn.Close()
                StrSel.LoadMe()
            End Try
        End If

    End Function

    Function RenameConnection(ByRef Conn As clsConnection, ByVal NewName As String) As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        'PrevObjTreeNode = Conn.ObjTreeNode

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
                    success = EditDescriptions(cmd, Conn.ConnectionName, NewName, Conn)
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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameConnection")
                RenameConnection = False
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

        'PrevObjTreeNode = DS.ObjTreeNode

        If ValidateNewName128(NewName) = False Then
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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameDatastore")
                RenameDatastore = False
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

    Function RenameTask(ByRef Task As clsTask, ByVal NewName As String, Optional ByVal Oldname As String = "") As Boolean

        Dim success As Boolean = False
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing

        'PrevObjTreeNode = Task.ObjTreeNode
        Dim NameOld As String = Task.TaskName
        If Oldname <> "" Then
            NameOld = Oldname
        End If

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

                success = EditTaskDatastores(cmd, NameOld, NewName, Task)
                If success = True Then
                    success = EditTaskMappings(cmd, NameOld, NewName, Task)
                End If
                If success = True Then
                    success = EditTasks(cmd, NameOld, NewName, Task)
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

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameTask")
                RenameTask = False
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

        'PrevObjTreeNode = Var.ObjTreeNode

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
                    success = EditTaskMappings(cmd, Var.VariableName, NewName, Var)
                End If

                If success = True Then
                    tran.Commit()
                    NameOfNodeBefore = NewName
                    RenameVariable = True
                Else
                    tran.Rollback()
                    RenameVariable = False
                End If

            Catch OE As Odbc.OdbcException
                tran.Rollback()
                LogODBCError(OE, "modRename RenameVariable")
                RenameVariable = False
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

            EditConnections = EditConnectionATTR(cmd, OldValue, NewValue, obj)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    EditConnections = EditConnectionATTR(cmd, OldValue, NewValue, obj)
            'End If

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditConnections", sql)
            EditConnections = False
        Catch ex As Exception
            LogError(ex, "modRename EditConnections", sql)
            EditConnections = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditConnectionATTR", sql)
            EditConnectionATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditConnectionsATTR", sql)
            EditConnectionATTR = False
        End Try

    End Function

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
                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDatastores & " Set DatastoreName= '" & NewValue & _
                                       "' WHERE DatastoreName= '" & OldValue & _
                                       "' AND ProjectName='" & objDS.Project.ProjectName & _
                                       "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                       "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                       "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    '    sql = "Update " & objDS.Project.tblDatastores & " Set DatastoreName= '" & NewValue & _
                    '                       "' WHERE DatastoreName= '" & OldValue & _
                    '                       "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                       "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If


                Case Else
                    EditDatastores = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDatastores = EditDatastoresATTR(cmd, OldValue, NewValue, obj)

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDatastores", sql)
            EditDatastores = False
        Catch ex As Exception
            LogError(ex, "modRename EditDatastores", sql)
            EditDatastores = False
        End Try

    End Function

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
                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDatastoresATTR & " Set DatastoreName= '" & NewValue & _
                                        "' WHERE DatastoreName= '" & OldValue & _
                                        "' AND ProjectName='" & objDS.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblDatastoresATTR & _
                    '" Set DatastoreName= '" & NewValue & _
                    '                    "' WHERE DatastoreName= '" & OldValue & _
                    '                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If


                Case Else
                    EditDatastoresATTR = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDatastoresATTR = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDatastoresATTR", sql)
            EditDatastoresATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditDatastoresATTR", sql)
            EditDatastoresATTR = False
        End Try

    End Function

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
                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDSselFields & " Set DatastoreName= '" & NewValue & _
                                       "' WHERE DatastoreName= '" & OldValue & _
                                       "' AND ProjectName='" & objDS.Project.ProjectName & _
                                       "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                       "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                       "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblDSselFields & " Set DatastoreName= '" & NewValue & _
                    '                   "' WHERE DatastoreName= '" & OldValue & _
                    '                   "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                   "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If


                    'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    cmd.CommandText = sql
                    '    Log(sql)
                    '    cmd.ExecuteNonQuery()

                    '    If objDS.Engine IsNot Nothing Then
                    '        sql = "Update " & objDS.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                    '                                "' WHERE Parent= '" & OldValue & _
                    '                                "' AND DatastoreName='" & NewValue & _
                    '                                "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                                "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                    '                                "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                    '                                "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    '    Else
                    '        sql = "Update " & objDS.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                    '                                "' WHERE Parent= '" & OldValue & _
                    '                                "' AND DatastoreName='" & NewValue & _
                    '                                "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                                "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    '    End If
                    'Else
                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDSselFields & " Set ParentName= '" & NewValue & _
                                            "' WHERE ParentName= '" & OldValue & _
                                            "' AND DatastoreName='" & NewValue & _
                                            "' AND ProjectName='" & objDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblDSselFields & " Set ParentName= '" & NewValue & _
                    '                        "' WHERE ParentName= '" & OldValue & _
                    '                        "' AND DatastoreName='" & NewValue & _
                    '                        "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                        "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If
                    'End If

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    'If objStr.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "Update " & objStr.Project.tblDSselFields & _
                    '    " Set StructureName= '" & NewValue & _
                    '    "',SelectionName= '" & NewValue & _
                    '    "' WHERE StructureName= '" & OldValue & _
                    '    "' AND SelectionName='" & OldValue & _
                    '    "' AND ProjectName='" & objStr.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    '    cmd.CommandText = sql
                    '    Log(sql)
                    '    cmd.ExecuteNonQuery()

                    '    sql = "Update " & objStr.Project.tblDSselFields & " Set StructureName= '" & NewValue & _
                    '    "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    '    cmd.CommandText = sql
                    '    Log(sql)
                    '    cmd.ExecuteNonQuery()

                    '    sql = "Update " & objStr.Project.tblDSselFields & " Set Parent= '" & NewValue & _
                    '    "' WHERE Parent= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"
                    'Else
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
                    'End If


                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    'If objSel.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "Update " & objSel.Project.tblDSselFields & " Set SelectionName= '" & NewValue & _
                    '    "' WHERE SelectionName= '" & OldValue & "' AND ProjectName='" & objSel.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                    '    "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

                    '    cmd.CommandText = sql
                    '    Log(sql)
                    '    cmd.ExecuteNonQuery()

                    '    sql = "Update " & objSel.Project.tblDSselFields & _
                    '    " Set Parent= '" & NewValue & _
                    '    "' WHERE Parent= '" & OldValue & _
                    '    "' AND ProjectName='" & objSel.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"
                    'Else
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
                    'End If


                Case Else
                    EditDSSELFIELDS = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditDSSELFIELDS = EditFkey(cmd, OldValue, NewValue, obj)

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDSSELFIELDS", sql)
            EditDSSELFIELDS = False
        Catch ex As Exception
            LogError(ex, "modRename EditDSselFields", sql)
            EditDSSELFIELDS = False
        End Try

    End Function

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

                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDSselections & " Set DatastoreName= '" & NewValue & _
                                       "' WHERE DatastoreName= '" & OldValue & _
                                       "' AND ProjectName='" & objDS.Project.ProjectName & _
                                       "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                       "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                       "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblDSselections & " Set DatastoreName= '" & NewValue & _
                    '                   "' WHERE DatastoreName= '" & OldValue & _
                    '                   "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                   "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblDSselections & _
                                        " Set Parent= '" & NewValue & _
                                        "' WHERE DatastoreName= '" & NewValue & _
                                        "' AND Parent= '" & OldValue & _
                                        "' AND ProjectName='" & objDS.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblDSselections & _
                    '                    " Set Parent= '" & NewValue & _
                    '                    "' WHERE DatastoreName= '" & NewValue & _
                    '                    "' AND Parent= '" & OldValue & _
                    '                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If


                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    'If objStr.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "Update " & objStr.Project.tblDSselections & " Set StructureName= '" & NewValue & _
                    '                        "',  SelectionName= '" & NewValue & "' WHERE StructureName= '" & OldValue & _
                    '                        "' AND SelectionName='" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    '                        "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    '    cmd.CommandText = sql
                    '    Log(sql)
                    '    cmd.ExecuteNonQuery()

                    '    sql = "Update " & objStr.Project.tblDSselections & " Set StructureName= '" & NewValue & _
                    '    "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                    'Else
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

                    'End If

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()

                    sql = "Update " & objStr.Project.tblDSselections & " Set Parent= '" & NewValue & _
                    "' WHERE Parent= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    'If objSel.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "Update " & objSel.Project.tblDSselections & _
                    '    " Set SelectionName= '" & NewValue & _
                    '    "' WHERE SelectionName= '" & OldValue & _
                    '    "' AND ProjectName='" & objSel.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                    '    "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"
                    'Else
                    sql = "Update " & objSel.Project.tblDSselections & _
                    " Set SelectionName= '" & NewValue & _
                    "' WHERE SelectionName= '" & OldValue & _
                    "' AND ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & _
                    "' AND DescriptionName= '" & objSel.ObjStructure.StructureName & "'"
                    'End If

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDSSELECTIONS", sql)
            EditDSSELECTIONS = False
        Catch ex As Exception
            LogError(ex, "modRename EditDSSELECTIONS", sql)
            EditDSSELECTIONS = False
        End Try

    End Function

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
                    'If ObjConn.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    sql = "Update " & ObjConn.Project.tblEngines & _
                    '    " Set ConnectionName= '" & NewValue & _
                    '    "' WHERE ConnectionName= '" & OldValue & _
                    '    "' AND ProjectName='" & ObjConn.Project.ProjectName & _
                    '    "' AND EnvironmentName= '" & ObjConn.Environment.EnvironmentName & "'"
                    'Else
                    GoTo goto1
                    'End If

                Case Else
                    EditEngines = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

goto1:
            EditEngines = EditEnginesATTR(cmd, OldValue, NewValue, obj)

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditEngines", sql)
            EditEngines = False
        Catch ex As Exception
            LogError(ex, "modRename EditEngines", sql)
            EditEngines = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditEnginesATTR", sql)
            EditEnginesATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditEnginesATTR", sql)
            EditEnginesATTR = False
        End Try

    End Function

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

            EditEnvironments = EditEnvironmentsATTR(cmd, OldValue, NewValue, obj)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    EditEnvironments = 
            'End If

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditEnvironments", sql)
            EditEnvironments = False
        Catch ex As Exception
            LogError(ex, "modRename EditEnvironments", sql)
            EditEnvironments = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditEnvironmentsATTR", sql)
            EditEnvironmentsATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditEnvironmentsATTR", sql)
            EditEnvironmentsATTR = False
        End Try

    End Function

    Function EditProjects(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean
        Dim sql As String = ""

        Try
            sql = "Update " & obj.Project.tblProjects & _
            " Set ProjectName= '" & NewValue & _
            "' WHERE ProjectName= '" & OldValue & "'"

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditProjects = EditProjectsATTR(cmd, OldValue, NewValue, obj)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    EditProjects = 
            'End If

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditProjects", sql)
            EditProjects = False
        Catch ex As Exception
            LogError(ex, "modRename EditProjects", sql)
            EditProjects = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditProjectsATTR", sql)
            EditProjectsATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditProjectsATTR", sql)
            EditProjectsATTR = False
        End Try

    End Function

    'done
    'Function EditSTRUCTFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

    '    Dim sql As String = ""
    '    Dim objProj As clsProject
    '    Dim objEnv As clsEnvironment
    '    Dim objStr As clsStructure

    '    Try
    '        Select Case obj.Type
    '            Case NODE_PROJECT
    '                objProj = CType(obj, clsProject)
    '                sql = "Update " & objProj.Project.tblStructFields & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

    '            Case NODE_ENVIRONMENT
    '                objEnv = CType(obj, clsEnvironment)
    '                sql = "Update " & objEnv.Project.tblStructFields & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

    '            Case NODE_STRUCT
    '                '/// sets structure name and parent name
    '                objStr = CType(obj, clsStructure)

    '                sql = "Update " & objStr.Project.tblStructFields & _
    '                " Set StructureName= '" & NewValue & _
    '                "',ParentName= '" & NewValue & _
    '                "' WHERE StructureName= '" & OldValue & _
    '                "' AND ProjectName='" & objStr.Project.ProjectName & _
    '                "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

    '            Case Else
    '                EditSTRUCTFIELDS = False
    '                Exit Function
    '        End Select

    '        cmd.CommandText = sql
    '        Log(sql)
    '        cmd.ExecuteNonQuery()

    '        EditSTRUCTFIELDS = True

    '    Catch ex As Exception
    '        LogError(ex, "modRename EditStructFields", sql)
    '        EditSTRUCTFIELDS = False
    '    End Try

    'End Function

    'done
    'Function EditStructures(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

    '    Dim sql As String = ""
    '    Dim objProj As clsProject
    '    Dim objEnv As clsEnvironment
    '    Dim objStr As clsStructure
    '    Dim ObjConn As clsConnection

    '    Try
    '        Select Case obj.Type
    '            Case NODE_PROJECT
    '                objProj = CType(obj, clsProject)
    '                sql = "Update " & objProj.Project.tblStructures & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

    '            Case NODE_ENVIRONMENT
    '                objEnv = CType(obj, clsEnvironment)
    '                sql = "Update " & objEnv.Project.tblStructures & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

    '            Case NODE_STRUCT
    '                objStr = CType(obj, clsStructure)
    '                sql = "Update " & objStr.Project.tblStructures & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

    '            Case NODE_CONNECTION
    '                objConn = CType(obj, clsConnection)
    '                sql = "Update " & ObjConn.Project.tblStructures & " Set ConnectionName= '" & NewValue & "' WHERE ConnectionName= '" & OldValue & "' AND ProjectName='" & ObjConn.Project.ProjectName & "' AND EnvironmentName= '" & ObjConn.Environment.EnvironmentName & "'"

    '            Case Else
    '                EditStructures = False
    '                Exit Function
    '        End Select

    '        cmd.CommandText = sql
    '        Log(sql)
    '        cmd.ExecuteNonQuery()

    '        EditStructures = True

    '    Catch ex As Exception
    '        LogError(ex, "modRename EditStructures", sql)
    '        EditStructures = False
    '    End Try

    'End Function

    'done
    'Function EditSTRSELFIELDS(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

    '    Dim sql As String = ""
    '    Dim objProj As clsProject
    '    Dim objEnv As clsEnvironment
    '    Dim objStr As clsStructure
    '    Dim objSel As clsStructureSelection

    '    Try
    '        Select Case obj.Type
    '            Case NODE_PROJECT
    '                objProj = CType(obj, clsProject)
    '                sql = "Update " & objProj.Project.tblStrSelFields & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

    '            Case NODE_ENVIRONMENT
    '                objEnv = CType(obj, clsEnvironment)
    '                sql = "Update " & objEnv.Project.tblStrSelFields & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

    '            Case NODE_STRUCT
    '                objStr = CType(obj, clsStructure)
    '                sql = "Update " & objStr.Project.tblStrSelFields & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

    '            Case NODE_STRUCT_SEL
    '                objSel = CType(obj, clsStructureSelection)
    '                sql = "Update " & objSel.Project.tblStrSelFields & " Set SelectionName= '" & NewValue & "' WHERE SelectionName= '" & OldValue & "' AND ProjectName='" & objSel.Project.ProjectName & "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

    '            Case Else
    '                EditSTRSELFIELDS = False
    '                Exit Function
    '        End Select

    '        cmd.CommandText = sql
    '        Log(sql)
    '        cmd.ExecuteNonQuery()

    '        EditSTRSELFIELDS = True

    '    Catch ex As Exception
    '        LogError(ex, "modRename EditStrSelFields", sql)
    '        EditSTRSELFIELDS = False
    '    End Try

    'End Function

    'done
    'Function EditStructureSelections(ByRef cmd As Data.Odbc.OdbcCommand, ByVal OldValue As String, ByVal NewValue As String, ByVal obj As INode) As Boolean

    '    Dim sql As String = ""
    '    Dim objProj As clsProject
    '    Dim objEnv As clsEnvironment
    '    Dim objStr As clsStructure
    '    Dim objSel As clsStructureSelection

    '    Try
    '        Select Case obj.Type
    '            Case NODE_PROJECT
    '                objProj = CType(obj, clsProject)
    '                sql = "Update " & objProj.Project.tblStructSel & " Set ProjectName= '" & NewValue & "' WHERE ProjectName= '" & OldValue & "'"

    '            Case NODE_ENVIRONMENT
    '                objEnv = CType(obj, clsEnvironment)
    '                sql = "Update " & objEnv.Project.tblStructSel & " Set EnvironmentName= '" & NewValue & "' WHERE EnvironmentName= '" & OldValue & "' AND ProjectName='" & objEnv.Project.ProjectName & "'"

    '            Case NODE_STRUCT
    '                objStr = CType(obj, clsStructure)
    '                sql = "Update " & objStr.Project.tblStructSel & " Set StructureName= '" & NewValue & "',SelectionName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "' AND ISSYSTEMSELECT= '1'"

    '                cmd.CommandText = sql
    '                Log(sql)
    '                cmd.ExecuteNonQuery()

    '                sql = "Update " & objStr.Project.tblStructSel & " Set StructureName= '" & NewValue & "' WHERE StructureName= '" & OldValue & "' AND ProjectName='" & objStr.Project.ProjectName & "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "' AND ISSYSTEMSELECT= '0'"

    '            Case NODE_STRUCT_SEL
    '                objSel = CType(obj, clsStructureSelection)
    '                sql = "Update " & objSel.Project.tblStructSel & " Set SelectionName= '" & NewValue & "' WHERE SelectionName= '" & OldValue & "' AND ProjectName= '" & objSel.Project.ProjectName & "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "' AND StructureName= '" & objSel.ObjStructure.StructureName & "'"

    '            Case Else
    '                EditStructureSelections = False
    '                Exit Function
    '        End Select

    '        cmd.CommandText = sql
    '        Log(sql)
    '        cmd.ExecuteNonQuery()

    '        EditStructureSelections = True

    '    Catch ex As Exception
    '        LogError(ex, "modRename EditStructureSelections", sql)
    '        EditStructureSelections = False
    '    End Try

    'End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDESCFIELDS", sql)
            EditDESCFIELDS = False
        Catch ex As Exception
            LogError(ex, "modRename EditDESCFIELDS", sql)
            EditDESCFIELDS = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDescriptions", sql)
            EditDescriptions = False
        Catch ex As Exception
            LogError(ex, "modRename EditDescriptions", sql)
            EditDescriptions = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDescriptionsATTR", sql)
            EditDescriptionsATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditDescriptionsATTR", sql)
            EditDescriptionsATTR = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDESCSELFIELDS", sql)
            EditDESCSELFIELDS = False
        Catch ex As Exception
            LogError(ex, "modRename EditDESCSELFIELDS", sql)
            EditDESCSELFIELDS = False
        End Try

    End Function

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
                    sql = "Update " & objSel.Project.tblDescriptionSelect & _
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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditDESCSelections", sql)
            EditDESCSelections = False
        Catch ex As Exception
            LogError(ex, "modRename EditDESCSelections", sql)
            EditDESCSelections = False
        End Try

    End Function

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

            EditSystems = EditSystemsATTR(cmd, OldValue, NewValue, obj)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    EditSystems = 
            'End If

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditSystems", sql)
            EditSystems = False
        Catch ex As Exception
            LogError(ex, "modRename EditSystems", sql)
            EditSystems = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditSystemsATTR", sql)
            EditSystemsATTR = False
        Catch ex As Exception
            LogError(ex, "modRename EditSystemsATTR", sql)
            EditSystemsATTR = False
        End Try

    End Function

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
                    'If objDS.Engine IsNot Nothing Then
                    sql = "Update " & objDS.Project.tblTaskDS & _
                                        " Set DatastoreName= '" & NewValue & _
                                        "' WHERE DatastoreName= '" & OldValue & _
                                        "' AND ProjectName='" & objDS.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objDS.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objDS.Engine.EngineName & "'"
                    'Else
                    'sql = "Update " & objDS.Project.tblTaskDS & _
                    '                    " Set DatastoreName= '" & NewValue & _
                    '                    "' WHERE DatastoreName= '" & OldValue & _
                    '                    "' AND ProjectName='" & objDS.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objDS.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                    objTask = CType(obj, clsTask)
                    'If objTask.Engine IsNot Nothing Then
                    sql = "Update " & objTask.Project.tblTaskDS & _
                                        " Set TaskName= '" & NewValue & _
                                        "' WHERE TaskName= '" & OldValue & _
                                        "' AND ProjectName='" & objTask.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objTask.Engine.EngineName & _
                                        "' AND TASKTYPE= " & objTask.TaskType
                    'Else
                    'sql = "Update " & objTask.Project.tblTaskDS & _
                    '                    " Set TaskName= '" & NewValue & _
                    '                    "' WHERE TaskName= '" & OldValue & _
                    '                    "' AND ProjectName='" & objTask.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    'End If

                Case Else
                    EditTaskDatastores = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTaskDatastores = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditTaskDatastores", sql)
            EditTaskDatastores = False
        Catch ex As Exception
            LogError(ex, "modRename EditTaskDatastores", sql)
            EditTaskDatastores = False
        End Try

    End Function

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

                    If EditMappingSrc(cmd, OldValue, NewValue, objTask) = True Then
                        'If objTask.Engine IsNot Nothing Then
                        sql = "Update " & objTask.Project.tblTaskMap & _
                                           " Set TaskName= '" & NewValue & _
                                           "' WHERE TaskName= '" & OldValue & _
                                           "' AND ProjectName='" & objTask.Project.ProjectName & _
                                           "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                           "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                           "' AND EngineName= '" & objTask.Engine.EngineName & _
                                           "' AND TASKTYPE= " & objTask.TaskType
                        'Else
                        'sql = "Update " & objTask.Project.tblTaskMap & _
                        '                   " Set TaskName= '" & NewValue & _
                        '                   "' WHERE TaskName= '" & OldValue & _
                        '                   "' AND ProjectName='" & objTask.Project.ProjectName & _
                        '                   "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                        'End If
                    End If

                Case NODE_SOURCEDATASTORE
                    objSDS = CType(obj, clsDatastore)

                    If EditMappingSrc(cmd, OldValue, NewValue, objSDS) = True Then
                        sql = "Update " & objSDS.Project.tblTaskMap & _
                                            " Set SourceDatastore= '" & NewValue & _
                                            "' WHERE SourceDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objSDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objSDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objSDS.Engine.EngineName & "'"
                    End If

                    'If objSDS.Engine IsNot Nothing Then

                    'Else
                    '    sql = "Update " & objSDS.Project.tblTaskMap & _
                    '                        " Set SourceDatastore= '" & NewValue & _
                    '                        "' WHERE SourceDatastore= '" & OldValue & _
                    '                        "' AND ProjectName='" & objSDS.Project.ProjectName & _
                    '                        "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_TARGETDATASTORE
                    objTDS = CType(obj, clsDatastore)

                    If EditMappingSrc(cmd, OldValue, NewValue, objTDS) = True Then
                        sql = "Update " & objTDS.Project.tblTaskMap & _
                                            " Set TargetDatastore= '" & NewValue & _
                                            "' WHERE TargetDatastore= '" & OldValue & _
                                            "' AND ProjectName='" & objTDS.Project.ProjectName & _
                                            "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & _
                                            "' AND SystemName= '" & objTDS.Engine.ObjSystem.SystemName & _
                                            "' AND EngineName= '" & objTDS.Engine.EngineName & "'"
                    End If

                    'If objTDS.Engine IsNot Nothing Then

                    'Else
                    '    sql = "Update " & objTDS.Project.tblTaskMap & _
                    '                        " Set TargetDatastore= '" & NewValue & _
                    '                        "' WHERE TargetDatastore= '" & OldValue & _
                    '                        "' AND ProjectName='" & objTDS.Project.ProjectName & _
                    '                        "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)

                    If EditMappingSrc(cmd, OldValue, NewValue, objStr) = True Then
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
                    End If

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)

                    EditTaskMappings = EditMappingSrc(cmd, OldValue, NewValue, objSel)
                    Exit Try

                Case NODE_VARIABLE
                    objVar = CType(obj, clsVariable)

                    EditTaskMappings = EditMappingSrc(cmd, OldValue, NewValue, objVar)
                    Exit Try

                Case Else
                    EditTaskMappings = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTaskMappings = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditTaskMappings", sql)
            EditTaskMappings = False
        Catch ex As Exception
            LogError(ex, "modRename EditTaskMappings", sql)
            EditTaskMappings = False
        End Try

    End Function

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
                    'If objTask.Engine IsNot Nothing Then
                    sql = "Update " & objTask.Project.tblTasks & _
                                        " Set TaskName= '" & NewValue & _
                                        "' WHERE TaskName= '" & OldValue & _
                                        "' AND ProjectName='" & objTask.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objTask.Engine.EngineName & _
                                        "' AND TASKTYPE= " & objTask.TaskType
                    'Else
                    'sql = "Update " & objTask.Project.tblTasks & _
                    '                    " Set TaskName= '" & NewValue & _
                    '                    "' WHERE TaskName= '" & OldValue & _
                    '                    "' AND ProjectName='" & objTask.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    'End If

                Case Else
                    EditTasks = False
                    Exit Function
            End Select

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            EditTasks = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditTasks", sql)
            EditTasks = False
        Catch ex As Exception
            LogError(ex, "modRename EditTasks", sql)
            EditTasks = False
        End Try

    End Function

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

            EditVariables = EditVariablesATTR(cmd, OldValue, NewValue, obj)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            '    EditVariables = 
            'End If

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditVariables", sql)
            EditVariables = False
        Catch ex As Exception
            LogError(ex, "modRename EditVariables", sql)
            EditVariables = False
        End Try

    End Function

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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditVariablesATTR", sql)
            EditVariablesATTR = False
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
            If loopRange > 50 Then
                ValidateNewName = False
                MsgBox("Name cannot be more than 50 characters," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                ValidateNewName = False
                Exit Function
            End If

            If NewName = "" Then
                ValidateNewName = False
                MsgBox("Name cannot be left blank," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                ValidateNewName = False
                Exit Function
            End If

            If NewName.StartsWith("%") = True And IsEng = True Then
                If NewName.Length <> 3 Then
                    ValidateNewName = False
                    MsgBox("Substitution Variables Must be in the format '%XX'," & Chr(13) & _
                    "where each X is a digit from 0 thru 9." & Chr(13) & _
                    "New name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                    ValidateNewName = False
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
                            ValidateNewName = False
                            Exit Function
                        End If
                    Else
                        If Not ((asciiInt >= 97 And asciiInt <= 122) Or (asciiInt >= 64 And asciiInt <= 90) Or (asciiInt >= 48 And asciiInt <= 57) Or (asciiInt = 38) Or (asciiInt = 35) Or (asciiInt = 45) Or (asciiInt = 95)) Then
                            ValidateNewName = False
                            MsgBox("Name must only consist of alpha-numeric characters, &, @, -, _, or #." & Chr(13) & "New Name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                            ValidateNewName = False
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
            ValidateNewName = False
        End Try

    End Function

    Function ValidateNewName128(ByVal NewName As String) As Boolean

        Dim Letter As Char  '// character of new name presently being analized
        Dim i As Integer   '// loop counter
        Dim loopRange As Integer = Len(NewName.Trim)  '// length of new name
        Dim asciiInt As Integer     '// the ascii code for the character being analized
        Dim FirstLetterFlag As Boolean = True  '// determines if the character is the first letter in the new name

        Try
            NewName = NewName.Trim
            If loopRange > 128 Then
                ValidateNewName128 = False
                MsgBox("Name cannot be more than 128 characters," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                Exit Function
            End If

            If NewName = "" Then
                ValidateNewName128 = False
                MsgBox("Name cannot be left blank," & (Chr(13)) & "Please enter a name", MsgBoxStyle.Information, MsgTitle)
                Exit Function
            End If

            
            'loopRange = Len(NewName)
            For i = 1 To loopRange
                Letter = Left(NewName, 1)
                asciiInt = Asc(Letter)
                If FirstLetterFlag = True Then
                    If Not ((asciiInt >= 97 And asciiInt <= 122) Or (asciiInt >= 64 And asciiInt <= 90) Or (asciiInt = 38)) Then
                        ValidateNewName128 = False
                        MsgBox("Name must start with an alphabetic character, &, or @." & Chr(13) & "New name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                        Exit Function
                    End If
                Else
                    If Not ((asciiInt >= 97 And asciiInt <= 122) Or (asciiInt >= 64 And asciiInt <= 90) Or (asciiInt >= 48 And asciiInt <= 57) Or (asciiInt = 38) Or (asciiInt = 35) Or (asciiInt = 45) Or (asciiInt = 95)) Then
                        ValidateNewName128 = False
                        MsgBox("Name must only consist of alpha-numeric characters, &, @, -, _, or #." & Chr(13) & "New Name will not be saved", MsgBoxStyle.Information, "Invalid Name")
                        Exit Function
                    End If
                End If
                NewName = Right(NewName, Len(NewName) - 1)
                FirstLetterFlag = False
            Next

            ValidateNewName128 = True

        Catch ex As Exception
            LogError(ex, "modRename ValidateNewName128")
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
                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    ParName = dr.Item("Parent")
                '    DSDir = dr.Item("DsDirection")
                '    StrName = dr.Item("StructureName")
                'Else
                ParName = dr.Item("ParentName")
                StrName = dr.Item("DescriptionName")
                'End If


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
                    'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                    '    DSDirArray.Add(DSDir)
                    'End If
                End If

            End While
            dr.Close()  '// must be closed when done so that the rest of the "non-query" part of the 
            '// transaction can continue

            '// Now that we have made new foreign keys using the new name
            '// we will update the DSselFields Table with the new Fkeys where all the other fields read in
            '// match, so that we put the right fKey in the right selection field
            For i = 0 To FkeyArray.Count - 1
                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    sql = "Update " & obj.Project.tblDSselFields & _
                '    " Set ForeignKey= '" & FkeyArray(i) & _
                '    "' WHERE FieldName= '" & FldNameArray(i) & _
                '    "' AND DatastoreName= '" & DSNameArray(i) & _
                '    "' AND EngineName= '" & EngNameArray(i) & _
                '    "' AND SystemName= '" & SysNameArray(i) & _
                '    "' AND SelectionName= '" & SelNameArray(i) & _
                '    "' AND StructureName= '" & StrNameArray(i) & _
                '    "' AND EnvironmentName= '" & EnvNameArray(i) & _
                '    "' AND ProjectName= '" & ProjNameArray(i) & _
                '    "' AND Parent= '" & ParNameArray(i) & _
                '    "' AND DsDirection= '" & DSDirArray(i) & "'"
                'Else
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
                'End If


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            EditFkey = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditFkey", sql)
            EditFkey = False
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
        Dim dt As New DataTable("temp")

        Dim objSDS As clsDatastore
        Dim objTDS As clsDatastore
        Dim objStr As clsStructure
        Dim objSel As clsStructureSelection
        Dim objTask As clsTask
        Dim objVar As clsVariable
        Dim sql As String = ""
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
                    'If objTask.Engine IsNot Nothing Then
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " _
                                       & objTask.Project.tblTaskMap & _
                                       " WHERE ProjectName='" & objTask.Project.ProjectName & _
                                       "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & _
                                       "' AND SystemName= '" & objTask.Engine.ObjSystem.SystemName & _
                                       "' AND EngineName= '" & objTask.Engine.EngineName & "'"

                    'Else
                    'sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " _
                    '                   & objTask.Project.tblTaskMap & _
                    '                   " WHERE ProjectName='" & objTask.Project.ProjectName & _
                    '                   "' AND EnvironmentName= '" & objTask.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_SOURCEDATASTORE
                    objSDS = CType(obj, clsDatastore)
                    'If objSDS.Engine IsNot Nothing Then
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                        objSDS.Project.tblTaskMap & _
                                        " WHERE ProjectName='" & objSDS.Project.ProjectName & _
                                        "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & _
                                        "' AND SystemName= '" & objSDS.Engine.ObjSystem.SystemName & _
                                        "' AND EngineName= '" & objSDS.Engine.EngineName & "'"
                    'Else
                    'sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    '                    objSDS.Project.tblTaskMap & _
                    '                    " WHERE ProjectName='" & objSDS.Project.ProjectName & _
                    '                    "' AND EnvironmentName= '" & objSDS.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_TARGETDATASTORE
                    objTDS = CType(obj, clsDatastore)
                    'If objTDS.Engine IsNot Nothing Then
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                                       objTDS.Project.tblTaskMap & _
                                       " WHERE ProjectName='" & objTDS.Project.ProjectName & _
                                       "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & _
                                       "' AND SystemName= '" & objTDS.Engine.ObjSystem.SystemName & _
                                       "' AND EngineName= '" & objTDS.Engine.EngineName & "'"
                    'Else
                    'sql = "SELECT MAPPINGSOURCE, SEQNO, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    '                   objTDS.Project.tblTaskMap & _
                    '                   " WHERE ProjectName='" & objTDS.Project.ProjectName & _
                    '                   "' AND EnvironmentName= '" & objTDS.Environment.EnvironmentName & "'"
                    'End If

                Case NODE_STRUCT
                    objStr = CType(obj, clsStructure)
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    objStr.Project.tblTaskMap & _
                    " WHERE ProjectName='" & objStr.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objStr.Environment.EnvironmentName & "'"

                Case NODE_STRUCT_SEL
                    objSel = CType(obj, clsStructureSelection)
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
                    objSel.Project.tblTaskMap & _
                    " WHERE ProjectName='" & objSel.Project.ProjectName & _
                    "' AND EnvironmentName= '" & objSel.ObjStructure.Environment.EnvironmentName & "'"

                Case NODE_VARIABLE
                    objVar = CType(obj, clsVariable)
                    sql = "SELECT MAPPINGSOURCE, MAPPINGID, TASKNAME, ENGINENAME, SYSTEMNAME, ENVIRONMENTNAME, PROJECTNAME FROM " & _
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
                MapString = GetVal(dr.Item("MAPPINGSOURCE"))
                SeqNo = GetVal(dr.Item("MAPPINGID"))
                TaskName = GetVal(dr.Item("TASKNAME"))
                EngineName = GetVal(dr.Item("ENGINENAME"))
                SystemName = GetVal(dr.Item("SYSTEMNAME"))
                EnvName = GetVal(dr.Item("ENVIRONMENTNAME"))
                ProjectName = GetVal(dr.Item("PROJECTNAME"))

                Dim NewMapSrc As New System.Text.StringBuilder(MapString)
                Dim MapSource As New System.Text.StringBuilder(MapString)

                Select Case obj.Type
                    Case NODE_LOOKUP, NODE_GEN, NODE_PROC, NODE_MAIN
                        NewMapSrc.Replace("(" & OldValue & ")", "(" & NewValue & ")")
                        'NewMapSrc.Replace("(" & OldValue & ".", "(" & NewValue & ".")
                        NewMapSrc.Replace("(" & OldValue & " ", "(" & NewValue & " ")
                        NewMapSrc.Replace("(" & OldValue & ",", "(" & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & ")", " " & NewValue & ")")
                        NewMapSrc.Replace(" " & OldValue & ",", " " & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & " ", " " & NewValue & " ")
                        NewMapSrc.Replace("," & OldValue & ")", "," & NewValue & ")")
                        NewMapSrc.Replace("," & OldValue & " ", "," & NewValue & " ")

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

                    Case NODE_STRUCT, NODE_STRUCT_SEL

                        NewMapSrc.Replace("'" & OldValue & "'", "'" & NewValue & "'")
                        NewMapSrc.Replace("." & OldValue & ".", "." & NewValue & ".")
                        NewMapSrc.Replace("(" & OldValue & " ", "(" & NewValue & " ")
                        NewMapSrc.Replace("(" & OldValue & ",", "(" & NewValue & ",")
                        NewMapSrc.Replace("(" & OldValue & ".", "(" & NewValue & ".")
                        NewMapSrc.Replace(" " & OldValue & ")", " " & NewValue & ")")
                        NewMapSrc.Replace(" " & OldValue & ",", " " & NewValue & ",")
                        NewMapSrc.Replace(" " & OldValue & ".", " " & NewValue & ".")
                        NewMapSrc.Replace("," & OldValue & ")", "," & NewValue & ")")
                        NewMapSrc.Replace("," & OldValue & " ", "," & NewValue & " ")
                        NewMapSrc.Replace("," & OldValue & ".", "," & NewValue & ".")

                    Case NODE_VARIABLE

                        NewMapSrc.Replace(OldValue, NewValue)

                End Select

                '// if nothing was replaced it means that old value didn't exist in that 
                '// mappingSource field, so ignore it and go to the next match
                '// if the old Mapping Source Field doesn't match the New "replaced" Mapping Source
                '// field, then we will need to update it, so save all of the row data to 
                '// arraylists so they can be replaced one at a time in the loop below
                If NewMapSrc.ToString <> MapSource.ToString Then

                    SeqNoArray.Add(SeqNo)
                    MapSourceArray.Add(NewMapSrc.ToString)
                    TaskNameArray.Add(TaskName)
                    EngineNameArray.Add(EngineName)
                    SystemNameArray.Add(SystemName)
                    EnvironmentNameArray.Add(EnvName)
                    ProjectNameArray.Add(ProjectName)
                End If
            End While

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditMappingSrc", sql)
            EditMappingSrc = False
            Exit Function
        Catch ex As Exception
            LogError(ex, "modRename EditMappingSrc", sql)
            EditMappingSrc = False
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
                "' WHERE MAPPINGID= " & CInt(SeqNoArray(i)) & _
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

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename EditMappingSrc", sql)
            EditMappingSrc = False
        Catch ex As Exception
            LogError(ex, "modRename EditMappingSrc", sql)
            EditMappingSrc = False
        End Try

        '/// sample of using Regular expressions to solve this problem
        'Dim RegExp1 As New System.Text.RegularExpressions.Regex("(^|[^-.0-9A-Za-z_])" & OldValue & "([^-_0-9A-Za-z]|$)")
        'Dim RegExp2 As New System.Text.RegularExpressions.Regex(MapString)

    End Function

#End Region

    '// functions in this region are called from the frmStructure form 
    '// when replacing structure files ... by Tom Karasch April and June 2007
    '// *** Rewritten: October 2010 by TK
#Region "Replace structure file"

    '/// added by TKarasch April 07 to replace Structure file ... 
    '// Main replacement function ... Needs refining
    '**********************************************************
    '/// Refined 1/09 to 4/09 by TKarasch
    Function ReplaceStrFile(ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure) As Boolean

        '// return value from "compareOldNewFldNames" function to determine status for replacement
        Dim IntRet As Integer
        '// Arraylists to get populated by CompareOldNewFldNames function
        '***** Array of New Fields in Incoming description that don't exist in the existing description
        Dim AddedFieldList As New ArrayList
        '***** Array of Old Fields in current description that don't exist in the existing description
        Dim DeletedFieldList As New ArrayList

        'Dim Engs As New Collection

        AddedFieldList.Clear()
        DeletedFieldList.Clear()
        Procs.Clear()
        'Engs.Clear()

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
            Replace = frm.ShowFldDiff(NewStrObj, OldStrObj, AddedFieldList, DeletedFieldList, Procs)

            If Replace = True Then
                ReplaceStrFile = SQLReplaceFile(NewStrObj, OldStrObj, AddedFieldList, DeletedFieldList, Procs)
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
    Function SQLReplaceFile(ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure, ByVal addList As ArrayList, ByVal delList As ArrayList, ByRef procedures As Collection) As Boolean

        '// All Comparison testing of fields has been done before we open Metadata
        '// so we are sure that Updating of metadata will work without errors
        '// and all deletes and adds will happen in one Transaction, so rollback is possible
        Dim cmd As New Odbc.OdbcCommand
        Dim tran As Odbc.OdbcTransaction = Nothing
        Dim cascade As Boolean = False
        Dim success As Boolean = True

        Try
            '// start SQL Tran
            cmd.Connection = cnn
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran


            '// Now delete fields from metadata that are no longer present in 
            '// If either delete or add operations fail, then ROLLBACK transaction
            If OldStrObj.DeleteATTR(cmd) = True Then
                If NewStrObj.InsertATTR(cmd) = True Then
                    If ChangeStrFieldsForReplace(cmd, NewStrObj, OldStrObj, addList, delList) = True Then
                        If UpdateDSSelFldsForStrChng(cmd, NewStrObj, OldStrObj, addList, delList) = True Then
                            If UpdateMappings(cmd, NewStrObj, OldStrObj, procedures, addList, delList) = True Then
                                '// All went well, then commit
                                tran.Commit()
                                success = True
                            Else
                                tran.Rollback()
                                success = False
                            End If
                        Else
                            tran.Rollback()
                            success = False
                        End If
                    Else
                        tran.Rollback()
                        success = False
                    End If
                Else
                    tran.Rollback()
                    success = False
                End If
            Else
                tran.Rollback()
                success = False
            End If


            SQLReplaceFile = success

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename SQLReplaceFile")
            tran.Rollback()
            SQLReplaceFile = False
        Catch ex As Exception
            LogError(ex, "modRename SQLReplaceFile")
            tran.Rollback()
            SQLReplaceFile = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Function ChangeStrFieldsForReplace(ByRef cmd As Odbc.OdbcCommand, ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure, ByVal addList As ArrayList, ByVal delList As ArrayList) As Boolean

        Try
            Dim sql As String = ""
            '/// First Delete fields no longer present in new structure and retain their seqnos
            '/// Table DescFields and descselflds
            '/DescFlds
            mainwindow.StatusBar1.Text = "Deleting fields no longer in Description File..."
            For Each fld As clsField In delList
                sql = "DELETE FROM " & NewStrObj.Project.tblDescriptionFields & " WHERE" & _
                " PROJECTNAME = " & NewStrObj.Project.GetQuotedText & _
                " AND ENVIRONMENTNAME= " & NewStrObj.Environment.GetQuotedText & _
                " AND DESCRIPTIONNAME= " & NewStrObj.GetQuotedText & _
                " AND FIELDNAME= " & fld.GetQuotedText

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next
            '/DescSelFlds
            For Each fld As clsField In delList
                sql = "DELETE FROM " & NewStrObj.Project.tblDescriptionSelFields & " WHERE" & _
                " PROJECTNAME = " & NewStrObj.Project.GetQuotedText & _
                " AND ENVIRONMENTNAME= " & NewStrObj.Environment.GetQuotedText & _
                " AND DESCRIPTIONNAME= " & NewStrObj.GetQuotedText & _
                " AND FIELDNAME= " & fld.GetQuotedText

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            '/// Next Add New Structure Fields into metadata for new structure based on seqno
            '/// Table DescFields
            mainwindow.StatusBar1.Text = "Inserting new fields added by New Description File..."
            For Each fld As clsField In addList
                sql = "INSERT INTO " & NewStrObj.Project.tblDescriptionFields & " (" & _
                    "ProjectName, " & _
                    "EnvironmentName, " & _
                    "DescriptionName, " & _
                    "FieldName, " & _
                    "ParentName, " & _
                    "SEQNO, " & _
                    "DescFieldDescription, " & _
                    "NChildren, " & _
                    "NLevel, " & _
                    "Ntimes, " & _
                    "NOccNo, " & _
                    "Datatype, " & _
                    "NOffSet, " & _
                    "NLength, " & _
                    "NScale, " & _
                    "CanNull, " & _
                    "ISKEY, " & _
                    "OrgName, " & _
                    "DateFormat, " & _
                    "Label, " & _
                    "InitVal, " & _
                    "ReType, " & _
                    "InValid, " & _
                    "ExtType, " & _
                    "IdentVal, " & _
                    "ForeignKey) " & _
                    "Values(" & _
                    NewStrObj.Project.GetQuotedText & "," & _
                    NewStrObj.Environment.GetQuotedText & "," & _
                    NewStrObj.GetQuotedText & "," & _
                    Quote(fld.FieldName) & "," & _
                    Quote(fld.ParentName) & "," & _
                    fld.SeqNo & "," & _
                    Quote(fld.FieldDesc) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                    fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                    Quote(fld.OrgName) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                    Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                    ");"

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            '/// Update all fields to reflect any changes in attributes
            '/// table DescFields
            mainwindow.StatusBar1.Text = "Updating all field attributes in New Description File..."
            For Each fld As clsField In NewStrObj.ObjFields
                sql = "UPDATE " & NewStrObj.Project.tblDescriptionFields & " SET " & _
                "DescFieldDescription= " & Quote(fld.FieldDesc) & "," & _
                "NChildren= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                "NLevel= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                "Ntimes= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                "NOccNo= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                "Datatype= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                "NOffSet= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                "NLength= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                "NScale= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                "CanNull= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                "ISKEY= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                "OrgName= " & Quote(fld.OrgName) & "," & _
                "DateFormat= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                "Label= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                "InitVal= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                "ReType= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                "InValid= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                "ExtType= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                "IdentVal= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                "ForeignKey= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                " WHERE " & _
                "ProjectName= " & NewStrObj.Project.GetQuotedText & _
                " AND EnvironmentName= " & NewStrObj.Environment.GetQuotedText & _
                " AND DescriptionName= " & NewStrObj.GetQuotedText & _
                " AND FieldName= " & Quote(fld.FieldName) & _
                " AND ParentName= " & Quote(fld.ParentName) & _
                " AND SEQNO= " & fld.SeqNo

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            ChangeStrFieldsForReplace = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename ChangeStrFieldsForReplace")
            ChangeStrFieldsForReplace = False
        Catch ex As Exception
            LogError(ex, "modRename ChangeStrFieldsForReplace")
            ChangeStrFieldsForReplace = False
        End Try

    End Function

    Function UpdateDSSelFldsForStrChng(ByRef cmd As Odbc.OdbcCommand, ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure, ByVal addList As ArrayList, ByVal delList As ArrayList) As Boolean

        Try
            Dim sql As String = ""

            OldStrObj.SysAllSelection.ObjDSselections = SearchSSForDSSelRef(OldStrObj.SysAllSelection)

            '/// For each DSSel affected, delete fields that no longer exist in new structure and retain seqnos
            '/// Table DSSelFlds
            mainwindow.StatusBar1.Text = "Deleting fields from Datastore Selections not in New Description File..."
            For Each DSSel As clsDSSelection In OldStrObj.SysAllSelection.ObjDSselections
                For Each fld As clsField In delList
                    sql = "DELETE FROM " & NewStrObj.Project.tblDSselFields & " WHERE " & _
                    "PROJECTNAME = " & DSSel.Project.GetQuotedText & _
                    " AND ENVIRONMENTNAME= " & DSSel.Environment.GetQuotedText & _
                    " AND SYSTEMNAME= " & DSSel.Engine.ObjSystem.GetQuotedText & _
                    " AND ENGINENAME= " & DSSel.Engine.GetQuotedText & _
                    " AND DATASTORENAME= " & DSSel.ObjDatastore.GetQuotedText & _
                    " AND DESCRIPTIONNAME= " & NewStrObj.GetQuotedText & _
                    " AND FIELDNAME= " & fld.GetQuotedText

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next
            Next

            '/// Next Add New DSSel Fields for new structure based on seqnos
            '/// Table DSSelFlds
            mainwindow.StatusBar1.Text = "Inserting all new fields into Datastore Selections..."
            For Each DSSel As clsDSSelection In OldStrObj.SysAllSelection.ObjDSselections
                For Each objfld As clsField In addList
                    sql = "INSERT INTO " & NewStrObj.Project.tblDSselFields & " (" & _
                            "ProjectName," & _
                            "EnvironmentName," & _
                            "SystemName," & _
                            "EngineName," & _
                            "DatastoreName," & _
                            "DescriptionName, " & _
                            "SelectionName, " & _
                            "FieldName, " & _
                            "ParentName, " & _
                            "SEQNO, " & _
                            "DescFieldDescription, " & _
                            "NChildren, " & _
                            "NLevel, " & _
                            "Ntimes, " & _
                            "NOccNo, " & _
                            "Datatype, " & _
                            "NOffSet, " & _
                            "NLength, " & _
                            "NScale, " & _
                            "CanNull, " & _
                            "ISKEY, " & _
                            "OrgName, " & _
                            "DateFormat, " & _
                            "Label, " & _
                            "InitVal, " & _
                            "ReType, " & _
                            "InValid, " & _
                            "ExtType, " & _
                            "IdentVal, " & _
                            "ForeignKey) " & _
                            "Values( " & _
                            DSSel.Project.GetQuotedText & "," & _
                            DSSel.Engine.ObjSystem.Environment.GetQuotedText & "," & _
                            DSSel.Engine.ObjSystem.GetQuotedText & "," & _
                            DSSel.Engine.GetQuotedText & "," & _
                            DSSel.ObjDatastore.GetQuotedText & "," & _
                            DSSel.ObjStructure.GetQuotedText & "," & _
                            DSSel.GetQuotedText & "," & _
                            objfld.GetQuotedText & "," & _
                            Quote(objfld.ParentName) & "," & _
                            objfld.SeqNo & "," & _
                            Quote(objfld.FieldDesc) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                            objfld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_ISKEY)) & "," & _
                            Quote(objfld.OrgName) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_DATEFORMAT)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_LABEL)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_INITVAL)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_RETYPE)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_INVALID)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_EXTTYPE)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_IDENTVAL)) & "," & _
                            Quote(objfld.GetFieldAttr(enumFieldAttributes.ATTR_FKEY)) & _
                            ");" '//TK 11/8/06 and 4/5/07 and 2/15/09

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next
            Next

            '/// Update all fields to reflect any changes in attributes from new Description File
            '/// Table DSSelFlds
            mainwindow.StatusBar1.Text = "Updating all field attributes in Datastores for new Description File..."
            For Each DSSel As clsDSSelection In OldStrObj.SysAllSelection.ObjDSselections
                For Each fld As clsField In NewStrObj.ObjFields
                    sql = "UPDATE " & DSSel.Project.tblDSselFields & " SET " & _
                    "NChildren= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_NCHILDREN) & "," & _
                    "NLevel= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) & "," & _
                    "Ntimes= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_TIMES) & "," & _
                    "NOccNo= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_OCCURS) & "," & _
                    "Datatype= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_DATATYPE)) & "," & _
                    "NOffSet= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_OFFSET) & "," & _
                    "NLength= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_LENGTH) & "," & _
                    "NScale= " & fld.GetFieldAttr(enumFieldAttributes.ATTR_SCALE) & "," & _
                    "CanNull= " & Quote(fld.GetFieldAttr(enumFieldAttributes.ATTR_CANNULL)) & _
                    " WHERE " & _
                    "ProjectName= " & DSSel.Project.GetQuotedText & _
                    " AND EnvironmentName= " & DSSel.Environment.GetQuotedText & _
                    " AND SYSTEMNAME= " & DSSel.Engine.ObjSystem.GetQuotedText & _
                    " AND ENGINENAME= " & DSSel.Engine.GetQuotedText & _
                    " AND DATASTORENAME= " & DSSel.ObjDatastore.GetQuotedText & _
                    " AND DescriptionName= " & DSSel.GetQuotedText & _
                    " AND FieldName= " & Quote(fld.FieldName) '& _
                    '" AND ParentName= " & Quote(fld.ParentName) '& _
                    '" AND SEQNO= " & fld.SeqNo

                    cmd.CommandText = sql
                    Log(sql)
                    cmd.ExecuteNonQuery()
                Next
            Next

            UpdateDSSelFldsForStrChng = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename UpdateDSSelFldsForStrChng")
            UpdateDSSelFldsForStrChng = False
        Catch ex As Exception
            LogError(ex, "modRename UpdateDSSelFldsForStrChng")
            UpdateDSSelFldsForStrChng = False
        End Try

    End Function

    '/// function used when replacing structure files.
    '// this function updates the mapping source and target fields
    '// that are in mappings that use the replaced structure 
    Function UpdateMappings(ByRef cmd As Odbc.OdbcCommand, ByRef NewStrObj As clsStructure, ByRef OldStrObj As clsStructure, ByRef procedures As Collection, ByVal addList As ArrayList, ByVal delList As ArrayList) As Boolean

        Try
            Dim sql As String = ""
            Dim Obj As Object
            Dim success As Boolean = True

            '/// For each procedure in the procedure section, delete mapping sources or targets for fields in the delList
            '/// Object manipulation
            mainwindow.StatusBar1.Text = "Updating all Mappings to reflect new Description File..."
            For Each task As clsTask In procedures
                For Each map As clsMapping In task.ObjMappings
                    Obj = map.MappingSource
                    If Obj IsNot Nothing Then
                        For Each fld As clsField In delList
                            If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                If CType(Obj, clsField).FieldName = fld.FieldName Then
                                    map.MappingSource = Nothing
                                    map.SourceType = enumMappingType.MAPPING_TYPE_NONE
                                    map.SourceDataStore = ""
                                    map.SourceParent = ""
                                    map.IsModified = True
                                End If
                            End If
                            If CType(Obj, INode).Type = NODE_FUN Then
                                If CType(Obj, clsSQFunction).SQFunctionWithInnerText.Contains(fld.FieldName) Then
                                    map.MappingSource = Nothing
                                    map.SourceType = enumMappingType.MAPPING_TYPE_NONE
                                    map.SourceDataStore = ""
                                    map.SourceParent = ""
                                    map.IsModified = True
                                End If
                            End If
                        Next
                    End If
                    Obj = map.MappingTarget
                    If Obj IsNot Nothing Then
                        For Each fld As clsField In delList
                            If CType(Obj, INode).Type = NODE_STRUCT_FLD Then
                                If CType(Obj, clsField).FieldName = fld.FieldName Then
                                    map.MappingTarget = Nothing
                                    map.TargetType = enumMappingType.MAPPING_TYPE_NONE
                                    map.TargetDataStore = ""
                                    map.TargetParent = ""
                                    map.IsModified = True
                                End If
                            End If
                            If CType(Obj, INode).Type = NODE_FUN Then
                                If CType(Obj, clsSQFunction).SQFunctionWithInnerText.Contains(fld.FieldName) Then
                                    map.MappingTarget = Nothing
                                    map.TargetType = enumMappingType.MAPPING_TYPE_NONE
                                    map.TargetDataStore = ""
                                    map.TargetParent = ""
                                    map.IsModified = True
                                End If
                            End If
                        Next
                    End If
                Next
                '/// Save the new mappings for the task
                '/// UPDATE(taskmappings)
                success = task.DeleteMappings(cmd)
                If success = True Then
                    success = task.AddNewMappings(cmd)
                Else
                    success = False
                End If
                If success = False Then
                    Exit For
                End If
            Next

            UpdateMappings = success

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "modRename UpdateMappings")
            UpdateMappings = False
        Catch ex As Exception
            LogError(ex, "modRename UpdateMappings")
            UpdateMappings = False
        End Try

    End Function

#End Region

End Module
