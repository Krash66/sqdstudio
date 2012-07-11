Public Module modDatabase

    Function GetNewId() As String
        Return System.Guid.NewGuid().ToString("N")
    End Function

    Function GetVal(ByVal obj As Object) As Object
        If (obj Is System.DBNull.Value) Then
            GetVal = Nothing
        Else
            GetVal = obj
        End If
    End Function

    Function GetStr(ByVal obj As Object) As String
        If obj Is Nothing Then
            GetStr = ""
        Else
            GetStr = obj.ToString
        End If
    End Function

    'Private m_blnDSNUpdated As Boolean = False

    'Public Property DSNUpdated() As Boolean
    '    Get
    '        Return m_blnDSNUpdated
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_blnDSNUpdated = Value
    '    End Set
    'End Property

    '/// UnUsed as of March 2007
#Region "DSN Methods - Added to keep DSN up to date after auto updates occur"

    'Constant Declaration
    Private Const ODBC_ADD_DSN As Integer = 1               ' Add data source
    Private Const ODBC_CONFIG_DSN As Integer = 2            ' Configure (edit) data source
    Private Const ODBC_REMOVE_DSN As Integer = 3            ' Remove data source
    Private Const vbAPINull As Long = 0                     ' NULL Pointer

    Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
         (ByVal hwndParent As Integer, ByVal fRequest As Integer, _
         ByVal lpszDriver As String, ByVal lpszAttributes As String) As Integer

    ' JDM - 3/28/2006 - Update DSN (after program update)
    Function UpdateDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseFilePath As String = "", _
            Optional ByVal Description As String = "" _
            ) As Boolean

        If Description.Trim = "" Then
            Description = "This is default datasource for SQData Studio 1.0"
        End If

        If Not System.IO.File.Exists(DatabaseFilePath) Then ' Copy from parent
            If System.IO.File.Exists(GetParentPath() & "PROJECTS\SQDMeta.mdb") Then
                System.IO.File.Move(GetParentPath() & "PROJECTS\SQDMeta.mdb", DatabaseFilePath)
            Else
                System.IO.File.Move(GetParentPath() & "Temp\SQDMeta.mdb", DatabaseFilePath)
            End If
        End If

        RemoveDSN(DsnName)
        Return CreateDSN(DsnName, DatabaseFilePath, Description)

    End Function

    Function CreateDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseFilePath As String = "", _
            Optional ByVal Description As String = "" _
            ) As Boolean

        Dim intRet As Integer
        Dim strDriver As String
        Dim strAttributes As String

        If DatabaseFilePath.Trim = "" Then
            DatabaseFilePath = System.AppDomain.CurrentDomain.BaseDirectory() & "SQDMeta.mdb"
        End If
        If Description.Trim = "" Then
            Description = "This is default datasource for SQData Studio 1.0"
        End If

        strDriver = "Microsoft Access Driver (*.mdb)" & Chr(0)
        strAttributes = "DSN=" & DsnName & Chr(0) & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0) & "DBQ=" & DatabaseFilePath & Chr(0) & "Description=" & Description & Chr(0)
        intRet = SQLConfigDataSource(vbAPINull, ODBC_ADD_DSN, strDriver, strAttributes)

        If intRet Then
            Return True
        Else
            Return False
        End If

    End Function

    Function RemoveDSN( _
            Optional ByVal DsnName As String = "SQDMeta", _
            Optional ByVal DatabaseName As String = "" _
            ) As Boolean

        Dim intRet As Integer
        Dim strDriver As String
        Dim strAttributes As String

        strDriver = "Microsoft Access Driver (*.mdb)"
        strAttributes = "DSN=" & DsnName & Chr(0)
        intRet = SQLConfigDataSource(vbAPINull, ODBC_REMOVE_DSN, strDriver, strAttributes)

        If intRet Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

    '//// For Future MODs
#Region "ObjectLoad"

    Function LoadProject(ByVal obj As clsProject, Optional ByVal Renamed As Boolean = False, Optional ByVal Updated As Boolean = False) As Boolean
        '// Modified 2/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp")
        Dim i As Integer
        Dim sql As String
        Dim MsgClear As Boolean

        Try
            cnn = New System.Data.Odbc.OdbcConnection(obj.Project.MetaConnectionString)
            cnn.Open()
            Dim cNode As TreeNode = Nothing

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            obj.LoadMe()

            '// Optional Passed Vars added By Tom Karasch for change(i.e. delete, etc...) and rename
            '/// to reload project according to what occured in the object tree.
            If Renamed = False Then
                If Updated = False Then  '// opening new project
                    '//Add project node
                    'cNode = AddNode(tvExplorer.Nodes, NODE_PROJECT, obj)
                    obj.SeqNo = cNode.Index '//store position
                    IsEventFromCode = True
                    'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    IsEventFromCode = False
                    MsgClear = True
                Else   '// Reloading project and tree to reflect MetaData Changes
                    cNode = obj.ObjTreeNode
                    cNode.Nodes.Clear()
                    IsEventFromCode = True
                    'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                    IsEventFromCode = False
                    MsgClear = False
                End If
            Else
                '// Obj is existing project with renamed objects so refill with name changes
                cNode = obj.ObjTreeNode
                cNode.Nodes.Clear()
                cNode.Text = obj.Project.ProjectName
                'SCmain.SplitterDistance = CType(obj, clsProject).MainSeparatorX
                CType(obj, clsProject).Save()
                'CType(obj, clsProject).SaveToRegistry()
                CType(obj, clsProject).SaveToXML()
                MsgClear = False
            End If

            'ShowUsercontrol(cNode, MsgClear)
            'tvExplorer.SelectedNode = cNode

            '//Now add and process each environment 
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,ENVIRONMENTDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblEnvironments & " where ProjectName=" & obj.GetQuotedText

            Log(sql)

            'Log(obj.Project.MetaConnectionString)    //// commented so Password is not in the log
            'Log("cnn.connectionstring >> " & cnn.ConnectionString)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            '//Add Environment folder node
            Dim objF As clsFolderNode
            objF = New clsFolderNode("Environments", NODE_FO_ENVIRONMENT)
            objF.Parent = CType(cNode.Tag, INode)
            cNode = AddNode(cNode, objF.Type, objF)
            objF.SeqNo = cNode.Index '//store position

            For i = 0 To dt.Rows.Count - 1
                '//Process this env
                LoadEnv(cNode, dt.Rows(i), cnn)
            Next


            cNode = cNode.Parent
            cNode.EnsureVisible()


        Catch ex As Exception
            LogError(ex, "fillProject")
        Finally
            ShowStatusMessage("")
        End Try

    End Function

    Function LoadEnv(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As New System.Data.DataTable("temp1")
        Dim i As Integer

        Dim obj As New clsEnvironment
        Dim sql As String = ""
        Dim Dir As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent '//Project->[Env Folder]->Env
            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
                obj.LocalDTDDir = GetVal(dr.Item("LocalDTDDir"))
                obj.LocalDDLDir = GetVal(dr.Item("LocalDDLDir"))
                obj.LocalCobolDir = GetVal(dr.Item("LocalCobolDir"))
                obj.LocalCDir = GetVal(dr.Item("LocalCDir"))
                obj.LocalScriptDir = GetVal(dr.Item("LocalScriptDir"))
                obj.LocalDBDDir = GetVal(dr.Item("LocalModelDir"))

                obj.EnvironmentDescription = GetVal(dr.Item("Description"))

            Else
                obj.EnvironmentName = GetVal(dr.Item("EnvironmentName"))
                obj.EnvironmentDescription = GetVal(dr.Item("ENVIRONMENTDESCRIPTION"))
            End If


            AddToCollection(obj.Project.Environments, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            '/// Added 6/4/07 by TK
            obj.GetStructureDir()

        Catch ex As Exception
            LogError(ex, "fillEnv")
        End Try
        '///////////////////////////////////////////////
        '//Now add connections
        '///////////////////////////////////////////////
        Try
            '//Add connection folder node under Env
            Dim cNodeCnn As TreeNode
            Dim objCnn As INode

            objCnn = New clsFolderNode("Connections", NODE_FO_CONNECTION)
            objCnn.Parent = CType(cNode.Tag, INode)
            cNodeCnn = AddNode(cNode, objCnn.Type, objCnn)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & objCnn.Project.tblConnections & " where EnvironmentName=" & obj.GetQuotedText & " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,CONNECTIONNAME,CONNECTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objCnn.Project.tblConnections & _
                            " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName =" & obj.GetQuotedText

                Log(sql)
            End If


            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                LoadConn(cNodeCnn, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "fillConn")
        End Try
        '///////////////////////////////////////////////
        '//Now add structures
        '///////////////////////////////////////////////
        Try
            '//Add Structure folder node under env
            Dim cNodeStruct As TreeNode
            Dim objStruct As INode
            objStruct = New clsFolderNode("Descriptions", NODE_FO_STRUCT)
            objStruct.Parent = CType(cNode.Tag, INode)
            cNodeStruct = AddNode(cNode, objStruct.Type, objStruct)

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblStructures & _
            '    " where EnvironmentName=" & obj.GetQuotedText & _
            '    " AND ProjectName=" & obj.Project.GetQuotedText & _
            '    " order by structuretype, structureName"
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,DESCRIPTIONTYPE,DESCRIPTIONDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblDescriptions & _
            " where ProjectName=" & obj.Project.GetQuotedText & _
            " AND EnvironmentName = " & obj.GetQuotedText & _
            " order by DESCRIPTIONTYPE, DescriptionName"
            Log(sql)
            'End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New System.Data.DataTable("temp2")

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode1 is root node under which we add other structures)
                LoadStruct(cNodeStruct, dt.Rows(i), cnn)
            Next

            cNodeStruct.Expand()

        Catch ex As Exception
            LogError(ex, "fillEnv-fillStr")
        End Try


        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add variables folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode, objVar.Type, objVar, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from variables where EngineName=" & obj.GetQuotedText & _
                " AND ENGINENAME=" & Quote(DBNULL) & _
                " AND SystemName=" & Quote(DBNULL) & _
                " AND EnvironmentName=" & obj.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables " & _
                            "where EnvironmentName=" & obj.GetQuotedText & _
                            " AND ProjectName=" & obj.Project.GetQuotedText & _
                            " AND SYSTEMNAME=" & Quote(DBNULL) & _
                            " AND ENGINENAME=" & Quote(DBNULL) & _
                            " ORDER BY VARIABLENAME"
                Log(sql)
            End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                LoadVar(cNodeVar, dt.Rows(i), cnn, True)
            Next
        Catch ex As Exception
            LogError(ex, "modObjects LoadEngine>VariableLoad")
        End Try



        '///////////////////////////////////////////////
        '//Now add systems
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Systems", NODE_FO_SYSTEM)
            objSys.Parent = CType(cNode.Tag, INode)
            cNodeSys = AddNode(cNode, objSys.Type, objSys)

            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblSystems & _
                " where EnvironmentName=" & obj.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,SYSTEMDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblSystems & " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName=" & obj.GetQuotedText
                Log(sql)
            End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                loadSys(cNodeSys, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "FillEnv-FillSys")
        End Try

    End Function

    Function LoadStruct(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsStructure
        Dim sql As String = ""
        '///////////////////////////////////////////////
        '//Construct structure object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Tag, INode).Parent
            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.StructureName = GetVal(dr.Item("StructureName"))
                obj.StructureType = GetVal(dr.Item("StructureType"))
                obj.fPath1 = GetVal(dr.Item("fPath1"))
                obj.fPath2 = GetVal(dr.Item("fPath2"))
                obj.IMSDBName = GetVal(dr.Item("IMSDBName"))
                obj.SegmentName = GetVal(dr.Item("SegmentName"))
                obj.StructureDescription = GetVal(dr.Item("Description"))
                obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->StructFolder->Struct
                Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
                If ConnName <> "" Then
                    For Each conn As clsConnection In obj.Environment.Connections
                        If conn.Text = ConnName Then
                            obj.Connection = conn
                        End If
                    Next
                End If
            Else
                obj.StructureName = GetVal(dr.Item("DescriptionName"))
                obj.StructureDescription = GetStr(GetVal(dr.Item("DescriptionDescription")))
                obj.StructureType = GetVal(dr.Item("DescriptionTYPE"))
            End If


            AddToCollection(obj.Environment.Structures, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddStructNode(obj, cNode)
            cNode.Expand()

        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

        '///////////////////////////////////////////////
        '//Now add fieldselections
        '///////////////////////////////////////////////
        Try
            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Select * from " & obj.Project.tblStructSel & " where StructureName=" & obj.GetQuotedText & _
            '    " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & _
            '    obj.Environment.Project.GetQuotedText & " order by selectionname"
            'Else
            sql = "Select PROJECTNAME,ENVIRONMENTNAME,DESCRIPTIONNAME,SELECTIONNAME,ISSYSTEMSEL,SELECTDESCRIPTION from " & obj.Project.tblDescriptionSelect & _
            " where ProjectName=" & obj.Project.GetQuotedText & _
            " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
            " AND DescriptionName = " & obj.GetQuotedText & _
            " order by selectionname"
            'End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode is root node under which we add other nodes)
                LoadStructSel(cNode, dt.Rows(i), cnn)
            Next
            cNode.Collapse()


        Catch ex As Exception
            LogError(ex, "fillStruct")
        End Try

    End Function

    Function LoadStructSel(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsStructureSelection

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.ObjStructure = CType(cNode.Tag, INode)
            If obj.ObjStructure.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSELECT"))
                obj.SelectionName = GetVal(dr.Item("SelectionName"))
                obj.SelectionDescription = GetVal(dr.Item("Description"))
            Else
                obj.IsSystemSelection = GetVal(dr.Item("ISSYSTEMSEL"))
                obj.SelectionName = GetVal(dr.Item("SelectionName"))
                obj.SelectionDescription = GetVal(dr.Item("SELECTDescription"))
            End If

            AddToCollection(obj.ObjStructure.StructureSelections, obj, obj.GUID)

            If obj.IsSystemSelection = "1" Then
                obj.ObjStructure.SysAllSelection = obj
                Exit Function '//dont add this node if system selection
            End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position
            cNode.Collapse()

        Catch ex As Exception
            LogError(ex, "fillstrSel")
        End Try

    End Function

    Function LoadSys(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer

        Dim obj As New clsSystem
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.SystemName = GetVal(dr.Item("SystemName"))
                obj.Port = GetVal(dr.Item("Port"))
                obj.Host = GetVal(dr.Item("Host"))
                obj.QueueManager = GetVal(dr.Item("QMgr"))
                obj.OSType = GetVal(dr.Item("OSType"))

                obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
                obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
                obj.DTDLib = GetVal(dr.Item("DTDLib"))
                obj.DDLLib = GetVal(dr.Item("DDLLib"))

                obj.SystemDescription = GetVal(dr.Item("Description"))
            Else
                obj.SystemName = GetVal(dr.Item("SystemName"))
                obj.SystemDescription = GetVal(dr.Item("SystemDescription"))
            End If

            AddToCollection(obj.Environment.Systems, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex, "fillSys")
        End Try

        '///////////////////////////////////////////////
        '//Now add engines
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeSys As TreeNode
            Dim objSys As INode
            objSys = New clsFolderNode("Engines", NODE_FO_ENGINE)
            objSys.Parent = CType(cNode.Tag, INode)
            cNodeSys = AddNode(cNode, objSys.Type, objSys)

            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & objSys.Project.tblEngines & " where SystemName=" & obj.GetQuotedText & " AND EnvironmentName=" & obj.Environment.GetQuotedText & " AND ProjectName=" & obj.Environment.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,ENGINEDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & objSys.Project.tblEngines & _
                            " where ProjectName=" & obj.Project.GetQuotedText & _
                            " AND EnvironmentName=" & obj.Environment.GetQuotedText & _
                            " AND SystemName=" & obj.GetQuotedText
                Log(sql)
            End If

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                '//Process this (cNode2 is root node under which we add other systems)
                LoadEngine(cNodeSys, dt.Rows(i), cnn)
            Next

        Catch ex As Exception
            LogError(ex, "fillSys")
        End Try

    End Function

    Function LoadEngine(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim da As System.Data.Odbc.OdbcDataAdapter
        Dim dt As System.Data.DataTable
        Dim i As Integer
        Dim Dir As String = ""
        Dim obj As New clsEngine
        Dim sql As String = ""

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try
            obj.EngineName = GetVal(dr.Item("EngineName"))
            obj.ObjSystem = CType(cNode.Parent.Tag, INode) 'Env->SysFolder->Sys

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.ReportEvery = GetVal(dr.Item("ReportEvery"))
            '    obj.CommitEvery = GetVal(dr.Item("CommitEvery"))
            '    obj.ReportFile = GetVal(dr.Item("ReportFile"))
            '    obj.CopybookLib = GetVal(dr.Item("CopybookLib"))
            '    obj.IncludeLib = GetVal(dr.Item("IncludeLib"))
            '    obj.DTDLib = GetVal(dr.Item("DTDLib"))
            '    obj.DDLLib = GetVal(dr.Item("DDLLib"))
            '    obj.EngineDescription = GetVal(dr.Item("Description"))
            '    Dim ConnName As String = GetVal(dr.Item("ConnectionName"))
            '    If ConnName <> "" Then
            '        For Each conn As clsConnection In obj.ObjSystem.Environment.Connections
            '            If conn.Text = ConnName Then
            '                obj.Connection = conn
            '            End If
            '        Next
            '    End If
            '    obj.LoadEngProps()
            'Else
            obj.EngineDescription = GetVal(dr.Item("EngineDescription"))
            'End If

            AddToCollection(obj.ObjSystem.Engines, obj, obj.GUID)

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now source datastore
        '///////////////////////////////////////////////
        Try
            Dim cNode1 As TreeNode
            Dim obj1 As INode

            obj1 = New clsFolderNode("Sources", NODE_FO_SOURCEDATASTORE)
            obj1.Parent = CType(cNode.Tag, INode)
            cNode1 = AddNode(cNode, obj1.Type, obj1, False)


            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='S' AND EngineName=" & _
                obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
                obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
                " where projectname=" & Quote(obj.Project.ProjectName) & _
                " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
                " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
                " and enginename=" & Quote(obj.EngineName) & _
                " and DSDIRECTION='S'" & _
                " order by datastorename"
                Dir = "S"
            End If

            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()



            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                LoadDataStore(cNode1, dt.Rows(i), cnn, Dir)
            Next

        Catch ex As Exception
            LogError(ex)
        End Try

        '///////////////////////////////////////////////
        '//Now add targets
        '///////////////////////////////////////////////
        Try
            Dim cNode2 As TreeNode
            Dim obj2 As INode


            obj2 = New clsFolderNode("Targets", NODE_FO_TARGETDATASTORE)
            obj2.Parent = CType(cNode.Tag, INode)
            cNode2 = AddNode(cNode, obj2.Type, obj2, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from " & obj.Project.tblDatastores & " where DsDirection='T' AND EngineName=" & _
                obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & _
                obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "SELECT PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,DATASTORENAME,DSDIRECTION,DSTYPE,DATASTOREDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID FROM " & obj.Project.tblDatastores & _
                " where projectname=" & Quote(obj.Project.ProjectName) & _
                " and environmentname=" & Quote(obj.ObjSystem.Environment.EnvironmentName) & _
                " and systemname=" & Quote(obj.ObjSystem.SystemName) & _
                " and enginename=" & Quote(obj.EngineName) & _
                " and DSDIRECTION='T'" & _
                " order by datastorename"
                Dir = "T"
            End If
            Log(sql)

            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            dt = New DataTable("temp")
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode1 is root node under which we add other structures)
                LoadDataStore(cNode2, dt.Rows(i), cnn, Dir)
            Next

        Catch ex As Exception
            LogError(ex, sql)
        End Try

        '///////////////////////////////////////////////
        '//Now add variables
        '///////////////////////////////////////////////
        Try
            '//Add systems folder node under env
            Dim cNodeVar As TreeNode
            Dim objVar As INode
            objVar = New clsFolderNode("Variables", NODE_FO_VARIABLE)
            objVar.Parent = CType(cNode.Tag, INode)
            cNodeVar = AddNode(cNode, objVar.Type, objVar, False)

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                sql = "Select * from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEDESCRIPTION from variables where EngineName=" & obj.GetQuotedText & " AND SystemName=" & obj.ObjSystem.GetQuotedText & " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & obj.ObjSystem.Environment.Project.GetQuotedText
                Log(sql)
            End If

            dt = New DataTable("temp")
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)
            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode2 is root node under which we add other systems)
                LoadVar(cNodeVar, dt.Rows(i), cnn)
            Next
        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>VariableLoad")
        End Try

        '///////////////////////////////////////////////
        '//Now add tasks (main, join, lookup, proc)
        '///////////////////////////////////////////////
        Try
            Dim cNodeMain As TreeNode
            'Dim cNodeLookup As TreeNode
            'Dim cNodeJoin As TreeNode
            Dim cNodeProc As TreeNode
            Dim objMain As INode
            'Dim objLook As INode
            'Dim objGen As INode
            Dim objProc As INode
            Dim ttype As enumTaskType

            '//Add Join folder
            'objJoin = New clsFolderNode("Join", NODE_FO_JOIN)
            'objJoin.Parent = CType(cNode.Tag, INode)
            'cNodeJoin = AddNode(cNode.Nodes, objJoin.Type, objJoin, False)

            '//Add Proc folder
            objProc = New clsFolderNode("Procedures", NODE_FO_PROC)
            objProc.Parent = CType(cNode.Tag, INode)
            cNodeProc = AddNode(cNode, objProc.Type, objProc, False)

            '//Add Lookup folder
            'objLook = New clsFolderNode("Lookup", NODE_FO_LOOKUP)
            'objLook.Parent = CType(cNode.Tag, INode)
            'cNodeLookup = AddNode(cNode.Nodes, objLook.Type, objLook, False)

            '//Add Main folder
            objMain = New clsFolderNode("Main Procedure(s)", NODE_FO_MAIN)
            objMain.Parent = CType(cNode.Tag, INode)
            cNodeMain = AddNode(cNode, objMain.Type, objMain, False)
            dt = New DataTable("temp")

            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                If obj.ObjSystem IsNot Nothing Then
                    sql = "Select * from " & obj.Project.tblTasks & _
                    " where EngineName=" & obj.GetQuotedText & _
                    " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
                    " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
                    " AND ProjectName=" & obj.Project.GetQuotedText & _
                    " Order by SeqNo"
                Else
                    sql = "Select * from " & obj.Project.tblTasks & _
                    " where EnvironmentName = " & obj.ObjSystem.Environment.GetQuotedText & _
                    " AND ProjectName=" & obj.Project.GetQuotedText & _
                    " Order by SeqNo"
                End If
            Else
                sql = "Select PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,TASKNAME,TASKTYPE,TASKSEQNO,TASKDESCRIPTION,CREATED_TIMESTAMP,UPDATED_TIMESTAMP,CREATED_USER_ID,UPDATED_USER_ID from " & obj.Project.tblTasks & _
                " where EngineName=" & obj.GetQuotedText & _
                " AND SystemName=" & obj.ObjSystem.GetQuotedText & _
                " AND EnvironmentName=" & obj.ObjSystem.Environment.GetQuotedText & _
                " AND ProjectName=" & obj.Project.GetQuotedText & _
                " Order by TASKSEQNO"
            End If

            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cnn)

            da.Fill(dt)
            da.Dispose()

            For i = 0 To dt.Rows.Count - 1
                ''//Process this (cNode is root node under which we add other nodes)
                ttype = GetVal(dt.Rows(i).Item("TaskType"))
                Select Case ttype
                    Case enumTaskType.TASK_MAIN
                        LoadTasks(cNodeMain, dt.Rows(i), cnn)
                    Case enumTaskType.TASK_GEN, enumTaskType.TASK_LOOKUP, enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc
                        LoadTasks(cNodeProc, dt.Rows(i), cnn)
                        'Case enumTaskType.TASK_JOIN, enumTaskType.TASK_LOOKUP
                        '    FillTasks(cNodeJoin, dt.Rows(i), cnn)
                        'Case 
                        '    FillTasks(cNodeLookup, dt.Rows(i), cnn)
                End Select
            Next
            cNode.Expand()

            For Each ds As clsDatastore In obj.Sources
                ds.SetIsMapped(False, True)
            Next
            For Each ds As clsDatastore In obj.Targets
                ds.SetIsMapped(False, True)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillEngine>fillTasks", sql)
            Return False
        End Try

    End Function

    Function LoadTasks(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean
        '// Modified 1/07 by Tom Karasch

        Dim obj As New clsTask
        Dim Otype As String

        '///////////////////////////////////////////////
        '//Construct object from database values and add
        '///////////////////////////////////////////////
        Try

            Otype = CType(cNode.Parent.Tag, INode).Type
            If Otype = NODE_ENVIRONMENT Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->TaskFolder->Any Task
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.TaskName = GetVal(dr.Item("TaskName"))
            obj.TaskType = GetVal(dr.Item("TaskType"))
            If obj.Parent.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.TaskDescription = GetVal(dr.Item("Description"))
            Else
                obj.TaskDescription = GetStr(GetVal(dr("TASKDESCRIPTION")))
            End If

            Select Case obj.TaskType
                'Case modDeclares.enumTaskType.TASK_JOIN, modDeclares.enumTaskType.TASK_LOOKUP
                '    If obj.Engine IsNot Nothing Then
                '        AddToCollection(obj.Engine.Joins, obj, obj.GUID)
                '    Else
                '        AddToCollection(obj.Environment.Joins, obj, obj.GUID)
                '    End If
                Case modDeclares.enumTaskType.TASK_MAIN
                    If obj.Engine IsNot Nothing Then
                        AddToCollection(obj.Engine.Mains, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    End If
                Case modDeclares.enumTaskType.TASK_PROC, enumTaskType.TASK_IncProc, _
                modDeclares.enumTaskType.TASK_GEN, modDeclares.enumTaskType.TASK_LOOKUP
                    If obj.Engine IsNot Nothing Then
                        AddToCollection(obj.Engine.Procs, obj, obj.GUID)
                    Else
                        AddToCollection(obj.Environment.Procedures, obj, obj.GUID)
                    End If
            End Select

            'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    obj.LoadDatastores()
            '    obj.LoadMappings()
            'End If

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store treeview node index

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function LoadVar(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal EnvLoad As Boolean = False) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsVariable
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            If EnvLoad = True Then
                obj.Engine = Nothing
                obj.Environment = CType(cNode.Parent.Tag, INode)
                obj.Parent = obj.Environment
            Else
                obj.Engine = CType(cNode.Parent.Tag, INode) 'Engine->VarFolder->Var
                obj.Environment = obj.Engine.ObjSystem.Environment
                obj.Parent = obj.Engine
            End If

            obj.VariableName = GetVal(dr.Item("VariableName"))


            If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.VarInitVal = GetVal(dr.Item("VarInitVal"))
                obj.VarSize = GetVal(dr.Item("VarSize"))
                obj.VariableDescription = GetVal(dr.Item("Description"))
            Else
                obj.VariableDescription = GetVal(dr.Item("VariableDescription"))
            End If

            If EnvLoad = True Then
                AddToCollection(obj.Environment.Variables, obj, obj.GUID)
            Else
                AddToCollection(obj.Engine.Variables, obj, obj.GUID)
            End If


            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex, "frmMain fillVar")
        End Try

    End Function

    Function LoadConn(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection) As Boolean

        '// Modified 1/07 by Tom Karasch
        Dim obj As New clsConnection
        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            If CType(cNode.Tag, INode).Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
                obj.ConnectionType = GetVal(dr.Item("ConnectionType"))
                obj.ConnectionDescription = GetVal(dr.Item("Description"))
                obj.Database = GetVal(dr.Item("DBName"))
                obj.UserId = GetVal(dr.Item("UserId"))
                obj.Password = GetVal(dr.Item("Password"))
                obj.DateFormat = GetVal(dr.Item("Dateformat"))
            Else
                obj.ConnectionName = GetVal(dr.Item("ConnectionName"))
                obj.ConnectionDescription = GetVal(dr.Item("ConnectionDescription"))
            End If

            obj.Environment = CType(cNode.Parent.Tag, INode) 'Env->ConnFolder->Conn
            AddToCollection(obj.Environment.Connections, obj, obj.GUID)

            cNode = AddNode(cNode, obj.Type, obj, False)
            obj.SeqNo = cNode.Index '//store position

        Catch ex As Exception
            LogError(ex)
        End Try

    End Function

    Function LoadDataStore(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal DSD As String = "") As Boolean

        '/// Modified by Tom Karasch 12/06
        Dim obj As New clsDatastore

        '///////////////////////////////////////////////
        '//Construct object form database values and add
        '///////////////////////////////////////////////
        Try
            obj.Parent = CType(cNode.Parent.Tag, INode) 'Engine->Folder(source/target)->ds
            obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

            If DSD = "" Then
                obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
                obj.DsDirection = GetVal(dr.Item("DsDirection"))
                obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
                obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
                'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
                obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
                obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
                'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
                obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

                obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
                obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
                obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
                obj.DatastoreDescription = GetVal(dr.Item("Description"))
                obj.OperationType = GetVal(dr.Item("OperationType"))
                obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
                obj.DsPort = GetVal(dr.Item("Port"))
                obj.DsUOW = GetVal(dr.Item("UOW"))
                'obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
                'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
                '/// Load Globals From File
                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                '    obj.LoadGlobals()
                'End If

                '//If AnyTree Object is renamed, then reload all datastore items to 
                '// Make sure renaming is propagated to all nodes
                obj.LoadItems()
            Else
                obj.DatastoreDescription = GetVal(dr.Item("DATASTOREDESCRIPTION"))

                If obj.DatastoreDescription Is Nothing Then
                    obj.DatastoreDescription = ""
                End If
                If DSD = "S" Then
                    obj.DsDirection = DS_DIRECTION_SOURCE
                Else
                    obj.DsDirection = DS_DIRECTION_TARGET
                End If

                obj.LoadItems(, True)

                obj.LoadMe()

            End If

            Select Case obj.DsDirection
                Case DS_DIRECTION_SOURCE
                    'If obj.Engine Is Nothing Then
                    'AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    'Else
                    AddToCollection(obj.Engine.Sources, obj, obj.GUID)
                    'End If
                Case DS_DIRECTION_TARGET
                    'If obj.Engine Is Nothing Then
                    'AddToCollection(obj.Environment.Datastores, obj, obj.GUID)
                    'Else
                    AddToCollection(obj.Engine.Targets, obj, obj.GUID)
                    'End If
            End Select

            ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

            cNode = AddNode(cNode, obj.Type, obj, False, obj.Text)
            obj.SeqNo = cNode.Index '//store position

            cNode.Text = obj.DsPhysicalSource

            '// now add Datastore selections to tree
            AddDSstructuresToTree(cNode, obj)

            cNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "frmMain FillDatastore")
            Return False
        End Try

    End Function

    'Function LoadDSbyType(ByVal cNode As TreeNode, ByVal dr As System.Data.DataRow, ByRef cnn As System.Data.Odbc.OdbcConnection, Optional ByVal Ver As enumMetaVersion = enumMetaVersion.V2) As Boolean

    '    Try
    '        Dim obj As New clsDatastore

    '        obj.Parent = CType(cNode.Tag, INode).Parent
    '        obj.DatastoreName = GetVal(dr.Item("DatastoreName"))

    '        'If Ver = enumMetaVersion.V2 Then
    '        '    obj.DatastoreType = GetVal(dr.Item("DatastoreType"))
    '        '    obj.DatastoreDescription = GetVal(dr.Item("Description"))
    '        '    obj.DsPhysicalSource = GetVal(dr.Item("DsPhysicalSource"))
    '        '    obj.DsDirection = GetVal(dr.Item("DsDirection"))
    '        '    obj.DsAccessMethod = GetVal(dr.Item("DsAccessMethod"))
    '        '    obj.DsCharacterCode = GetVal(dr.Item("DsCharacterCode"))
    '        '    'obj.IsOrdered = GetVal(dr.Item("IsOrdered"))
    '        '    obj.IsIMSPathData = GetVal(dr.Item("IsIMSPathData"))
    '        '    obj.IsSkipChangeCheck = GetVal(dr.Item("ISSKIPCHGCHECK"))
    '        '    'obj.IsBeforeImage = GetVal(dr.Item("IsBeforeImage")) '//new by npatel on 8/10/05
    '        '    obj.ExceptionDatastore = GetVal(dr.Item("ExceptDatastore"))

    '        '    obj.TextQualifier = GetVal(dr.Item("TextQualifier"))
    '        '    obj.RowDelimiter = GetVal(dr.Item("RowDelimiter"))
    '        '    obj.ColumnDelimiter = GetVal(dr.Item("ColumnDelimiter"))
    '        '    obj.DatastoreDescription = GetVal(dr.Item("Description"))
    '        '    obj.OperationType = GetVal(dr.Item("OperationType"))
    '        '    obj.DsQueMgr = GetVal(dr.Item("QueMgr"))
    '        '    obj.DsPort = GetVal(dr.Item("Port"))
    '        '    obj.DsUOW = GetVal(dr.Item("UOW"))
    '        '    obj.Poll = GetVal(dr.Item("SelectEvery")) '// added by TK 11/9/2006
    '        '    'obj.IsCmmtKey = GetVal(dr.Item("OnCmmtKey")) '// added by TK 11/9/2006
    '        '    '/// Load Globals From File
    '        '    'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
    '        '    '    obj.LoadGlobals()
    '        '    'End If

    '        '    '//If AnyTree Object is renamed, then reload all datastore items to 
    '        '    '// Make sure renaming is propagated to all nodes
    '        '    obj.LoadItems()
    '        'Else
    '        obj.DatastoreDescription = GetStr(GetVal(dr.Item("DATASTOREDESCRIPTION")))
    '        obj.DatastoreType = GetVal(dr.Item("DSTYPE"))

    '        obj.LoadItems(, True)

    '        obj.LoadMe()
    '        'End If


    '        AddToCollection(obj.Environment.Datastores, obj, obj.GUID)

    '        ShowStatusMessage("Loading ....[" & obj.Key.Replace("-", "->") & "]")

    '        cNode = AddDSNode(obj, cNode)
    '        obj.SeqNo = cNode.Index '//store position

    '        cNode.Text = obj.DsPhysicalSource

    '        '// now add Datastore selections to tree
    '        AddDSstructuresToTree(cNode, obj)
    '        cNode.Collapse()

    '        Return True

    '    Catch ex As Exception
    '        LogError(ex, "frmMain FillDSbyType")
    '        Return False
    '    End Try

    'End Function

    '//This will add Datastore node under proper Category Folder
    '//cNode : is "Datastores" folder node
    '//obj : is object to represent struct node
    '// [Datastores]
    '//     |
    '//    [+]-[<DSType>]
    '//         |
    '//         [+]-DS1
    '//         [+]-DS2
    'Function AddDSNode(ByVal obj As clsDatastore, ByVal cNode As TreeNode) As TreeNode
    '    '//Now look for Datastore type and insert it in proper folder
    '    Dim nd As TreeNode = Nothing
    '    Dim IsFound As Boolean

    '    Try
    '        For Each nd In cNode.Nodes
    '            If nd.Text = GetDSFolderText(obj.DatastoreType) Then
    '                IsFound = True
    '                Exit For
    '            End If
    '        Next

    '        '//if folder for structure category is not found then add new folder 
    '        '//and then add structure node under it
    '        If IsFound = True Then
    '            '//Add struct node under struct category
    '            cNode = AddNode(nd.Nodes, obj.Type, obj, False)
    '            obj.SeqNo = cNode.Index '//store position
    '        Else
    '            Dim objFol As INode
    '            Dim objFolType As String = GetFolType(obj.DatastoreType)

    '            objFol = New clsFolderNode(GetDSFolderText(obj.DatastoreType), objFolType)

    '            objFol.Parent = CType(cNode.Parent.Tag, INode)
    '            '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
    '            nd = AddNode(cNode.Nodes, objFol.Type, objFol, False)
    '            '//Add struct node under struct category
    '            cNode = AddNode(nd.Nodes, obj.Type, obj, False, obj.DatastoreName)
    '            obj.SeqNo = cNode.Index '//store position
    '        End If
    '        Return cNode

    '    Catch ex As Exception
    '        LogError(ex, "frmMain AddDSNode")
    '        Return Nothing
    '    End Try

    'End Function

    'Function GetDSFolderText(ByVal DSType As enumDatastore) As String

    '    Select Case DSType
    '        Case modDeclares.enumDatastore.DS_BINARY
    '            Return "BINARY"
    '        Case modDeclares.enumDatastore.DS_DB2CDC
    '            Return "DB2CDC"
    '        Case modDeclares.enumDatastore.DS_DB2LOAD
    '            Return "DB2LOAD"
    '        Case modDeclares.enumDatastore.DS_DELIMITED
    '            Return "DELIMITED"
    '        Case modDeclares.enumDatastore.DS_GENERICCDC
    '            Return "GENERICCDC"
    '        Case modDeclares.enumDatastore.DS_HSSUNLOAD
    '            Return "HSSUNLOAD"
    '        Case modDeclares.enumDatastore.DS_IBMEVENT
    '            Return "IBMEVENT"
    '            'Case modDeclares.enumDatastore.DS_IMSCDC
    '            '    Return "IMSCDC"
    '        Case modDeclares.enumDatastore.DS_IMSDB
    '            Return "IMSDB"
    '            'Case modDeclares.enumDatastore.DS_IMSLE
    '            '    Return "IMSLE"
    '            'Case modDeclares.enumDatastore.DS_IMSLEBATCH
    '            '    Return "IMSLEBATCH"
    '        Case modDeclares.enumDatastore.DS_INCLUDE
    '            Return "INCLUDE"
    '        Case modDeclares.enumDatastore.DS_ORACLECDC
    '            Return "ORACLECDC"
    '        Case modDeclares.enumDatastore.DS_RELATIONAL
    '            Return "RELATIONAL"
    '        Case modDeclares.enumDatastore.DS_SUBVAR
    '            Return "SUBVAR"
    '        Case modDeclares.enumDatastore.DS_TEXT
    '            Return "TEXT"
    '            'Case modDeclares.enumDatastore.DS_TRBCDC
    '            '    Return "TRBCDC"
    '        Case modDeclares.enumDatastore.DS_UNKNOWN
    '            Return "UNKNOWN"
    '        Case modDeclares.enumDatastore.DS_VSAM
    '            Return "VSAM"
    '        Case modDeclares.enumDatastore.DS_VSAMCDC
    '            Return "VSAMCDC"
    '        Case modDeclares.enumDatastore.DS_XML
    '            Return "XML"
    '            'Case modDeclares.enumDatastore.DS_XMLCDC
    '            '    Return "XMLCDC"
    '        Case Else
    '            Return "Unknown"
    '    End Select

    'End Function

    'Function GetFolType(ByVal DStype As enumDatastore) As String

    '    Select Case DStype
    '        Case enumDatastore.DS_BINARY
    '            Return DS_BINARY
    '        Case enumDatastore.DS_DB2CDC
    '            Return DS_DB2CDC
    '        Case enumDatastore.DS_DB2LOAD
    '            Return DS_DB2LOAD
    '        Case enumDatastore.DS_DELIMITED
    '            Return DS_DELIMITED
    '        Case enumDatastore.DS_GENERICCDC
    '            Return DS_GENERICCDC
    '        Case enumDatastore.DS_HSSUNLOAD
    '            Return DS_HSSUNLOAD
    '        Case enumDatastore.DS_IBMEVENT
    '            Return DS_IBMEVENT
    '        Case enumDatastore.DS_IMSCDCLE
    '            Return DS_IMSCDCLE
    '        Case enumDatastore.DS_IMSDB
    '            Return DS_IMSDB
    '            'Case enumDatastore.DS_IMSLE
    '            '    Return DS_IMSLE
    '            'Case enumDatastore.DS_IMSLEBATCH
    '            '    Return DS_IMSLEBATCH
    '        Case enumDatastore.DS_INCLUDE
    '            Return DS_INCLUDE
    '        Case enumDatastore.DS_ORACLECDC
    '            Return DS_ORACLECDC
    '        Case enumDatastore.DS_RELATIONAL
    '            Return DS_RELATIONAL
    '        Case enumDatastore.DS_SUBVAR
    '            Return DS_SUBVAR
    '        Case enumDatastore.DS_TEXT
    '            Return DS_TEXT
    '            'Case enumDatastore.DS_TRBCDC
    '            '    Return DS_TRBCDC
    '        Case enumDatastore.DS_UNKNOWN
    '            Return DS_UNKNOWN
    '        Case enumDatastore.DS_VSAM
    '            Return DS_VSAM
    '        Case enumDatastore.DS_VSAMCDC
    '            Return DS_VSAMCDC
    '        Case enumDatastore.DS_XML
    '            Return DS_XML
    '            'Case enumDatastore.DS_XMLCDC
    '            '    Return DS_XMLCDC
    '        Case Else
    '            Return NODE_FO_DATASTORE
    '    End Select

    'End Function

    Function AddDSstructuresToTree(ByVal pNode As TreeNode, ByVal obj As clsDatastore, Optional ByVal MapAs As Boolean = False) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            pNode.Nodes.Clear()
            obj.LoadItems(, True)


            For i = 0 To obj.ObjSelections.Count - 1
                Dim DSselobj As clsDSSelection = obj.ObjSelections(i)

                'If obj.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
                DSselobj.LoadItems(, True)
                'End If

                If Not DSselobj.IsChildDSSelection = True Or MapAs = True Then
                    cnode = AddNode(pNode, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    Call addDSSelChildrenToTree(cnode, DSselobj)
                End If
            Next

            pNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "main AddDSstructuresToTree")
            Return False
        End Try

    End Function

    Private Function addDSSelChildrenToTree(ByRef pNode As TreeNode, ByRef obj As clsDSSelection) As Boolean

        Dim i As Integer
        Dim cnode As TreeNode

        Try
            For i = 0 To obj.ObjDSSelections.Count - 1
                Dim DSselobj As clsDSSelection
                DSselobj = obj.ObjDSSelections(i)
                DSselobj.LoadItems(, True)
                If DSselobj.Parent Is obj Then
                    cnode = AddNode(pNode, DSselobj.Type, DSselobj, True)
                    DSselobj.SeqNo = pNode.Index '//store position
                    Call addDSSelChildrenToTree(cnode, DSselobj) '// recurse for all children
                End If
            Next
            pNode.Collapse()

            Return True

        Catch ex As Exception
            LogError(ex, "main addDSselChildrenToTree")
            Return False
        End Try

    End Function

    '//This will add Structure node under proper Category Folder
    '//cNode : is "Structures" folder node
    '//obj : is object to represent struct node
    '// [Structures]
    '//     |
    '//    [+]-[<category>]
    '//         |
    '//         [+]-Struct1
    '//         [+]-struct2
    Function AddStructNode(ByVal obj As clsStructure, ByVal cNode As TreeNode) As TreeNode
        '//Now look for Structure type and insert it in proper folder
        Dim nd As TreeNode = Nothing
        Dim IsFound As Boolean

        Try
            For Each nd In cNode.Nodes
                If nd.Text = GetStructureFolderText(obj.StructureType) Then
                    IsFound = True
                    Exit For
                End If
            Next

            '//if folder for structure category is not found then add new folder 
            '//and then add structure node under it
            If IsFound = True Then
                '//Add struct node under struct category
                cNode = AddNode(nd, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            Else
                Dim objFol As INode
                objFol = New clsFolderNode(GetStructureFolderText(obj.StructureType), NODE_FO_STRUCT)
                objFol.Parent = CType(cNode.Parent.Tag, INode)
                '//Add Struct Type Folder (i.e. [XML] [COBOLIMS])
                nd = AddNode(cNode, objFol.Type, objFol, False)
                '//Add struct node under struct category
                cNode = AddNode(nd, obj.Type, obj, False)
                obj.SeqNo = cNode.Index '//store position
            End If
            Return cNode

        Catch ex As Exception
            Log("AddStructNode=>" & ex.Message)
            Return Nothing
        End Try

    End Function

    Function GetStructureFolderText(ByVal stType As enumStructure) As String

        Select Case stType
            Case modDeclares.enumStructure.STRUCT_C
                Return "CHeader"
            Case modDeclares.enumStructure.STRUCT_COBOL
                Return "COBOL"
            Case modDeclares.enumStructure.STRUCT_COBOL_IMS
                Return "COBOLIMS"
            Case modDeclares.enumStructure.STRUCT_IMS
                Return "IMS"
            Case modDeclares.enumStructure.STRUCT_REL_DDL
                Return "DDL"
            Case modDeclares.enumStructure.STRUCT_REL_DML, enumStructure.STRUCT_REL_DML_FILE
                Return "DML"
            Case modDeclares.enumStructure.STRUCT_XMLDTD
                Return "XMLDTD"
            Case Else
                Return "Unknown"
        End Select

    End Function

    '//////////////////////////////////////////////////////////////////
    '//// Now the tree is built from loaded data in the MetaDataDB ////
    '//////////////////////////////////////////////////////////////////

#End Region

End Module