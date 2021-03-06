Public Class clsEngine
    Implements INode

    Private m_EngineName As String = ""
    Private m_ReportFile As String = ""
    Private m_ReportEvery As Integer
    Private m_CommitEvery As Integer
    Private m_ForceCommit As Boolean = False
    Private m_EngineDescription As String = ""
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_ObjSystem As clsSystem
    Private m_IsModified As Boolean
    Private m_IsRenamed As Boolean = False
    Private m_MapGroupItems As Boolean = False
    Private m_Connection As clsConnection
    Private m_DateFormat As String = ""
    Private m_Main As String = ""
    Private m_IsLoaded As Boolean = False
    Private m_EngVersion As String

    Public Sources As New Collection
    Public Targets As New Collection
    Public Variables As New Collection
    Public Mains As New Collection
    Public Procs As New Collection
    Public Gens As New Collection
    Public Lookups As New Collection
    Public DBDLib As String = ""
    Public CopybookLib As String = ""
    Public IncludeLib As String = ""
    Public DTDLib As String = ""
    Public DDLLib As String = ""
    '/// for windows Systems
    Public EXEdir As String = ""
    Public BATdir As String = ""
    Public CDCdir As String = ""
    '/// File writing for Non-Metadata Property values
    Dim fsEng As System.IO.FileStream
    Dim objWriteEng As System.IO.StreamWriter
    Dim objReadEng As System.IO.StreamReader
    '/// AddFlow additions
    Private m_ObjAddFlow As AddFlow
    Private m_ObjTabPage As TabPage
    Private m_ObjAddFlowCtl As ctlAddFlowTab
    Private m_CurAddFlowDiagPath As String
    Private m_IsDiagramChanged As Boolean

#Region "INode Implementation"

    Public Function GetQuotedText(Optional ByVal bFix As Boolean = True) As String Implements INode.GetQuotedText
        If bFix = True Then
            Return "'" & FixStr(Me.Text.Trim) & "'"
        Else
            Return "'" & Me.Text.Trim & "'"
        End If
    End Function

    Public Property SeqNo() As Integer Implements INode.SeqNo
        Get
            Return m_SeqNo
        End Get
        Set(ByVal Value As Integer)
            m_SeqNo = Value
        End Set
    End Property

    Public ReadOnly Property GUID() As String Implements INode.GUID
        Get
            Return m_GUID
        End Get
    End Property

    Public Function IsFolderNode() As Boolean Implements INode.IsFolderNode
        If Strings.Left(Me.Type, 3) = "FO_" Then
            IsFolderNode = True
        Else
            IsFolderNode = False
        End If
    End Function

    Public Property ObjTreeNode() As System.Windows.Forms.TreeNode Implements INode.ObjTreeNode
        Get
            Return m_ObjTreeNode
        End Get
        Set(ByVal Value As System.Windows.Forms.TreeNode)
            m_ObjTreeNode = Value
        End Set
    End Property

    Public Property Parent() As INode Implements INode.Parent
        Get
            Return Me.ObjSystem
        End Get
        Set(ByVal Value As INode)
            m_ObjSystem = Value
        End Set
    End Property
    '
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Try
            Dim obj As New clsEngine
            Dim cmd As System.Data.Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)

            obj.EngineName = Me.EngineName
            obj.EngineDescription = Me.EngineDescription
            obj.ReportFile = Me.ReportFile
            obj.ReportEvery = Me.ReportEvery
            obj.CommitEvery = Me.CommitEvery
            obj.SeqNo = Me.SeqNo
            obj.CopybookLib = Me.CopybookLib
            obj.IncludeLib = Me.IncludeLib
            obj.DTDLib = Me.DTDLib
            obj.DDLLib = Me.DDLLib
            obj.IsModified = Me.IsModified
            obj.Parent = NewParent 'Me.Parent
            obj.Connection = Me.Connection
            obj.EXEdir = Me.EXEdir
            obj.BATdir = Me.BATdir
            obj.CDCdir = Me.CDCdir
            obj.EngVersion = Me.EngVersion

            If Cascade = True Then

                '//clone all sources under Engine
                For Each src As clsDatastore In Me.Sources

                    Dim NewSrc As New clsDatastore
                    NewSrc = src.Clone(obj, True, cmd)
                    NewSrc.Engine = obj
                    AddToCollection(obj.Sources, NewSrc, NewSrc.GUID)
                Next


                '//clone all targets under Engine
                For Each tar As clsDatastore In Me.Targets

                    Dim NewTar As New clsDatastore
                    NewTar = tar.Clone(obj, True, cmd)
                    NewTar.Engine = obj
                    AddToCollection(obj.Targets, NewTar, NewTar.GUID)
                Next


                '//clone all variables under Engine
                For Each var As clsVariable In Me.Variables

                    Dim NewVar As New clsVariable
                    NewVar = var.Clone(obj, True, cmd)
                    NewVar.Engine = obj
                    AddToCollection(obj.Variables, NewVar, NewVar.GUID)
                Next


                '//clone all main procs under Engine
                For Each man As clsTask In Me.Mains
                    'man.LoadMe(cmd)

                    Dim NewMan As New clsTask
                    NewMan = man.Clone(obj, True, cmd)
                    NewMan.Engine = obj
                    AddToCollection(obj.Mains, NewMan, NewMan.GUID)
                Next


                '//clone all procs under Engine
                For Each prc As clsTask In Me.Procs
                    'prc.LoadMe(cmd)

                    Dim NewPrc As New clsTask
                    NewPrc = prc.Clone(obj, True, cmd)
                    NewPrc.Engine = obj
                    AddToCollection(obj.Procs, NewPrc, NewPrc.GUID)
                Next

                'Dim NewGen As clsTask
                '//clone all procs under Engine
                'For Each gen As clsTask In Me.Gens
                '    'gen.LoadMe()
                '    Dim NewGen As New clsTask
                '    NewGen = gen.Clone(obj, True, cmd)
                '    NewGen.Engine = obj
                '    AddToCollection(obj.Procs, NewGen, NewGen.GUID)
                'Next
                'Dim NewJon As clsTask
                ''//clone all joins under Engine
                'For Each jon As clsTask In Me.Gens
                '    NewJon = jon.Clone(obj, True, cmd)
                '    NewJon.Engine = obj
                '    AddToCollection(obj.Gens, NewJon, NewJon.GUID)
                'Next

                'Dim NewLoc As clsTask
                ''//clone all lookups under Engine
                'For Each loc As clsTask In Me.Lookups
                '    NewLoc = loc.Clone(obj, True, cmd)
                '    NewLoc.Engine = obj
                '    AddToCollection(obj.Lookups, NewLoc, NewLoc.GUID)
                'Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsEngine Clone")
            Return Nothing
        End Try

    End Function

    Public Property IsModified() As Boolean Implements INode.IsModified
        Get
            Return m_IsModified
        End Get
        Set(ByVal Value As Boolean)
            m_IsModified = Value
        End Set
    End Property

    Public ReadOnly Property Key() As String Implements INode.Key
        Get
            'Key = EngineId
            Key = Me.ObjSystem.Environment.Project.Text & KEY_SAP & Me.ObjSystem.Environment.Text & KEY_SAP & Me.ObjSystem.Text & KEY_SAP & Me.Text
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = EngineName
        End Get
        Set(ByVal Value As String)
            EngineName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_ENGINE
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.Parent.Project
            Return p
        End Get
    End Property

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim sql As String = ""
        Dim cmd As New Odbc.OdbcCommand
        Dim ConnStr As String = ""
        Dim tran As Odbc.OdbcTransaction = Nothing

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            If Me.Connection IsNot Nothing Then
                ConnStr = Me.Connection.ConnectionName
            End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblEngines & " " & _
            '    "set EngineName='" & FixStr(Me.EngineName) & "' , " & _
            '    "ReportEvery=" & Me.ReportEvery & " , " & _
            '    "CommitEvery=" & Me.CommitEvery & " , " & _
            '    "ReportFile='" & FixStr(Me.ReportFile) & "' , " & _
            '    "CopybookLib='" & FixStr(Me.CopybookLib) & "' , " & _
            '    "IncludeLib='" & FixStr(Me.IncludeLib) & "' , " & _
            '    "DTDLib='" & FixStr(Me.DTDLib) & "' , " & _
            '    "DDLLib='" & FixStr(Me.DDLLib) & "' , " & _
            '    "Description='" & FixStr(Me.EngineDescription) & "' , " & _
            '    "CONNECTIONNAME='" & ConnStr & "' " & _
            '    "where EngineName=" & Me.GetQuotedText & " AND SystemName=" & Me.ObjSystem.GetQuotedText & " AND EnvironmentName=" & Me.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.ObjSystem.Environment.Project.GetQuotedText
            'Else
            sql = "Update " & Me.Project.tblEngines & " set EngineName='" & FixStr(Me.EngineName) & "'," & _
            "EngineDescription='" & FixStr(Me.EngineDescription) & "' " & _
            "where EngineName=" & Me.GetQuotedText & " AND SystemName=" & Me.ObjSystem.GetQuotedText & _
            " AND EnvironmentName=" & Me.ObjSystem.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.ObjSystem.Environment.Project.GetQuotedText

            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.DeleteATTR(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If
            If Me.InsertATTR(cmd) = False Then
                tran.Rollback()
                Save = False
                Exit Try
            End If

            Me.IsModified = False

            tran.Commit()

            Save = SaveAddFlowXML()


        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsEngine Save", sql)
            tran.Rollback()
            Save = False
        Catch ex As Exception
            LogError(ex, "clsEngine Save", sql)
            tran.Rollback()
            Save = False
        Finally
            'cnn.Close()
        End Try

    End Function

    '/// OverHauled by TKarasch    April,May 07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            If Cascade = True Then
                Dim i As Integer

                For i = 1 To Me.Mains.Count
                    If Me.Mains(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Mains, "") '//clear all

                For i = 1 To Me.Procs.Count
                    If Me.Procs(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Procs, "") '//clear all  

                For i = 1 To Me.Lookups.Count
                    If Me.Lookups(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Lookups, "") '//clear all

                For i = 1 To Me.Variables.Count
                    If Me.Variables(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Variables, "") '//clear all

                For i = 1 To Me.Gens.Count
                    If Me.Gens(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Gens, "") '//clear all

                For i = 1 To Me.Sources.Count
                    If Me.Sources(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Sources, "")  '//clear all

                For i = 1 To Me.Targets.Count
                    If Me.Targets(i).Delete(cmd, cnn, Cascade, False) = False Then
                        Delete = False
                        Exit Try
                    End If
                Next

                RemoveFromCollection(Me.Targets, "") '//clear all
            End If

            If Me.DeleteATTR(cmd) = False Then
                Delete = False
                Exit Try
            End If

            sql = "Delete From " & Me.Project.tblEngines & " where  EngineName=" & Me.GetQuotedText & " AND SystemName=" & Me.ObjSystem.GetQuotedText & " AND EnvironmentName=" & Me.ObjSystem.Environment.GetQuotedText & " AND ProjectName=" & Me.ObjSystem.Environment.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                RemoveFromCollection(Me.ObjSystem.Engines, Me.GUID)
            End If

            '/// AddFlow Additions
            If Me.ObjTabPage IsNot Nothing Then
                Dim TC As TabControl = CType(Me.ObjTabPage.Parent, TabControl)
                TC.TabPages.Remove(Me.ObjTabPage)
            End If
            '/////////////////////////

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsEngine Delete", sql)
            Delete = False
        End Try

    End Function

    '// modified by TK and KS 11/3/06
    '/// OverHauled by TKarasch    April,May 07
    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand
        Dim sql As String = ""
        Dim ConnText As String
        Dim tran As Odbc.OdbcTransaction = Nothing

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn
            tran = cnn.BeginTransaction()
            cmd.Transaction = tran

            If Me.Connection Is Nothing Then
                ConnText = ""
            Else
                ConnText = Me.Connection.Text
            End If

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    '//When we add new record we need to find Unique Database Record ID.
            '    sql = "INSERT INTO " & Me.Project.tblEngines & "(ProjectName,EnvironmentName,SystemName,EngineName,ReportEvery,CommitEvery,ReportFile,Description,CopybookLib,IncludeLib,DTDLib,DDLLib,ConnectionName) " & _
            '          " Values(" & Me.ObjSystem.Environment.Project.GetQuotedText & "," & Me.ObjSystem.Environment.GetQuotedText & "," & Me.ObjSystem.GetQuotedText & "," & Me.GetQuotedText & "," & Me.ReportEvery & "," & Me.CommitEvery & ",'" & FixStr(Me.ReportFile) & "','" & FixStr(EngineDescription) & "','" & FixStr(CopybookLib) & "','" & FixStr(IncludeLib) & "','" & FixStr(DTDLib) & "','" & FixStr(DDLLib) & "','" & ConnText & "')"
            'Else
            '//When we add new record we need to find Unique Database Record ID.
            sql = "INSERT INTO " & Me.Project.tblEngines & "(ProjectName,EnvironmentName,SystemName,EngineName,EngineDescription) " & _
            " Values(" & Me.Project.GetQuotedText & "," & _
            Me.ObjSystem.Environment.GetQuotedText & "," & _
            Me.ObjSystem.GetQuotedText & "," & _
            Me.GetQuotedText & ",'" & _
            FixStr(Me.EngineDescription) & "')"

            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.DeleteATTR(cmd) = False Then
                tran.Rollback()
                AddNew = False
                Exit Try
            End If
            If Me.InsertATTR(cmd) = False Then
                tran.Rollback()
                AddNew = False
                Exit Try
            End If

            If AddToCollection(Me.ObjSystem.Engines, Me, Me.GUID) = False Then
                tran.Rollback()
                AddNew = False
                Exit Try
            End If
            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste this flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As clsDatastore In Me.Sources
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsDatastore In Me.Targets
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsVariable In Me.Variables
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
                'For Each child As clsTask In Me.Gens
                '    If child.AddNew(cmd, True) = False Then
                '        tran.Rollback()
                '        AddNew = False
                '        Exit Try
                '    End If
                'Next
                For Each child As clsTask In Me.Procs
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsTask In Me.Mains
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsTask In Me.Lookups
                    If child.AddNew(cmd, True) = False Then
                        tran.Rollback()
                        AddNew = False
                        Exit Try
                    End If
                Next
            End If

            tran.Commit()
            IsModified = False
            AddNew = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsEngine AddNew", sql)
            tran.Rollback()
            AddNew = False
        Catch ex As Exception
            LogError(ex, "clsEngine AddNew", sql)
            tran.Rollback()
            AddNew = False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// function added by KS and TK 11/3/06
    '/// OverHauled by TKarasch    April,May 07
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            '//When we add new record we need to find Unique Database Record ID.
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    '//When we add new record we need to find Unique Database Record ID.
            '    sql = "INSERT INTO " & Me.Project.tblEngines & _
            '    "(ProjectName,EnvironmentName,SystemName,EngineName,ReportEvery,CommitEvery,ReportFile,Description,CopybookLib,IncludeLib,DTDLib,DDLLib,ConnectionName) " & _
            '          " Values(" & Me.Project.GetQuotedText & "," & Me.ObjSystem.Environment.GetQuotedText & "," & _
            '          Me.ObjSystem.GetQuotedText & "," & Me.GetQuotedText & "," & Me.ReportEvery & "," & Me.CommitEvery & ",'" & _
            '          FixStr(Me.ReportFile) & "','" & FixStr(EngineDescription) & "','" & FixStr(CopybookLib) & "','" & _
            '          FixStr(IncludeLib) & "','" & FixStr(DTDLib) & "','" & FixStr(DDLLib) & "','" & Me.Connection.Text & "')"
            'Else
            '//When we add new record we need to find Unique Database Record ID.
            sql = "INSERT INTO " & Me.Project.tblEngines & "(ProjectName,EnvironmentName,SystemName,EngineName,EngineDescription) " & _
            " Values(" & Me.Project.GetQuotedText & "," & _
            Me.ObjSystem.Environment.GetQuotedText & "," & _
            Me.ObjSystem.GetQuotedText & "," & _
            Me.GetQuotedText & ",'" & _
            FixStr(Me.EngineDescription) & "')"


            'End If
            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If Me.DeleteATTR(cmd) = False Then
                AddNew = False
                Exit Try
            End If
            If Me.InsertATTR(cmd) = False Then
                AddNew = False
                Exit Try
            End If

            AddToCollection(Me.ObjSystem.Engines, Me, Me.GUID)

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As clsDatastore In Me.Sources
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsDatastore In Me.Targets
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsVariable In Me.Variables
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
                'For Each child As clsTask In Me.Gens
                '    If child.AddNew(cmd, True) = False Then
                '        AddNew = False
                '        Exit Try
                '    End If
                'Next
                For Each child As clsTask In Me.Procs
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsTask In Me.Mains
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
                For Each child As clsTask In Me.Lookups
                    If child.AddNew(cmd, True) = False Then
                        AddNew = False
                        Exit Try
                    End If
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsEngine AddNew", sql)
            AddNew = False
        Catch ex As Exception
            LogError(ex, "clsEngine AddNew", sql)
            AddNew = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then Exit Try
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
            Dim cmd As System.Data.Odbc.OdbcCommand
            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As New DataTable("temp")
            Dim dr As DataRow
            Dim sql As String = ""
            'Dim strAttrs As String
            Dim Attrib As String = ""
            Dim Value As String = ""

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "SELECT ENGINEATTRB,ENGINEATTRBVALUE FROM " & Me.Project.tblEnginesATTR & _
            " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & _
            "' AND ENVIRONMENTNAME='" & FixStr(Me.ObjSystem.Environment.EnvironmentName) & _
            "' AND SYSTEMNAME='" & FixStr(Me.ObjSystem.SystemName) & _
            "' AND ENGINENAME='" & FixStr(Me.EngineName) & "'"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetVal(dr("ENGINEATTRB"))
                Select Case Attrib
                    Case "COMMITEVERY"
                        Me.CommitEvery = GetVal(dr("ENGINEATTRBVALUE"))
                    Case "CONNECTIONNAME"
                        Me.Connection = SearchForConn(Me, GetVal(dr("ENGINEATTRBVALUE")))
                    Case "COPYBOOKLIB"
                        Me.CopybookLib = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "DATEFORMAT"
                        Me.DateFormat = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "DBDLIB"
                        Me.DBDLib = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "DDLLIB"
                        Me.DDLLib = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "DTDLIB"
                        Me.DTDLib = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "FORCECOMMIT"
                        Me.ForceCommit = GetVal(dr("ENGINEATTRBVALUE"))
                    Case "INCLUDELIB"
                        Me.IncludeLib = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "REPORTEVERY"
                        Me.ReportEvery = GetVal(dr("ENGINEATTRBVALUE"))
                    Case "REPORTFILE"
                        Me.ReportFile = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "MAIN"
                        Me.Main = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "ENGVER"
                        Me.EngVersion = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "EXEDIR"
                        Me.EXEdir = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "BATDIR"
                        Me.BATdir = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "CDCDIR"
                        Me.CDCdir = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                    Case "ADDFLOWPATH"
                        Me.CurAddFlowDiagPath = GetStr(GetVal(dr("ENGINEATTRBVALUE")))
                End Select
            Next

            Me.IsDiagramChanged = False
            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsEngine LoadMe")
            Return False
        End Try

    End Function

    Public Property IsRenamed() As Boolean Implements INode.IsRenamed
        Get
            Return m_IsRenamed
        End Get
        Set(ByVal Value As Boolean)
            m_IsRenamed = Value
        End Set
    End Property

    Property IsLoaded() As Boolean Implements INode.IsLoaded
        Get
            Return m_IsLoaded
        End Get
        Set(ByVal value As Boolean)
            m_IsLoaded = value
        End Set
    End Property

    Public Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean Implements INode.ValidateNewObject

        'Dim cnn As New System.Data.Odbc.OdbcConnection(Me.Project.MetaConnectionString)
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim dr As System.Data.Odbc.OdbcDataReader
        Dim sql As String = ""
        Dim readName As String
        Dim testName As String = NewName
        Dim NameValid As Boolean = True

        Try
            If testName = "" Then
                testName = Me.Text
            End If

            'cnn.Open()
            cmd = cnn.CreateCommand

            sql = "Select ENGINENAME from " & Me.Project.tblEngines & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.ObjSystem.Environment.EnvironmentName & "' AND SYSTEMNAME='" & Me.ObjSystem.SystemName & "'"

            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("ENGINENAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The New Engine name you have chosen already exists" & (Chr(13)) & "in this System, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsEngine ValidateNewObject")
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property EngineName() As String
        Get
            Return m_EngineName
        End Get
        Set(ByVal Value As String)
            m_EngineName = Value
        End Set
    End Property

    Public Property ReportFile() As String
        Get
            Return m_ReportFile
        End Get
        Set(ByVal Value As String)
            m_ReportFile = Value
        End Set
    End Property

    Public Property ReportEvery() As Integer
        Get
            Return m_ReportEvery
        End Get
        Set(ByVal Value As Integer)
            m_ReportEvery = Value
        End Set
    End Property

    Public Property CommitEvery() As Integer
        Get
            Return m_CommitEvery
        End Get
        Set(ByVal Value As Integer)
            m_CommitEvery = Value
        End Set
    End Property

    Public Property ForceCommit() As Boolean
        Get
            Return m_ForceCommit
        End Get
        Set(ByVal value As Boolean)
            m_ForceCommit = value
        End Set
    End Property

    Public Property EngineDescription() As String
        Get
            Return m_EngineDescription
        End Get
        Set(ByVal Value As String)
            m_EngineDescription = Value
        End Set
    End Property

    Public Property ObjSystem() As clsSystem
        Get
            Return m_ObjSystem
        End Get
        Set(ByVal Value As clsSystem)
            m_ObjSystem = Value
        End Set
    End Property

    Public Property MapGroupItems() As Boolean
        Get
            Return m_MapGroupItems
        End Get
        Set(ByVal value As Boolean)
            m_MapGroupItems = value
        End Set
    End Property

    Public Property Connection() As clsConnection
        Get
            Return m_Connection
        End Get
        Set(ByVal value As clsConnection)
            m_Connection = value
        End Set
    End Property

    Public Property DateFormat() As String
        Get
            Return m_DateFormat
        End Get
        Set(ByVal value As String)
            m_DateFormat = value
        End Set
    End Property

    Public Property Main() As String
        Get
            Return m_Main
        End Get
        Set(ByVal value As String)
            m_Main = value
        End Set
    End Property

    Public Property EngVersion() As String
        Get
            Return m_EngVersion
        End Get
        Set(ByVal value As String)
            m_EngVersion = value
        End Set
    End Property

    Public Property ObjAddFlow() As AddFlow
        Get
            Return m_ObjAddFlow
        End Get
        Set(ByVal value As AddFlow)
            m_ObjAddFlow = value
        End Set
    End Property

    Public Property ObjTabPage() As TabPage
        Get
            Return m_ObjTabPage
        End Get
        Set(ByVal value As TabPage)
            m_ObjTabPage = value
        End Set
    End Property

    Public Property ObjAddFlowCtl() As ctlAddFlowTab
        Get
            Return m_ObjAddFlowCtl
        End Get
        Set(ByVal value As ctlAddFlowTab)
            m_ObjAddFlowCtl = value
        End Set
    End Property

    Public Property CurAddFlowDiagPath() As String
        Get
            If m_CurAddFlowDiagPath = "" Then
                m_CurAddFlowDiagPath = GetAppProj() & _
                Me.Project.ProjectName & "." & _
                Me.ObjSystem.Environment.EnvironmentName & "." & _
                Me.ObjSystem.SystemName & "." & _
                Me.EngineName & ".current.xml"
            End If
            Return m_CurAddFlowDiagPath
        End Get
        Set(ByVal value As String)
            m_CurAddFlowDiagPath = value
        End Set
    End Property

    Public Property IsDiagramChanged() As Boolean
        Get
            Return m_IsDiagramChanged
        End Get
        Set(ByVal value As Boolean)
            m_IsDiagramChanged = value
        End Set
    End Property

#End Region

#Region "Methods"

    '/// added 5/07 by Tom Karasch to make sure no duplicate names are used for
    '/// Tasks, datastores or variables in any given Engine
    '/// it uses the object tree instead of going to MetaData for speed
    Function FindDupNames(ByVal NewObj As INode) As Boolean

        Dim Duplicate As Boolean = False

        Try
            Select Case NewObj.Type
                Case modDeclares.NODE_TARGETDATASTORE
                    For Each TgtDS As clsDatastore In Me.Targets
                        If TgtDS.Text = NewObj.Text Then
                            Duplicate = True
                            Exit Select
                        End If
                    Next
                Case modDeclares.NODE_PROC
                    For Each Proc As clsTask In Me.Procs
                        If Proc.Text = NewObj.Text Then
                            Duplicate = True
                            Exit Select
                        End If
                    Next
                Case modDeclares.NODE_MAIN
                    For Each Main As clsTask In Me.Mains
                        If Main.Text = NewObj.Text Then
                            Duplicate = True
                            Exit Select
                        End If
                    Next
            End Select

            Return Duplicate

        Catch ex As Exception
            LogError(ex, "clsEngine findDupNames")
            Return False
        End Try

    End Function

    Function InsertATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As Odbc.OdbcCommand
        Dim sql As String = ""
        Dim Attrib As String = ""
        Dim Value As String = ""

        Try
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            For i As Integer = 0 To 16
                Select Case i
                    Case 0
                        Attrib = "COMMITEVERY"
                        Value = Me.CommitEvery
                    Case 1
                        Attrib = "CONNECTIONNAME"
                        If Me.Connection IsNot Nothing Then
                            Value = Me.Connection.ConnectionName
                        Else
                            Value = ""
                        End If
                    Case 2
                        Attrib = "COPYBOOKLIB"
                        Value = Me.CopybookLib
                    Case 3
                        Attrib = "DATEFORMAT"
                        Value = Me.DateFormat
                    Case 4
                        Attrib = "DDLLIB"
                        Value = Me.DDLLib
                    Case 5
                        Attrib = "DTDLIB"
                        Value = Me.DTDLib
                    Case 6
                        Attrib = "FORCECOMMIT"
                        Value = Me.ForceCommit
                    Case 7
                        Attrib = "INCLUDELIB"
                        Value = Me.IncludeLib
                    Case 8
                        Attrib = "REPORTEVERY"
                        Value = Me.ReportEvery
                    Case 9
                        Attrib = "REPORTFILE"
                        Value = Me.ReportFile
                    Case 10
                        Attrib = "MAIN"
                        Value = Me.Main
                    Case 11
                        Attrib = "ENGVER"
                        Value = Me.EngVersion
                    Case 12
                        Attrib = "DBDLIB"
                        Value = Me.DBDLib
                    Case 13
                        Attrib = "EXEDIR"
                        Value = Me.EXEdir
                    Case 14
                        Attrib = "BATDIR"
                        Value = Me.BATdir
                    Case 15
                        Attrib = "CDCDIR"
                        Value = Me.CDCdir
                    Case 16
                        Attrib = "ADDFLOWPATH"
                        Value = Me.CurAddFlowDiagPath
                End Select

                sql = "INSERT INTO " & Me.Project.tblEnginesATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,ENGINEATTRB,ENGINEATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & "','" & _
                FixStr(Me.ObjSystem.Environment.EnvironmentName) & "','" & _
                FixStr(Me.ObjSystem.SystemName) & "','" & _
                FixStr(Me.EngineName) & "','" & _
                FixStr(Attrib) & "','" & _
                FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            InsertATTR = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsEngine InsertATTR", sql)
            InsertATTR = False
        Catch ex As Exception
            LogError(ex, "clsEngine InsertATTR", sql)
            InsertATTR = False
        End Try

    End Function

    Function DeleteATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As Odbc.OdbcCommand
        Dim sql As String = ""
        Dim Attrib As String = ""
        Dim Value As String = ""

        Try
            If INcmd IsNot Nothing Then
                cmd = INcmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "DELETE FROM " & Me.Project.tblEnginesATTR & _
            " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.ObjSystem.Environment.EnvironmentName) & _
            " AND SYSTEMNAME=" & Quote(Me.ObjSystem.SystemName) & _
            " AND ENGINENAME=" & Quote(Me.EngineName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            DeleteATTR = True

        Catch OE As Odbc.OdbcException
            LogODBCError(OE, "clsEngine InsertATTR", sql)
            DeleteATTR = False
        Catch ex As Exception
            LogError(ex, "clsEngine DeleteATTR", sql)
            DeleteATTR = False
        End Try

    End Function

    Function SaveAddFlowXML() As Boolean

        Try
            If IsDiagramChanged = False Then
                Return True
                Exit Function
            End If

            Me.ObjAddFlowCtl.SaveFile(Me.CurAddFlowDiagPath)

            Return True

        Catch ex As Exception
            LogError(ex, "clsEngine SaveAddFlow")
            Return False
        End Try

    End Function

    Function LoadAddFlowXML() As Boolean

        Try
            If IsDiagramChanged = False Then
                Return True
                Exit Function
            End If

            '/// Load Addflow Diagram from File if it didn't change
            Me.ObjAddFlowCtl.LoadFile(Me.CurAddFlowDiagPath)

            Return True

        Catch ex As Exception
            LogError(ex, "clsEngine LoadAddFlow")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class