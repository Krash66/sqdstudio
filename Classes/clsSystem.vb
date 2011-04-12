Public Class clsSystem
    Implements INode

    Private m_SystemName As String = ""
    Private m_OSType As String = ""
    Private m_Port As String = ""
    Private m_Host As String = ""
    Private m_QueMgr As String = ""
    Private m_SystemDescription As String = ""
    Private m_IsModified As Boolean
    Private m_ObjEnvironment As clsEnvironment
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_IsLoaded As Boolean = False

    Public DBDLib As String = ""
    Public CopybookLib As String = ""
    Public IncludeLib As String = ""
    Public DTDLib As String = ""
    Public DDLLib As String = ""

    Public Engines As New Collection

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
            Return Me.m_ObjEnvironment
        End Get
        Set(ByVal Value As INode)
            m_ObjEnvironment = Value
        End Set
    End Property

    '/// OverHauled by TKarasch  April May 07
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsSystem
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)

            obj.SystemName = Me.SystemName
            obj.SystemDescription = Me.SystemDescription
            obj.OSType = Me.OSType
            obj.Port = Me.Port
            obj.Host = Me.Host
            obj.QueueManager = Me.QueueManager
            obj.CopybookLib = Me.CopybookLib
            obj.IncludeLib = Me.IncludeLib
            obj.DTDLib = Me.DTDLib
            obj.DDLLib = Me.DDLLib
            obj.SeqNo = Me.SeqNo
            obj.IsModified = Me.IsModified
            obj.Parent = NewParent 'Me.Parent

            If Cascade = True Then
                Dim NewEng As clsEngine
                '//clone all engines under system
                For Each eng As clsEngine In Me.Engines
                    eng.LoadMe(cmd)

                    NewEng = eng.Clone(obj, True, cmd)
                    NewEng.ObjSystem = obj
                    AddToCollection(obj.Engines, NewEng, NewEng.GUID)
                Next
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsSystem Clone")
            Return Nothing
        End Try

    End Function

    Public ReadOnly Property Key() As String Implements INode.Key
        Get
            Key = Me.Environment.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = SystemName
        End Get
        Set(ByVal Value As String)
            SystemName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_SYSTEM
        End Get
    End Property

    Public ReadOnly Property Project() As clsProject Implements INode.Project
        Get
            Dim p As INode
            p = Me.Parent.Project
            Return p
        End Get
    End Property

    Public Property IsModified() As Boolean Implements INode.IsModified
        Get
            Return m_IsModified
        End Get
        Set(ByVal Value As Boolean)
            m_IsModified = Value
        End Set
    End Property

    Public Function Save(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.Save

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblSystems & " set SystemName='" & FixStr(Me.SystemName) & "'" & _
            '                    " , Host='" & FixStr(Me.Host) & "'" & _
            '                    " , Port='" & FixStr(Me.Port) & "'" & _
            '                    " , QMgr='" & FixStr(Me.QueueManager) & "'" & _
            '                    " , CopybookLib='" & FixStr(Me.CopybookLib) & "'" & _
            '                    " , IncludeLib='" & FixStr(Me.IncludeLib) & "'" & _
            '                    " , DTDLib='" & FixStr(Me.DTDLib) & "'" & _
            '                    " , DDLLib='" & FixStr(Me.DDLLib) & "'" & _
            '                    " , Description='" & FixStr(Me.SystemDescription) & "'" & _
            '                    " , OSType='" & FixStr(Me.OSType) & "' " & _
            '                    " where SystemName=" & Me.GetQuotedText & " AND EnvironmentName = '" & _
            '                    Me.Environment.Text & "'" & " And ProjectName = '" & Me.Environment.Project.Text & "'"
            'Else
            sql = "Update " & Me.Project.tblSystems & " set SystemName='" & FixStr(Me.SystemName) & "'" & _
                            " , SystemDescription='" & FixStr(Me.SystemDescription) & "'" & _
                            " where SystemName=" & Me.GetQuotedText & " AND EnvironmentName = '" & _
                            Me.Environment.Text & "'" & " And ProjectName = '" & Me.Environment.Project.Text & "'"

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)
            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Me.IsModified = False
            Save = True

        Catch ex As Exception
            LogError(ex, "clsSystem Save", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '/// OverHauled by TKarasch   April May 07
    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "INSERT INTO " & Me.Project.tblSystems & "(ProjectName " & _
            '               ",EnvironmentName" & _
            '               ",SystemName" & _
            '               ",Host" & _
            '               ",Port" & _
            '               ",QMgr" & _
            '               ",CopybookLib" & _
            '               ",IncludeLib" & _
            '               ",DTDLib" & _
            '               ",DDLLib" & _
            '               ",Description" & _
            '               ",OSType) " & _
            '         " Values(" & Me.Environment.Parent.GetQuotedText & _
            '               "," & Me.Environment.GetQuotedText & _
            '               "," & Me.GetQuotedText & _
            '               ",'" & FixStr(Me.Host) & _
            '               "','" & FixStr(Me.Port) & _
            '               "','" & FixStr(Me.QueueManager) & _
            '               "','" & FixStr(Me.CopybookLib) & _
            '               "','" & FixStr(Me.IncludeLib) & _
            '               "','" & FixStr(Me.DTDLib) & _
            '               "','" & FixStr(Me.DDLLib) & _
            '               "','" & FixStr(SystemDescription) & _
            '               "','" & FixStr(Me.OSType) & "')"
            'Else
            sql = "INSERT INTO " & Me.Project.tblSystems & "(ProjectName " & _
                       ",EnvironmentName" & _
                       ",SystemName" & _
                       ",SystemDescription) " & _
                 " Values(" & Me.Project.GetQuotedText & _
                       "," & Me.Environment.GetQuotedText & _
                       "," & Me.GetQuotedText & _
                       ",'" & FixStr(SystemDescription) & "')"

            Me.InsertATTR(cmd)
            'end If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            AddToCollection(Me.Environment.Systems, Me, Me.GUID)

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As clsEngine In Me.Engines
                    child.AddNew(True)
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsSystem AddNew", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// Added by TK and KS 11/6/2006
    '/// OverHauled by TKarasch   April May 07
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "INSERT INTO " & Me.Project.tblSystems & "(ProjectName " & _
            '               ",EnvironmentName" & _
            '               ",SystemName" & _
            '               ",Host" & _
            '               ",Port" & _
            '               ",QMgr" & _
            '               ",CopybookLib" & _
            '               ",IncludeLib" & _
            '               ",DTDLib" & _
            '               ",DDLLib" & _
            '               ",Description" & _
            '               ",OSType) " & _
            '         " Values(" & Me.Environment.Parent.GetQuotedText & _
            '               "," & Me.Environment.GetQuotedText & _
            '               "," & Me.GetQuotedText & _
            '               ",'" & FixStr(Me.Host) & _
            '               "','" & FixStr(Me.Port) & _
            '               "','" & FixStr(Me.QueueManager) & _
            '               "','" & FixStr(Me.CopybookLib) & _
            '               "','" & FixStr(Me.IncludeLib) & _
            '               "','" & FixStr(Me.DTDLib) & _
            '               "','" & FixStr(Me.DDLLib) & _
            '               "','" & FixStr(SystemDescription) & _
            '               "','" & FixStr(Me.OSType) & "')"
            'Else
            sql = "INSERT INTO " & Me.Project.tblSystems & "(ProjectName,EnvironmentName,SystemName,SystemDescription) " & _
                 " Values(" & Me.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.GetQuotedText & _
                 ",'" & FixStr(SystemDescription) & "')"

            Me.InsertATTR(cmd)
            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            AddToCollection(Me.Environment.Systems, Me, Me.GUID)

            '//Add all child object to database if Cascade flag is true. 
            '//Generally when performing Clipboard copy/paste thi flag is set so 
            '//entire copied object tree can be added to database by just calling 
            '//AddNew method of parent node
            If Cascade = True Then
                For Each child As INode In Engines
                    child.AddNew(True)
                Next
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsSystem AddNew", sql)
            Return False
        End Try

    End Function

    '/// OverHauled by TKarasch   April May 07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            If Cascade = True Then
                Dim o As INode
                For Each o In Me.Engines
                    o.Delete(cmd, cnn, Cascade)
                Next
                RemoveFromCollection(Me.Engines, "") '//clear collection
            End If

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                Me.DeleteATTR(cmd)
            End If

            sql = "Delete From " & Me.Project.tblSystems & " where SystemName=" & Me.GetQuotedText & " AND EnvironmentName = " & Me.Environment.GetQuotedText & " And ProjectName = " & Me.Environment.Project.GetQuotedText

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                RemoveFromCollection(Me.Environment.Systems, Me.GUID)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsSystem Delete", sql)
            Return False
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


            sql = "SELECT SYSTEMATTRB,SYSTEMATTRBVALUE FROM " & Me.Project.tblSystemsATTR & _
            " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & _
            "' AND ENVIRONMENTNAME='" & FixStr(Me.Environment.EnvironmentName) & _
            "' AND SystemNAME='" & FixStr(Me.SystemName) & "'"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetStr(GetVal(dr("SYSTEMATTRB")))
                Select Case Attrib
                    Case "DBDLIB"
                        Me.DBDLib = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "COPYBOOKLIB"
                        Me.CopybookLib = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "DDLLIB"
                        Me.DDLLib = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "DTDLIB"
                        Me.DTDLib = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "HOST"
                        Me.Host = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "INCLUDELIB"
                        Me.IncludeLib = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "OSTYPE"
                        Me.OSType = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "PORT"
                        Me.Port = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                    Case "QMGR"
                        Me.QueueManager = GetStr(GetVal(dr("SYSTEMATTRBVALUE")))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsSystem LoadMe")
            Return False
        End Try

    End Function

    Property IsLoaded() As Boolean Implements INode.IsLoaded
        Get
            Return m_IsLoaded
        End Get
        Set(ByVal value As Boolean)
            m_IsLoaded = value
        End Set
    End Property

    Public Property IsRenamed() As Boolean Implements INode.IsRenamed
        Get
            Return m_IsRenamed
        End Get
        Set(ByVal Value As Boolean)
            m_IsRenamed = Value
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

            sql = "Select * from " & Me.Project.tblSystems & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & "'"

            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("SYSTEMNAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The New Name you have chosen already exists," & (Chr(13)) & "Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsSystem ValidateNewObject", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property SystemName() As String
        Get
            Return m_SystemName
        End Get
        Set(ByVal Value As String)
            m_SystemName = Value
        End Set
    End Property

    Public Property OSType() As String
        Get
            Return m_OSType
        End Get
        Set(ByVal Value As String)
            m_OSType = Value
        End Set
    End Property

    Public Property Port() As String
        Get
            Return m_Port
        End Get
        Set(ByVal Value As String)
            m_Port = Value
        End Set
    End Property

    Public Property Host() As String
        Get
            Return m_Host
        End Get
        Set(ByVal Value As String)
            m_Host = Value
        End Set
    End Property

    Public Property QueueManager() As String
        Get
            Return m_QueMgr
        End Get
        Set(ByVal Value As String)
            m_QueMgr = Value
        End Set
    End Property

    Public Property SystemDescription() As String
        Get
            Return m_SystemDescription
        End Get
        Set(ByVal Value As String)
            m_SystemDescription = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            Return m_ObjEnvironment
        End Get
        Set(ByVal Value As clsEnvironment)
            m_ObjEnvironment = Value
        End Set
    End Property

#End Region

#Region "Methods"

    Function InsertATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As New Odbc.OdbcCommand
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

            cmd.Connection = cnn

            For i As Integer = 0 To 8
                Select Case i
                    Case 0
                        Attrib = "COPYBOOKLIB"
                        Value = Me.CopybookLib
                    Case 1
                        Attrib = "DDLLIB"
                        Value = Me.DDLLib
                    Case 2
                        Attrib = "DTDLIB"
                        Value = Me.DTDLib
                    Case 3
                        Attrib = "INCLUDELIB"
                        Value = Me.IncludeLib
                    Case 4
                        Attrib = "HOST"
                        Value = Me.Host
                    Case 5
                        Attrib = "OSTYPE"
                        Value = Me.OSType
                    Case 6
                        Attrib = "PORT"
                        Value = Me.Port
                    Case 7
                        Attrib = "QMGR"
                        Value = Me.QueueManager
                    Case 8
                        Attrib = "DBDLIB"
                        Value = Me.DBDLib
                End Select

                sql = "INSERT INTO " & Me.Project.tblSystemsATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,SYSTEMATTRB,SYSTEMATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & "','" & FixStr(Me.Environment.EnvironmentName) & _
                "','" & FixStr(Me.SystemName) & "','" & FixStr(Attrib) & "','" & FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsSystem InsertATTR")
            Return False
        End Try

    End Function

    Function DeleteATTR(Optional ByRef INcmd As System.Data.Odbc.OdbcCommand = Nothing) As Boolean

        Dim cmd As New Odbc.OdbcCommand
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

            cmd.Connection = cnn

            sql = "DELETE FROM " & Me.Project.tblSystemsATTR & _
            " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
            " AND SYSTEMNAME=" & Quote(Me.SystemName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsSystem DeleteATTR")
            Return False
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class