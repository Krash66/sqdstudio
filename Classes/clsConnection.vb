<Serializable()> _
Public Class clsConnection
    Implements INode

    Private m_ConnectionName As String = ""
    Private m_ConnectionType As String = ""
    Private m_Database As String = ""
    Private m_UserId As String = ""
    Private m_Password As String = ""
    Private m_DateFormat As String = ""
    Private m_ConnectionDescription As String = ""
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_Environment As clsEnvironment
    Private m_IsLoaded As Boolean = False


#Region "INode Implementation"

    ''' <summary>
    ''' Inserts single quotes for ease of insersion in to SQL statements
    ''' </summary>
    ''' <param name="bFix">Boolean var to fix string before returning value</param>
    ''' <returns>Text as single quoted Text Value of Inode</returns>
    ''' <remarks></remarks>
    ''' 
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
            Return Me.Environment
        End Get
        Set(ByVal Value As INode)
            m_Environment = Value
        End Set
    End Property
    '
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone
        '/// Fix by TKarasch May 07
        Try
            Dim obj As New clsConnection
            Dim cmd As New Odbc.OdbcCommand

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)

            'obj.ObjTreeNode = Me.ObjTreeNode
            'obj.ConnectionId = Me.ConnectionId
            obj.ConnectionName = Me.ConnectionName
            obj.ConnectionDescription = Me.ConnectionDescription
            obj.Database = Me.Database
            obj.DateFormat = Me.DateFormat
            obj.UserId = Me.UserId
            obj.Password = Me.Password
            obj.SeqNo = Me.SeqNo
            obj.ConnectionType = Me.ConnectionType

            obj.IsModified = Me.IsModified
            obj.Parent = NewParent 'Me.Parent

            If Cascade = True Then
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsConnection Clone")
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
            'Key = ConnectionId
            Key = Me.Environment.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
        End Get
    End Property

    Public Property Text() As String Implements INode.Text
        Get
            Text = ConnectionName
        End Get
        Set(ByVal Value As String)
            ConnectionName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_CONNECTION
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

        Dim sql As String = ""
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New System.Data.Odbc.OdbcCommand

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblConnections & " set ConnectionName=" & Me.GetQuotedText & " , UserId='" & FixStr(Me.UserId) & "' , DBName='" & FixStr(Me.Database) & "' , Password='" & FixStr(Me.Password) & "', ConnectionType='" & FixStr(Me.ConnectionType) & "' , DateFormat='" & FixStr(Me.DateFormat) & "', Description='" & FixStr(Me.ConnectionDescription) & "' where ConnectionName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            'Else
            sql = "Update " & Me.Project.tblConnections & _
            " set ConnectionName=" & Me.GetQuotedText & _
            ", ConnectionDescription='" & FixStr(Me.ConnectionDescription) & _
            "' where ConnectionName=" & Me.GetQuotedText & _
            " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
            " AND ProjectName=" & Me.Project.GetQuotedText

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)

            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Me.IsModified = False

            Save = True

        Catch ex As Exception
            LogError(ex)
            Save = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            If Cascade = True Then
            End If

            If Me.Project.ProjectMetaVersion = enumMetaVersion.V3 Then
                Me.DeleteATTR(cmd)
            End If

            sql = "Delete From " & Me.Project.tblConnections & " where ConnectionName=" & Me.GetQuotedText & " AND EnvironmentName=" & Me.Environment.GetQuotedText & " AND ProjectName=" & Me.Environment.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                RemoveFromCollection(Me.Environment.Connections, Me.GUID)
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsConnection Delete")
            Delete = False
        End Try

    End Function

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
            '    sql = "INSERT INTO " & Me.Project.tblConnections & _
            '                "(ProjectName,EnvironmentName,ConnectionName,ConnectionType,UserId,DBName,Password,Description,DateFormat) " & _
            '                " Values(" & Me.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.GetQuotedText & ",'" & _
            '                FixStr(ConnectionType) & "','" & FixStr(Me.UserId) & "','" & FixStr(Me.Database) & "','" & FixStr(Me.Password) & _
            '                "','" & FixStr(ConnectionDescription) & "','" & FixStr(Me.DateFormat) & "')"
            'Else
            sql = "INSERT INTO " & Me.Project.tblConnections & _
                        "(ProjectName,EnvironmentName,ConnectionName,ConnectionDescription) " & _
                        " Values(" & Me.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.GetQuotedText & ",'" & _
                        FixStr(ConnectionDescription) & "')"

            Me.InsertATTR(cmd)

            'End If


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Add connection into systems's connection collection
            AddToCollection(Me.Environment.Connections, Me, Me.GUID)

            AddNew = True
            IsModified = False

        Catch ex As Exception
            LogError(ex, "clsConnection AddNew")
            AddNew = False
        Finally
            'cnn.Close()
        End Try

    End Function

    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            '//When we add new record we need to find Unique Database Record ID.
            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "INSERT INTO " & Me.Project.tblConnections & _
            '                "(ProjectName,EnvironmentName,ConnectionName,ConnectionType,UserId,DBName,Password,Description,DateFormat) " & _
            '                " Values(" & Me.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.GetQuotedText & ",'" & _
            '                FixStr(ConnectionType) & "','" & FixStr(Me.UserId) & "','" & FixStr(Me.Database) & "','" & FixStr(Me.Password) & _
            '                "','" & FixStr(ConnectionDescription) & "','" & FixStr(Me.DateFormat) & "')"
            'Else
            sql = "INSERT INTO " & Me.Project.tblConnections & _
                        "(ProjectName,EnvironmentName,ConnectionName,ConnectionDescription) " & _
                        " Values(" & Me.Project.GetQuotedText & "," & Me.Environment.GetQuotedText & "," & Me.GetQuotedText & ",'" & _
                        FixStr(ConnectionDescription) & "')"

            Me.InsertATTR(cmd)
            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Add connection into systems's connection collection
            AddToCollection(Me.Environment.Connections, Me, Me.GUID)

            AddNew = True
            IsModified = False

        Catch ex As Exception
            LogError(ex, "clsConnection AddNew")
            AddNew = False
        End Try

    End Function

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Try
            If Me.IsLoaded = True Then Exit Try

            Dim cmd As New System.Data.Odbc.OdbcCommand
            Dim da As System.Data.Odbc.OdbcDataAdapter
            Dim dt As New DataTable("temp")
            Dim dr As DataRow
            Dim sql As String = ""
            Dim Attrib As String = ""
            Dim Value As String = ""

            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            sql = "SELECT CONNECTIONATTRB,CONNECTIONATTRBVALUE FROM " & Me.Project.tblConnectionsATTR & _
            " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & "' AND ENVIRONMENTNAME='" & _
            FixStr(Me.Environment.EnvironmentName) & "' AND CONNECTIONNAME='" & FixStr(Me.ConnectionName) & "'"

            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetStr(GetVal(dr("CONNECTIONATTRB")))
                Select Case Attrib
                    Case "CONNECTIONTYPE"
                        Me.ConnectionType = GetStr(GetVal(dr("CONNECTIONATTRBVALUE")))
                    Case "DATEFORMAT"
                        Me.DateFormat = GetStr(GetVal(dr("CONNECTIONATTRBVALUE")))
                    Case "DBNAME"
                        Me.Database = GetStr(GetVal(dr("CONNECTIONATTRBVALUE")))
                    Case "PASSWORD"
                        Me.Password = GetStr(GetVal(dr("CONNECTIONATTRBVALUE")))
                    Case "USERID"
                        Me.UserId = GetStr(GetVal(dr("CONNECTIONATTRBVALUE")))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsConnection LoadMe")
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

            sql = "Select CONNECTIONNAME from " & Me.Project.tblConnections & " where PROJECTNAME='" & Me.Project.ProjectName & "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & "'"

            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("CONNECTIONNAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The New connection name you have chosen already exists" & (Chr(13)) & "in this system, Please enter a different Name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsConnection ValidateNewObject")
            ValidateNewObject = False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

#Region "Properties"

    Public Property ConnectionName() As String
        Get
            Return m_ConnectionName
        End Get
        Set(ByVal Value As String)
            m_ConnectionName = Value
        End Set
    End Property

    Public Property DateFormat() As String
        Get
            Return m_DateFormat
        End Get
        Set(ByVal Value As String)
            m_DateFormat = Value
        End Set
    End Property

    Public Property ConnectionType() As String
        Get
            Return m_ConnectionType
        End Get
        Set(ByVal Value As String)
            m_ConnectionType = Value
        End Set
    End Property

    Public Property Database() As String
        Get
            Return m_Database
        End Get
        Set(ByVal Value As String)
            m_Database = Value
        End Set
    End Property

    Public Property UserId() As String
        Get
            Return m_UserId
        End Get
        Set(ByVal Value As String)
            m_UserId = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return m_Password
        End Get
        Set(ByVal Value As String)
            m_Password = Value
        End Set
    End Property

    Public Property ConnectionDescription() As String
        Get
            Return m_ConnectionDescription
        End Get
        Set(ByVal Value As String)
            m_ConnectionDescription = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            Return m_Environment
        End Get
        Set(ByVal value As clsEnvironment)
            m_Environment = value
        End Set
    End Property

#End Region

#Region "methods"

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

            'cmd.Connection = cnn

            For i As Integer = 0 To 4
                Select Case i
                    Case 0
                        Attrib = "CONNECTIONTYPE"
                        Value = Me.ConnectionType
                    Case 1
                        Attrib = "DATEFORMAT"
                        Value = Me.DateFormat
                    Case 2
                        Attrib = "DBNAME"
                        Value = Me.Database
                    Case 3
                        Attrib = "PASSWORD"
                        Value = Me.Password
                    Case 4
                        Attrib = "USERID"
                        Value = Me.UserId
                End Select

                sql = "INSERT INTO " & Me.Project.tblConnectionsATTR & _
                "(PROJECTNAME,ENVIRONMENTNAME,CONNECTIONNAME,CONNECTIONATTRB,CONNECTIONATTRBVALUE) " & _
                "Values('" & FixStr(Me.Project.ProjectName) & "','" & FixStr(Me.Environment.EnvironmentName) & _
                "','" & FixStr(Me.ConnectionName) & "','" & FixStr(Attrib) & "','" & FixStr(Value) & "')"


                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsConnection InsertATTR")
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

            'cmd.Connection = cnn

            sql = "DELETE FROM " & Me.Project.tblConnectionsATTR & " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
            " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & " AND CONNECTIONNAME=" & Quote(Me.ConnectionName)

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsConnection DeleteATTR")
            Return False
        End Try

    End Function


#End Region

    Sub New()
        m_GUID = GetNewId()
    End Sub

End Class