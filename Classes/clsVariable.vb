
Public Class clsVariable
    Implements INode

    Private m_VariableType As enumVariableType
    Private m_VariableName As String = ""
    Private m_CorrectedVariableName As String = ""
    Private m_VarSize As Integer = 0
    Private m_VarInitVal As String = ""
    Private m_VariableDescription As String = ""
    Private m_ObjEngine As clsEngine
    Private m_IsModified As Boolean
    Private m_ObjTreeNode As TreeNode
    Private m_ObjParent As INode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_IsRenamed As Boolean = False
    Private m_Environment As clsEnvironment
    Private m_IsLoaded As Boolean = False


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
            If m_ObjParent Is Nothing Then
                If Me.m_ObjEngine Is Nothing Then
                    Return Me.m_Environment
                Else
                    Return Me.m_ObjEngine
                End If
            Else
                Return Me.m_ObjParent
            End If
        End Get
        Set(ByVal Value As INode)
            If Value.Type = NODE_ENGINE Then
                m_ObjEngine = Value
            End If
            If Value.Type = NODE_ENVIRONMENT Then
                m_Environment = Value
            End If
            m_ObjParent = Value
        End Set
    End Property

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsVariable
        Dim cmd As Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            Me.LoadMe(cmd)


            obj.VariableName = Me.VariableName
            obj.VariableType = Me.VariableType
            obj.VarInitVal = Me.VarInitVal
            obj.VarSize = Me.VarSize
            obj.Parent = NewParent
            obj.SeqNo = Me.SeqNo

            If NewParent.Type <> NODE_ENGINE Then
                obj.Engine = Nothing
            End If

            Return obj

        Catch ex As Exception
            LogError(ex, "clsVariable Clone")
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
            If Me.Engine IsNot Nothing Then
                Key = Me.Engine.ObjSystem.Environment.Project.Text & KEY_SAP & Me.Engine.ObjSystem.Environment.Text & KEY_SAP & Me.Engine.ObjSystem.Text & KEY_SAP & Me.Engine.Text & KEY_SAP & Me.Text
            Else
                Key = Me.Project.Text & KEY_SAP & Me.Environment.Text & KEY_SAP & Me.Text
            End If
        End Get
    End Property

    '//8/15/05
    Public Property Text() As String Implements INode.Text
        Get
            Text = VariableName
        End Get
        Set(ByVal Value As String)
            VariableName = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            If VariableType = modDeclares.enumVariableType.VTYPE_LOCAL Or VariableType = modDeclares.enumVariableType.VTYPE_GLOBAL Then
                Return NODE_VARIABLE
            Else
                Return NODE_CONST
            End If
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
        Dim cmd As New System.Data.Odbc.OdbcCommand
        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            
            If Me.Engine Is Nothing Then
                sql = "Update " & Me.Project.tblVariables & " set VariableName=" & Me.GetQuotedText & _
                                " , EnvironmentName=" & Me.Environment.GetQuotedText & _
                                " , ProjectName=" & Me.Project.GetQuotedText & _
                                " , VariableDescription=" & "'" & FixStr(Me.VariableDescription) & "'" & _
                                " where VariableName=" & Me.GetQuotedText & _
                                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                                " AND ProjectName=" & Me.Project.GetQuotedText
                '" AND EngineName=*" & _
                '" AND SystemName=*" & _
            Else
                sql = "Update " & Me.Project.tblVariables & " set VariableName=" & Me.GetQuotedText & _
                                " , EngineName=" & Me.Engine.GetQuotedText & _
                                " , SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                                " , EnvironmentName=" & Me.Environment.GetQuotedText & _
                                " , ProjectName=" & Me.Project.GetQuotedText & _
                                " , VariableDescription=" & "'" & FixStr(Me.VariableDescription) & "'" & _
                                " where VariableName=" & Me.GetQuotedText & _
                                " AND EngineName=" & Me.Engine.GetQuotedText & _
                                " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                                " AND ProjectName=" & Me.Project.GetQuotedText
            End If

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)



            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Save = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsVariable Save", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// Modified by TK and KS 11/6/2006 
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

            
            '//When we add new record we need to find Unique Database Record ID.
            If Me.Engine Is Nothing Then
                sql = "INSERT INTO " & Me.Project.tblVariables & _
                                "(VariableName,EngineName,SystemName,EnvironmentName,ProjectName,VariableDescription) " & _
                                "Values(" & Me.GetQuotedText & "," & _
                                Quote(DBNULL) & "," & _
                                Quote(DBNULL) & "," & _
                                Me.Environment.GetQuotedText & "," & _
                                Me.Project.GetQuotedText & ",'" & _
                                FixStr(Me.VariableDescription) & "')"
            Else
                sql = "INSERT INTO " & Me.Project.tblVariables & _
                                "(VariableName,EngineName,SystemName,EnvironmentName,ProjectName,VariableDescription) " & _
                                "Values(" & Me.GetQuotedText & "," & _
                                Me.Engine.GetQuotedText & "," & _
                                Me.Engine.ObjSystem.GetQuotedText & "," & _
                                Me.Environment.GetQuotedText & "," & _
                                Me.Project.GetQuotedText & ",'" & _
                                FixStr(Me.VariableDescription) & "')"
            End If

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Add Var into Env or Engines's Var collection
            If Me.Engine Is Nothing Then
                AddToCollection(Me.Environment.Variables, Me, Me.GUID)
            Else
                AddToCollection(Me.Engine.Variables, Me, Me.GUID)
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsVariable AddNew", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

    '// new function added by KS and TK 11/6/2006
    '/// OverHauled by TKarasch   April May 07
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Dim sql As String = ""

        Try
            Me.Text = Me.Text.Trim

            
            '//When we add new record we need to find Unique Database Record ID.
            If Me.Engine Is Nothing Then
                sql = "INSERT INTO " & Me.Project.tblVariables & _
                                 "(VariableName,EngineName,SystemName,EnvironmentName,ProjectName,VariableDescription) " & _
                                 "Values(" & Me.GetQuotedText & "," & _
                                 Quote(DBNULL) & "," & _
                                 Quote(DBNULL) & "," & _
                                 Me.Environment.GetQuotedText & "," & _
                                 Me.Project.GetQuotedText & ",'" & _
                                 FixStr(Me.VariableDescription) & "')"
            Else
                sql = "INSERT INTO " & Me.Project.tblVariables & _
                                "(VariableName,EngineName,SystemName,EnvironmentName,ProjectName,VariableDescription) " & _
                                "Values(" & Me.GetQuotedText & "," & _
                                Me.Engine.GetQuotedText & "," & _
                                Me.Engine.ObjSystem.GetQuotedText & "," & _
                                Me.Environment.GetQuotedText & "," & _
                                Me.Project.GetQuotedText & ",'" & _
                                FixStr(Me.VariableDescription) & "')"
            End If

            Me.DeleteATTR(cmd)
            Me.InsertATTR(cmd)


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '//Add Var into Env or Engines's Var collection
            If Me.Engine Is Nothing Then
                AddToCollection(Me.Environment.Variables, Me, Me.GUID)
            Else
                AddToCollection(Me.Engine.Variables, Me, Me.GUID)
            End If

            AddNew = True
            Me.IsModified = False

        Catch ex As Exception
            LogError(ex, "clsVariable AddNew", sql)
            Return False
        End Try

    End Function

    '/// OverHauled by TKarasch   April May 07
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Dim sql As String = ""

        Try
            
            If Me.Engine Is Nothing Then
                sql = "Delete From " & Me.Project.tblVariables & _
                 " where VariableName=" & Me.GetQuotedText & _
                 " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                 " AND ProjectName=" & Me.Project.GetQuotedText
            Else
                sql = "Delete From " & Me.Project.tblVariables & _
                " where VariableName=" & Me.GetQuotedText & _
                " AND EngineName=" & Me.Engine.GetQuotedText & _
                " AND SystemName=" & Me.Engine.ObjSystem.GetQuotedText & _
                " AND EnvironmentName=" & Me.Environment.GetQuotedText & _
                " AND ProjectName=" & Me.Project.GetQuotedText
            End If

            Me.DeleteATTR(cmd)


            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            '/// the variables are removed from the Database, 
            '/// now we remove the child objects from the engines
            '/// IF we are removing it at the top (under environment)
            '/// and remove them from the tree
            If Me.Engine Is Nothing Then
                For Each sys As clsSystem In Me.Environment.Systems
                    For Each eng As clsEngine In sys.Engines
                        For Each var As clsVariable In eng.Variables
                            If var.VariableName = Me.VariableName Then
                                var.ObjTreeNode.Remove()
                                RemoveFromCollection(eng.Variables, var.GUID)
                            End If
                        Next
                    Next
                Next
            End If


            '/// Now finally we remove the objects from the Object Model
            If RemoveFromParentCollection = True Then
                '//Remove from parent collection
                If Me.Engine Is Nothing Then
                    RemoveFromCollection(Me.Environment.Variables, Me.GUID)
                Else
                    RemoveFromCollection(Me.Engine.Variables, Me.GUID)
                    RemoveFromCollection(Me.Environment.Variables, Me.GUID)
                End If
            End If

            Delete = True

        Catch ex As Exception
            LogError(ex, "clsVariable Delete", sql)
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

            If Me.Engine Is Nothing Then
                sql = "Select * from " & Me.Project.tblVariables & _
                " where PROJECTNAME=" & Me.Project.GetQuotedText & _
                " AND ENVIRONMENTNAME=" & Me.Environment.GetQuotedText
            Else
                sql = "Select * from " & Me.Project.tblVariables & _
                " where PROJECTNAME='" & Me.Project.ProjectName & _
                "' AND ENVIRONMENTNAME='" & Me.Environment.EnvironmentName & _
                "' AND SYSTEMNAME='" & Me.Engine.ObjSystem.SystemName & _
                "' AND ENGINENAME='" & Me.Engine.EngineName & "'"
            End If


            cmd.CommandText = sql
            Log(sql)
            dr = cmd.ExecuteReader

            While dr.Read
                readName = dr("VARIABLENAME")
                If testName.Equals(readName, StringComparison.CurrentCultureIgnoreCase) = True Then
                    NameValid = False
                    MsgBox("The new Variable name you have chosen already exists," & (Chr(13)) & "in this Environment, Please enter a different name", MsgBoxStyle.Information, MsgTitle)
                    Exit While
                End If
            End While
            dr.Close()

            Return NameValid

        Catch ex As Exception
            LogError(ex, "clsVariable ValidateNewObject", sql)
            Return False
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


            If Me.Engine IsNot Nothing Then
                sql = "SELECT VARIABLEATTRB,VARIABLEATTRBVALUE FROM " & Me.Project.tblVariablesATTR & _
                                " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & _
                                "' AND ENVIRONMENTNAME='" & FixStr(Me.Environment.EnvironmentName) & _
                                "' AND SystemNAME='" & FixStr(Me.Engine.ObjSystem.SystemName) & _
                                "' AND ENGINENAME='" & FixStr(Me.Engine.EngineName) & _
                                "' AND VariableName='" & FixStr(Me.VariableName) & "'"
            Else
                sql = "SELECT VARIABLEATTRB,VARIABLEATTRBVALUE FROM " & Me.Project.tblVariablesATTR & _
                                " WHERE PROJECTNAME='" & FixStr(Me.Project.ProjectName) & _
                                "' AND ENVIRONMENTNAME='" & FixStr(Me.Environment.EnvironmentName) & _
                                "' AND SystemNAME='" & DBNULL & _
                                "' AND ENGINENAME='" & DBNULL & _
                                "' AND VariableName='" & FixStr(Me.VariableName) & "'"
            End If


            cmd.CommandText = sql
            Log(sql)
            da = New System.Data.Odbc.OdbcDataAdapter(sql, cmd.Connection)
            da.Fill(dt)
            da.Dispose()

            For i As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(i)

                Attrib = GetVal(dr("VARIABLEATTRB"))
                Select Case Attrib
                    Case "VARINITVAL"
                        Me.VarInitVal = GetVal(dr("VARIABLEATTRBVALUE"))
                    Case "VARSIZE"
                        Me.VarSize = GetVal(dr("VARIABLEATTRBVALUE"))
                    Case "DESCRIPTION"
                        Me.VariableDescription = GetStr(GetVal(dr("VARIABLEATTRBVALUE")))
                End Select
            Next

            Me.IsLoaded = True

            Return True

        Catch ex As Exception
            LogError(ex, "clsVariable LoadMe")
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

#End Region

#Region "Properties"

    Public Property VariableName() As String
        Get
            Return m_VariableName
        End Get
        Set(ByVal Value As String)
            m_VariableName = Value
        End Set
    End Property

    Public Property CorrectedVariableName() As String
        Get
            Return m_CorrectedVariableName
        End Get
        Set(ByVal Value As String)
            m_CorrectedVariableName = Value
        End Set
    End Property

    Public Property VariableType() As enumVariableType
        Get
            Return m_VariableType
        End Get
        Set(ByVal Value As enumVariableType)
            m_VariableType = Value
        End Set
    End Property

    Public Property VariableDescription() As String
        Get
            Return m_VariableDescription
        End Get
        Set(ByVal Value As String)
            m_VariableDescription = Value
        End Set
    End Property

    Public Property VarInitVal() As String
        Get
            Return m_VarInitVal
        End Get
        Set(ByVal Value As String)
            m_VarInitVal = Value
        End Set
    End Property

    Public Property VarSize() As Integer
        Get
            Return m_VarSize
        End Get
        Set(ByVal Value As Integer)
            m_VarSize = Value
        End Set
    End Property

    Public Property Engine() As clsEngine
        Get
            Return m_ObjEngine
        End Get
        Set(ByVal Value As clsEngine)
            m_ObjEngine = Value
        End Set
    End Property

    Public Property Environment() As clsEnvironment
        Get
            If m_Environment Is Nothing Then
                Return Me.Engine.ObjSystem.Environment
            Else
                Return m_Environment
            End If
        End Get
        Set(ByVal Value As clsEnvironment)
            m_Environment = Value
        End Set
    End Property

#End Region

#Region "Methods"

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

            For i As Integer = 0 To 2
                Select Case i
                    Case 0
                        Attrib = "VARINITVAL"
                        Value = Me.VarInitVal
                    Case 1
                        Attrib = "VARSIZE"
                        Value = Me.VarSize
                    Case 2
                        Attrib = "DESCRIPTION"
                        Value = Me.VariableDescription
                End Select

                If Me.Engine Is Nothing Then
                    sql = "INSERT INTO " & Me.Project.tblVariablesATTR & _
                                    "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEATTRB,VARIABLEATTRBVALUE) " & _
                                    "Values('" & FixStr(Me.Project.ProjectName) & "','" & _
                                    FixStr(Me.Environment.EnvironmentName) & "','" & _
                                    DBNULL & "','" & _
                                    DBNULL & "','" & _
                                    FixStr(Me.VariableName) & "','" & _
                                    FixStr(Attrib) & "','" & _
                                    FixStr(Value) & "')"
                Else
                    sql = "INSERT INTO " & Me.Project.tblVariablesATTR & _
                                    "(PROJECTNAME,ENVIRONMENTNAME,SYSTEMNAME,ENGINENAME,VARIABLENAME,VARIABLEATTRB,VARIABLEATTRBVALUE) " & _
                                    "Values('" & FixStr(Me.Project.ProjectName) & "','" & _
                                    FixStr(Me.Environment.EnvironmentName) & "','" & _
                                    FixStr(Me.Engine.ObjSystem.SystemName) & "','" & _
                                    FixStr(Me.Engine.EngineName) & "','" & _
                                    FixStr(Me.VariableName) & "','" & _
                                    FixStr(Attrib) & "','" & _
                                    FixStr(Value) & "')"
                End If

                cmd.CommandText = sql
                Log(sql)
                cmd.ExecuteNonQuery()
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "clsVariable InsertATTR")
            Return False
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

            If Me.Engine Is Nothing Then
                sql = "DELETE FROM " & Me.Project.tblVariablesATTR & _
                            " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                            " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                            " AND VARIABLENAME=" & Quote(Me.VariableName)
                ' " AND SYSTEMNAME=*" & _
                ' " AND ENGINENAME=*" & _
            Else
                sql = "DELETE FROM " & Me.Project.tblVariablesATTR & _
                            " WHERE PROJECTNAME=" & Quote(Me.Project.ProjectName) & _
                            " AND ENVIRONMENTNAME=" & Quote(Me.Environment.EnvironmentName) & _
                            " AND SYSTEMNAME=" & Quote(Me.Engine.ObjSystem.SystemName) & _
                            " AND ENGINENAME=" & Quote(Me.Engine.EngineName) & _
                            " AND VARIABLENAME=" & Quote(Me.VariableName)
            End If
            

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            Return True

        Catch ex As Exception
            LogError(ex, "clsVariable DeleteATTR")
            Return False
        End Try

    End Function

   
#End Region

    Sub New()
        m_GUID = GetNewId()
    End Sub

    Sub New(ByVal Name As String, ByVal VType As enumVariableType)
        m_GUID = GetNewId()
        VariableName = Name
        VariableType = VType
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
