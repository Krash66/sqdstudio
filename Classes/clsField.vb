<Serializable()> _
Public Class clsField
    Implements INode

    Private m_Structure As clsStructure
    Dim m_OrgName As String = ""
    Dim m_arrAttrs(18) As String
    '//When we parse Att="...." Attribute from XML store values seperated by "~" in arrAttrs
    Private m_AttrCount As Integer
    Private m_FieldName As String
    Private m_FieldDesc As String
    Private m_FieldText As String
    '//sometimes field name may contain reserved words and so user can store 
    '//corrected name using this field
    Private m_CorrectedFieldName As String '//new :by npatel on 9/7/05 issue#48
    Private m_IsModified As Boolean
    Private m_ObjParent As INode
    Private m_ObjTreeNode As TreeNode
    Private m_GUID As String
    Private m_SeqNo As Integer = 0
    Private m_DescModified As Boolean
    Private m_IsRenamed As Boolean = False
    Private m_IsLoaded As Boolean = True

    Public Mappings As Collection '/// collection of Mappings using this field
    Public m_ParentName As String
    Public m_ParentStructureName As String

#Region "INode Implementation"

    Public Function GetQuotedText(Optional ByVal bFix As Boolean = True) As String Implements INode.GetQuotedText
        If bFix = True Then
            Return "'" & FixStr(Me.FieldName.Trim) & "'"
        Else
            Return "'" & Me.FieldName.Trim & "'"
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

    '//This requires special handling. 
    Public Property Parent() As INode Implements INode.Parent
        Get
            If m_Structure Is Nothing Then
                Return Me.m_ObjParent
            Else
                Return Me.m_Structure
            End If
        End Get
        Set(ByVal Value As INode)
            If Value.Type = NODE_STRUCT Then
                m_Structure = Value
            Else
                m_ObjParent = Value
            End If
        End Set
    End Property

    '/// OverHauled by TKarasch April 07 
    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object Implements INode.Clone

        Dim obj As New clsField
        Dim cmd As New Odbc.OdbcCommand

        Try
            If Incmd IsNot Nothing Then
                cmd = Incmd
            Else
                cmd = New Odbc.OdbcCommand
                cmd.Connection = cnn
            End If

            'obj.ObjTreeNode = Me.ObjTreeNode
            obj.IsModified = True
            obj.Struct = Me.Struct
            obj.Text = Me.Text
            obj.FieldName = Me.FieldName
            obj.FieldDesc = Me.FieldDesc
            obj.FieldDescModified = Me.FieldDescModified
            obj.OrgName = Me.OrgName
            obj.Parent = NewParent 'Me.Parent
            obj.ParentName = Me.ParentName
            obj.ParentStructureName = Me.ParentStructureName
            obj.SeqNo = Me.SeqNo
            obj.m_AttrCount = Me.m_AttrCount

            '// copy all field attributes
            Me.m_arrAttrs.CopyTo(obj.m_arrAttrs, 0)

            Return obj

        Catch ex As Exception
            LogError(ex, "clsField Clone")
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

    Public Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    '// Function added 11/3/06 by KS and TK
    Public Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean Implements INode.AddNew

        Return True

    End Function

    Public ReadOnly Property Key() As String Implements INode.Key
        Get
            'Return m_FieldId
            Return Me.Struct.Environment.Project.Text & KEY_SAP & Me.Struct.Environment.Text & KEY_SAP & Me.Struct.Text & KEY_SAP & Me.FieldName
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

        Return True

    End Function

    '// updated 4/20/07 by TK although I think this function is unused
    Public Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean Implements INode.Delete

        Try
            Dim sql As String

            sql = "Delete From StructFields where  FieldName=" & Me.GetQuotedText & " AND StructureName=" & Me.Struct.GetQuotedText & " AND EnvironmentName=" & Me.Struct.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete From strselFields where  FieldName=" & Me.GetQuotedText & " AND StructureName=" & Me.Struct.GetQuotedText & " AND EnvironmentName=" & Me.Struct.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            sql = "Delete From dsselFields where  FieldName=" & Me.GetQuotedText & " AND StructureName=" & Me.Struct.GetQuotedText & " AND EnvironmentName=" & Me.Struct.Environment.GetQuotedText & " AND ProjectName=" & Me.Project.GetQuotedText

            Log(sql)
            cmd.CommandText = sql
            cmd.ExecuteNonQuery()

            Delete = True

        Catch ex As Exception
            LogError(ex)
            Return False
        End Try

    End Function

    Public Property Text() As String Implements INode.Text
        Get
            If m_FieldText = "" Then
                Text = Me.FieldName
            Else
                Text = m_FieldText
            End If
        End Get
        Set(ByVal Value As String)
            m_FieldText = Value
        End Set
    End Property

    Public ReadOnly Property Type() As String Implements INode.Type
        Get
            Return NODE_STRUCT_FLD
        End Get
    End Property

    Public Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadItems

        Return True

    End Function

    Function LoadMe(Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean Implements INode.LoadMe

        Return True

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

        Return True

    End Function

#End Region

#Region "Properties"

    Public ReadOnly Property AttrCount() As Integer
        Get
            Return m_AttrCount
        End Get
    End Property

    Public Property CorrectedFieldName() As String
        Get
            Return m_CorrectedFieldName
        End Get
        Set(ByVal Value As String)
            m_CorrectedFieldName = Value
        End Set
    End Property

    Public Property OrgName() As String
        Get
            Return m_OrgName
        End Get
        Set(ByVal Value As String)
            m_OrgName = Value
        End Set
    End Property

    Public Property ParentName() As String
        Get
            Return m_ParentName
        End Get
        Set(ByVal Value As String)
            m_ParentName = Value
        End Set
    End Property

    Public Property ParentStructureName() As String
        Get
            If m_ParentStructureName = "" Then
                If Not Me.Parent Is Nothing Then
                    If Me.Parent.Type = NODE_STRUCT Then
                        Return Me.Parent.Text
                    Else
                        Return ""
                    End If
                Else
                    Return ""
                End If
            Else
                Return m_ParentStructureName
            End If
        End Get
        Set(ByVal Value As String)
            m_ParentStructureName = Value
        End Set
    End Property

    Public Property FieldName() As String
        Get
            Return m_FieldName
        End Get
        Set(ByVal Value As String)
            m_FieldName = Value
        End Set
    End Property

    Public Property Struct() As clsStructure
        Get
            Return m_Structure
        End Get
        Set(ByVal Value As clsStructure)
            m_Structure = Value
        End Set
    End Property

    Public Property FieldDesc() As String
        Get
            Return m_FieldDesc
        End Get
        Set(ByVal Value As String)
            m_FieldDesc = Value
        End Set
    End Property

    Public Property FieldDescModified() As Boolean
        Get
            Return m_DescModified
        End Get
        Set(ByVal Value As Boolean)
            m_DescModified = Value
        End Set
    End Property

#End Region

#Region "methods"

    Public Overridable Function GetFieldAttr(ByVal Attribute As enumFieldAttributes) As Object
        '//Attribute is actual position in array so always zerobased and 
        '// m_AttrCount is always 1 more than arrayupperbound

        GetFieldAttr = IIf(Len(m_arrAttrs(Attribute)) > 0, m_arrAttrs(Attribute), "")
        '// dont return uninitialized item it may cause error

    End Function

    '//set multiple attributes from xml attr string
    '//you can manually set also by creating string attr1~attr2~ ... attrn
    Public Overridable Function SetFieldAtt(ByVal strFldAttrs As String, Optional ByVal splitChar As String = "~") As Boolean

        Dim arr() As String

        Try
            arr = Split(strFldAttrs, splitChar)
            'cannull
            If ((arr(8) = "0") Or (arr(8) = "N") Or (arr(8) = "No") Or (arr(8) = "NO")) Then
                arr(8) = "No"
            Else
                arr(8) = "Yes"
            End If
            'iskey
            If ((arr(9) = "1") Or (arr(9) = "YES") Or (arr(9) = "Yes") Or (arr(9) = "Y")) Then
                arr(9) = "Yes"
            Else
                arr(9) = "No"
            End If
            'Fix for Oracle "Numeric" it is a decimal
            If arr(4) = "(null)" Then
                arr(4) = "NUMERIC"
            End If

            If arr.GetUpperBound(0) + 1 > 0 Then
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_NCHILDREN) = arr(0) '9
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_LEVEL) = arr(1)     '10
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_TIMES) = arr(2)     '11
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_OCCURS) = arr(3)    '12
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_DATATYPE) = arr(4)  '13
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_OFFSET) = arr(5)    '14
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_LENGTH) = arr(6)    '15
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_SCALE) = arr(7)     '16
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_CANNULL) = arr(8)   '17
                m_arrAttrs(modDeclares.enumFieldAttributes.ATTR_ISKEY) = arr(9)     '0

                m_AttrCount = arr.GetUpperBound(0) + 1
                SetFieldAtt = True
            Else
                SetFieldAtt = False
            End If

        Catch ex As Exception
            LogError(ex, "clsField SettFieldAtt")
            SetFieldAtt = False
        End Try

    End Function

    Public Function SetSingleFieldAttr(ByVal Attribute As modDeclares.enumFieldAttributes, ByVal value As Object) As Boolean

        Try
            If (Attribute = modDeclares.enumFieldAttributes.ATTR_ISKEY) Then
                If ((value = "1") Or (value = "YES") Or (value = "Yes") Or (value = "Y")) Then
                    value = "Yes"
                Else
                    value = "No"
                End If
            End If


            If (Attribute = modDeclares.enumFieldAttributes.ATTR_CANNULL) Then
                If ((value = "1") Or (value = "YES") Or (value = "Yes") Or (value = "Y")) Then
                    value = "Yes"
                Else
                    value = "No"
                End If
            End If

            m_arrAttrs(Attribute) = value

            Return True

        Catch ex As Exception
            LogError(ex, "clsField SetSingleFieldAttr")
            Return False
        End Try

    End Function

    '/// added by TKarasch  April 07 to add Field Descriptions
    Public Function UpdateFieldDesc() As Boolean

        Dim sql As String = ""
        'Dim cnn As System.Data.Odbc.OdbcConnection = Nothing
        Dim cmd As New Odbc.OdbcCommand

        Try
            'cnn = New Odbc.OdbcConnection(Project.MetaConnectionString)
            'cnn.Open()
            cmd.Connection = cnn

            'If Me.Project.ProjectMetaVersion = enumMetaVersion.V2 Then
            '    sql = "Update " & Me.Project.tblStructFields & " " & _
            '                   "set FIELDDESC='" & Me.FieldDesc & "' " & _
            '                   "where FieldName= " & Me.GetQuotedText & _
            '                   " AND StructureName=" & Me.Struct.GetQuotedText & _
            '                   " AND EnvironmentName=" & Me.Struct.Environment.GetQuotedText & _
            '                   " AND ProjectName=" & Me.Project.GetQuotedText
            'Else
            sql = "Update " & Me.Project.tblDescriptionFields & " " & _
                           "set DESCFIELDDESCRIPTION='" & Me.FieldDesc & "' " & _
                           "where FieldName= " & Me.GetQuotedText & _
                           " AND DescriptionName=" & Me.Struct.GetQuotedText & _
                           " AND EnvironmentName=" & Me.Struct.Environment.GetQuotedText & _
                           " AND ProjectName=" & Me.Project.GetQuotedText
            'End If

            cmd.CommandText = sql
            Log(sql)
            cmd.ExecuteNonQuery()

            UpdateFieldDesc = True
            Me.FieldDescModified = False

        Catch ex As Exception
            LogError(ex, "clsField UpdateFieldDesc", sql)
            Return False
        Finally
            'cnn.Close()
        End Try

    End Function

#End Region

    Public Sub New()
        m_GUID = GetNewId()
    End Sub

End Class
