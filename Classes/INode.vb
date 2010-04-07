'//All classes to be stored as objects in the treeview of treenodes will implement this
'//interface so it is easy to access them. The tree nodes themselves are objects that 
'//are indexed and can be manipulated. The classes INode implements are objects stored 
'//in the objects of the tree nodes. These objects have type information so that as 
'//they are stored or read from the MetaDataDB they can be put in to the folder nodes of the 
'//tree (Which are INodes with no real properties or methods i.e. dummy data, placeholders)
'//written out when saved, or added when read from the MetaDataDB. This means that the 
'//user sees folders for ease of use, but the MetaDataDB contains no info about folders 
'//of the tree.
'///// In summary, INode is like a bluprint for an object, and each class that uses this
'///// blueprint must define all the object properties the blueprint requires. That way, all
'///// objects are built from the same blueprint, and can interact with one another, even
'///// though they were built by different classes. Also, all objects built from INode
'///// can be treated the same way by "treeview" when they are moved around in the tree view.

''' <summary>INode is like a bluprint for an SQData object, and each class that uses this blueprint must define all the object properties the blueprint requires. All objects built from INode can be treated the same way by "treeview" when they are moved around in the tree view.
''' </summary>
Public Interface INode

    ReadOnly Property GUID() As String

    ReadOnly Property Key() As String

    ReadOnly Property Type() As String

    ReadOnly Property Project() As clsProject

    Property Text() As String '//Read/Write : 8/15/05

    Property IsModified() As Boolean

    Property Parent() As INode

    Property ObjTreeNode() As TreeNode

    Property SeqNo() As Integer '//Added on 5/21/05 : This indicates Index for TreeNode representing this object this will help to change position of object in tree node

    Function Save(Optional ByVal Cascade As Boolean = False) As Boolean

    Overloads Function AddNew(Optional ByVal Cascade As Boolean = False) As Boolean

    Overloads Function AddNew(ByRef cmd As Odbc.OdbcCommand, Optional ByVal Cascade As Boolean = False) As Boolean

    Function Delete(ByRef cmd As Odbc.OdbcCommand, ByRef cnn As Odbc.OdbcConnection, Optional ByVal Cascade As Boolean = True, Optional ByVal RemoveFromParentCollection As Boolean = True) As Boolean

    Function Clone(ByVal NewParent As INode, Optional ByVal Cascade As Boolean = True, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Object

    Function IsFolderNode() As Boolean

    Function GetQuotedText(Optional ByVal bFix As Boolean = True) As String

    '// Added By TKarasch to check MetaData before adding new Inode Objects
    Function ValidateNewObject(Optional ByVal NewName As String = "", Optional ByVal InReg As Boolean = False) As Boolean

    '/// Added By TKarasch to load Object itself, with or without 


    '/// Added by TKarasch to load all children of Inode Objects from MetaData
    Function LoadItems(Optional ByVal Reload As Boolean = False, Optional ByVal TreeLode As Boolean = False, Optional ByRef Incmd As Odbc.OdbcCommand = Nothing) As Boolean

    '/// Added By TKarasch as global flag to determine if Object was renamed
    Property IsRenamed() As Boolean

End Interface
