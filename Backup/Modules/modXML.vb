Imports System.Data
Imports System.Xml
Imports System.Xml.XPath
Public Module modXML

    Function EncodeXMLFile(ByVal file As String) As Boolean

        Try
            Dim s As String

            s = LoadTextFile(file)
            s = s.Replace("&", "&amp;")

            EncodeXMLFile = SaveTextFile(file, s)

        Catch ex As Exception
            LogError(ex, "modXML EncodeXMLFile")
            EncodeXMLFile = False
        End Try

    End Function

    'Function GetDataTableFromXML(ByVal XMLFilePath As String, ByVal PCap As String) As DataTable

    '    Dim xDoc As New XPathDocument(XMLFilePath)
    '    Dim xNav As XPathNavigator = xDoc.CreateNavigator()
    '    Dim xIterator As XPathNodeIterator = xNav.Select("/SideBar/Panel")
    '    Dim dt As New DataTable
    '    Dim dr As DataRow

    '    Try
    '        dt.Columns.Add("Attribute")
    '        dt.Columns.Add("Value")

    '        While xIterator.MoveNext()
    '            dr = dt.NewRow()
    '            dr.Item(0) = xIterator.Current.GetAttribute("Text", "")
    '            dr.Item(1) = xIterator.Current.GetAttribute("URL", "")
    '            dt.Rows.Add(dr)
    '        End While

    '        GetDataTableFromXML = dt

    '    Catch ex As Exception
    '        LogError(ex, "modXML GetDataTableFromXML")
    '        Return Nothing
    '    End Try

    'End Function

    '/////************** UnUsed AsOf 1/2007 **********************
    'Public Class AsyncLoading
    '    Public tv As TreeView
    '    Public arrList As ArrayList
    '    Public FileName As String
    '    Public SaveAttrs As String
    '    Public OnlyAddCheckedItem As Boolean = True
    '    '//exact same signature of tree.nodes.add(node as treenode)
    '    Private Delegate Sub AddDelegate(ByVal nd As TreeNode)

    '    '////////////////////////////////////////////////////////////////
    '    Sub New(ByRef m_tv As TreeView, ByRef m_arrList As ArrayList)
    '        tv = m_tv
    '        arrList = m_arrList
    '    End Sub

    '    Public Sub LoadTreeviewFromFldArrlist_Async()
    '        LoadTreeviewFromFldArrlist(tv.Nodes, arrList)
    '    End Sub
    '    '////////////////////////////////////////////////////////////////
    '    Sub New(ByRef m_tv As TreeView, ByRef m_FileName As String, Optional ByVal m_SaveAttrs As Boolean = False)
    '        tv = m_tv
    '        FileName = m_FileName
    '        SaveAttrs = m_SaveAttrs
    '    End Sub

    '    Public Sub LoadTreeViewFromXmlFile_Async()
    '        LoadTreeViewFromXmlFile(tv, FileName, tv, SaveAttrs)
    '    End Sub
    '    '////////////////////////////////////////////////////////////////
    '    Sub New(ByVal m_FileName As String, ByRef m_arrList As ArrayList)
    '        FileName = m_FileName
    '        arrList = m_arrList
    '    End Sub

    '    Public Sub LoadFldArrFromXmlFile_Async()
    '        LoadFldArrFromXmlFile(FileName, arrList)
    '    End Sub
    '    '////////////////////////////////////////////////////////////////
    '    Sub New(ByVal m_arrList As ArrayList, ByRef m_tv As TreeView, Optional ByVal m_OnlyAddCheckedItem As Boolean = True)
    '        arrList = m_arrList
    '        tv = m_tv
    '        OnlyAddCheckedItem = m_OnlyAddCheckedItem
    '    End Sub

    '    Public Sub LoadFldArrFromTreeview_Async()
    '        LoadFldArrFromTreeview(arrList, tv, OnlyAddCheckedItem)
    '    End Sub

    'End Class

    '/////////////////////////////////////////////////////////////////////////////////////////
    ' Load a TreeView control from an XML file.
    '//If SaveAttrs=True then routine will store clsFields object in Treenode.Tag

#Region "Load Trees from XML"

    Public Function LoadTreeViewFromXmlFile(ByVal Sender As Object, ByVal file_name As String, ByVal trv As TreeView, Optional ByVal SaveAttrs As Boolean = False, Optional ByVal XMLType As enumXMLType = modDeclares.enumXMLType.XMLFOR_FIELD) As Boolean

        Dim xml_doc As New XmlDocument
        ' Load the XML document.
        Try
            EncodeXMLFile(file_name) '//encode some special characters like & to &amp;

            xml_doc.Load(file_name)

            trv.BeginUpdate()

            ' Add the root node's children to the TreeView.
            trv.Nodes.Clear()

            AddTreeViewChildNodes(Sender, trv.Nodes, xml_doc.DocumentElement, SaveAttrs, XMLType)

            If trv.GetNodeCount(True) > 0 Then trv.Nodes.Item(0).EnsureVisible()

            Return True

        Catch ex As Exception
            LogError(ex, "modXML LoadTreeViewFromXmlFile")
            Return False
        Finally
            trv.EndUpdate()
        End Try

    End Function

    ' Add the children of this XML node
    ' to this child nodes collection.
    Sub AddTreeViewChildNodes(ByVal Sender As Object, ByVal parent_tvnodes As TreeNodeCollection, ByVal xml_node As XmlNode, Optional ByVal SaveAttrs As Boolean = False, Optional ByVal XMLType As enumXMLType = modDeclares.enumXMLType.XMLFOR_FIELD)

        Try
            For Each child_xmlnode As XmlNode In xml_node.ChildNodes
                ' Make the new TreeView node.
                Dim new_tvnode As TreeNode

                '//Now store fieldattributes in Tag if requested
                If SaveAttrs = True Then
                    '//If saveattrs true means we want to save object in treenode.tag
                    Dim obj As INode = GetObjectFromXMLAttr(child_xmlnode, XMLType)
                    '//function visible false on first load
                    If obj.Type = NODE_FUN Then
                        new_tvnode = AddNodeToCol(parent_tvnodes, obj.Type, obj, False)
                    ElseIf obj.Type = NODE_STRUCT_FLD Then
                        obj.Text = CType(obj, clsField).FieldName
                        CType(obj, clsField).Parent = obj
                        CType(obj, clsField).ParentStructureName = obj.Text
                        new_tvnode = AddNodeToCol(parent_tvnodes, obj.Type, obj, True)
                    Else
                        new_tvnode = AddNodeToCol(parent_tvnodes, obj.Type, obj, True)
                    End If
                    new_tvnode.Tag = obj
                Else
                    new_tvnode = parent_tvnodes.Add(child_xmlnode.Attributes("ID").Value)
                End If

                ' Recursively make this node's descendants.
                AddTreeViewChildNodes(Sender, new_tvnode.Nodes, child_xmlnode, SaveAttrs, XMLType)

                ' If this is a leaf node, make sure it's visible.
                If new_tvnode.Nodes.Count = 0 Then new_tvnode.EnsureVisible()
            Next 'child_xmlnode

        Catch ex As Exception
            LogError(ex, "modXML AddTreeViewChildNodes")
        End Try

    End Sub

    '//This function will return a object from XML node which can be stored in treenode.tag
    Private Function GetObjectFromXMLAttr(ByVal xml_node As XmlNode, Optional ByVal XMLType As enumXMLType = modDeclares.enumXMLType.XMLFOR_FIELD) As Object

        Try
            Select Case XMLType
                Case modDeclares.enumXMLType.XMLFOR_FIELD
                    Dim objFld As New clsField
                    With objFld
                        .SetFieldAtt(xml_node.Attributes("Att").Value)
                        .FieldName = xml_node.Attributes("ID").Value
                        .OrgName = IIf(xml_node.Attributes("OrgName") Is Nothing, _
                                       "", _
                                       xml_node.Attributes("OrgName").Value)
                        .ParentName = xml_node.Attributes("Parent").Value
                    End With
                    Return objFld
                    Exit Function

                Case modDeclares.enumXMLType.XMLFOR_SQFUNCTIONS
                    If xml_node.Attributes("NODETYPE").Value <> "Heading" Then
                        Dim objFun As New clsSQFunction
                        With objFun
                            .SQFunctionName = xml_node.Name
                            .SQFunctionDescription = xml_node.Attributes("DESC").Value
                            .SQFunctionSyntax = xml_node.Attributes("SYNTAX").Value
                            .ParaCount = xml_node.Attributes("NPARM").Value
                            .IsTemplate = IIf(xml_node.Attributes("NODETYPE").Value = "Template", True, False)
                        End With
                        Return objFun
                    Else
                        Dim objFolder As New clsFolderNode(xml_node.Name, _
                                                           IIf(xml_node.Name = "Templates", _
                                                               NODE_FO_TEMPLATE, NODE_FO_FUNCTION))
                        Return objFolder
                    End If
                    Exit Function

                Case modDeclares.enumXMLType.XMLFOR_SEGMENTS
                    Return Nothing
                    Exit Function
                    '//TODO
                    '//We really dont use segment attributes so just node text is fine to show. 
                    '//At this moment we dont use other info

                Case Else
                    Return Nothing
                    Exit Function

            End Select

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modXML GetObjectFromXMLAttr")
            Return Nothing
        End Try

    End Function

    '/////////////////////////////////////////////////////////////////////////////////////////
    '// If XML is IMS Segment then Return will be Att of first root node. Att is Database name
    '/////////////////////////////////////////////////////////////////////////////////////////
    'Load a fields and their attributes from an XML file to an ArrayList.
    Function LoadFldArrFromXmlFile(ByVal file_name As String, ByRef arrList As ArrayList) As String

        ' Load the XML document.
        Dim xml_doc As New XmlDocument
        Dim IMSDBName As String

        Try
            EncodeXMLFile(file_name) '//encode some special characters like & to &amp;
            xml_doc.Load(file_name)

            If Not xml_doc.DocumentElement.Attributes("Att") Is Nothing Then
                '//IMSDbName is stored in root node in "Att" attibute
                IMSDBName = xml_doc.DocumentElement.Attributes("Att").Value
                IMSDBName = IMSDBName.Substring(1, IMSDBName.IndexOf("~", 1) - 1)
                LoadFldArrFromXmlFile = IMSDBName
            Else
                LoadFldArrFromXmlFile = ""
            End If

            arrList.Clear()
            AddXMLItemToArr(arrList, xml_doc.DocumentElement)

        Catch ex As Exception
            LogError(ex, "modXML LoadFldArrFromXmlFile")
            Return ""
        End Try

    End Function

    ' Add the children of this XML node
    ' to this child nodes collection.
    Sub AddXMLItemToArr(ByRef arrList As ArrayList, ByVal xml_node As XmlNode)

        Try
            For Each child_xmlnode As XmlNode In xml_node.ChildNodes
                '//Now store fieldattributes in Tag if requested
                Dim objAttrs As New clsField

                objAttrs.SetFieldAtt(child_xmlnode.Attributes("Att").Value)
                objAttrs.FieldName = child_xmlnode.Attributes("ID").Value
                objAttrs.OrgName = IIf(child_xmlnode.Attributes("OrgName") Is Nothing, "", child_xmlnode.Attributes("OrgName").Value)

                arrList.Add(objAttrs)

                ' Recursively make this node's descendants.
                AddXMLItemToArr(arrList, child_xmlnode)

            Next 'child_xmlnode

        Catch ex As Exception
            LogError(ex, " modXML AddXMLItemToArr")
        End Try

    End Sub

    Public Function LoadFunctionFromXML(ByVal tv As TreeView) As Boolean
        '//only load once then cache in memory for faster loading
        Try
            If tv.GetNodeCount(True) <= 0 Then
                '//Load from XML
                tv.ImageList = modDeclares.imgListSmall
                LoadTreeViewFromXmlFile(tv, GetAppPath() & "sqdsyfun.xmq", tv, True, modDeclares.enumXMLType.XMLFOR_SQFUNCTIONS)

                For Each nd As TreeNode In tv.Nodes
                    nd.Collapse()
                Next
                '//add very recently used function node
                Dim objFolder As New clsFolderNode("Recently Used", NODE_FO_FUNCTION_RECENT)
                AddTreeNode(tv, NODE_FO_FUNCTION_RECENT, objFolder)
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modXML LoadFunctionFromXML")
            Return False
        End Try

    End Function

#End Region

    '/////////////////////////////////////////////////////////////////////////////////////////
    '//strFileToParse : Input file path which is to be parsed and exported to XML
    '//strCommandCode : 3 char command code which will tell SQDUIIMP.exe which type of input file
    '//par1 : only needed when we use DBS command which export segment info in XML
    '//par2 : not used
    '//
    '//List of Commands and ouputfile prefix for SQDUIIMP.exe
    '//
    '"DBO" for DBD (xml file prefix: SD)
    '"DBS" for COBOL with DBD segment combination (xml file prefix: SS)
    '"COB" for COBOL without DBD segment combination (xml file prefix: SC)
    '"DTD" for XML DTD (xml file prefix: ST)
    '"DDL" for SQL DDL (xml file prefix: SL)
    '"INC" for C Header (xml file prefix: SI)
    'Function GetSQDumpXML(ByVal strFileToParse As String _
    '                , Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN _
    '                , Optional ByVal par1 As String = "" _
    '                , Optional ByVal par2 As String = "") As String

    '    Dim strCommandCode As String = ""
    '    Dim strFilePrefix As String = ""
    '    Dim strTempDir As String = GetAppPath() & "Temp\"
    '    Dim args As String

    '    Try
    '        Select Case StructType
    '            Case modDeclares.enumStructure.STRUCT_COBOL
    '                strFilePrefix = "SC"
    '                strCommandCode = "COB"
    '            Case modDeclares.enumStructure.STRUCT_COBOL_IMS
    '                strFilePrefix = "SS"
    '                strCommandCode = "DBS"
    '            Case modDeclares.enumStructure.STRUCT_IMS
    '                strFilePrefix = "SD"
    '                strCommandCode = "DBO"
    '            Case modDeclares.enumStructure.STRUCT_XMLDTD
    '                strFilePrefix = "ST"
    '                strCommandCode = "DTD"
    '            Case modDeclares.enumStructure.STRUCT_REL_DDL
    '                strFilePrefix = "SL"
    '                strCommandCode = "DDL"
    '            Case modDeclares.enumStructure.STRUCT_C
    '                strFilePrefix = "SH"
    '                strCommandCode = "INC"
    '            Case modDeclares.enumStructure.STRUCT_REL_DML_FILE
    '                strFilePrefix = "SR"
    '                strCommandCode = "DML"
    '        End Select

    '        If StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
    '            GetSQDumpXML = strTempDir & "\" & strFilePrefix & par1 & ".xml"
    '        Else
    '            GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"
    '        End If

    '        '//delete previous temp file
    '        If IO.File.Exists(GetSQDumpXML) = True Then
    '            IO.File.Delete(GetSQDumpXML)
    '            GetSQDumpXML = ""
    '        End If

    '        '//delete previous log file
    '        If IO.File.Exists(IO.Path.Combine(GetAppTemp(), "sqduiimp.log")) Then
    '            IO.File.Delete(IO.Path.Combine(GetAppTemp(), "sqduiimp.log"))
    '        End If


    '        args = Quote(strFileToParse, """") & " " & strCommandCode & " " & Quote(strTempDir, """")

    '        If strCommandCode = "DBS" Then
    '            'par1=> segment name and par2=cob file name
    '            args = args & " " & par1 & " " & par2
    '        Else
    '            args = args
    '        End If

    '        Debug.Write("sqduiimp.exe" & args)
    '        Log("sqduiimp.exe " & args)

    '        Try
    '            '//run out little exe with command line args so it produces meta data in XML format
    '            Dim si As New System.Diagnostics.ProcessStartInfo
    '            Dim myProcess As System.Diagnostics.Process

    '            si.CreateNoWindow = True
    '            si.WindowStyle = ProcessWindowStyle.Hidden

    '            '#If CONFIG = "ETI" Then
    '            '                si.FileName = GetAppPath() & "ETIuiimp.exe"
    '            '#Else
    '            si.FileName = GetAppPath() & "sqduiimp.exe"
    '            '#End If

    '            si.Arguments = args
    '            Log(si.FileName & " " & si.Arguments)
    '            'Log(args)
    '            '//Create a new process to parse input file and dump to XML
    '            myProcess = System.Diagnostics.Process.Start(si)

    '            '//wait until task is done
    '            myProcess.WaitForExit()

    '            MsgBox("exit code >>> " & myProcess.ExitCode, MsgBoxStyle.Information)

    '            If myProcess.ExitCode <> 0 Then
    '                If MsgBox("Error occurred while parsing the file [" & strFileToParse & "]." & vbCrLf & "Do you want to see the log?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNoCancel, MsgTitle) = MsgBoxResult.Yes Then
    '                    If IO.File.Exists(IO.Path.Combine(GetAppPath(), "sqduiimp.log")) Then
    '                        Shell("notepad " & IO.Path.Combine(GetAppPath(), "sqduiimp.log"), AppWinStyle.NormalFocus)
    '                    End If
    '                End If
    '                Return ""
    '            Else
    '                '//Now return output file path : <temp folder>\<fileprefix><input filename no extension>.XML
    '                If StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
    '                    GetSQDumpXML = strTempDir & "\" & strFilePrefix & par1 & ".xml"
    '                Else
    '                    GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"
    '                End If
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "modXML GetSQDumpXML .. interior process")
    '            Return ""
    '        End Try

    '    Catch ex As Exception
    '        LogError(ex, "modXML GetSQDumpXML")
    '        Return ""
    '    End Try

    'End Function




    '/////////////////////////////////////////////////////////////////////////////////////////
    '//strFileToParse : Input file path which is to be parsed and exported to XML
    '//strCommandCode : 3 char command code which will tell SQDUIIMP.exe which type of input file
    '//par1 : only needed when we use DBS command which export segment info in XML
    '//par2 : not used
    '//
    '//List of Commands and ouputfile prefix for SQDUIIMP.exe
    '//
    '"DBO" for DBD (xml file prefix: SD)
    '"DBS" for COBOL with DBD segment combination (xml file prefix: SS)
    '"COB" for COBOL without DBD segment combination (xml file prefix: SC)
    '"DTD" for XML DTD (xml file prefix: ST)
    '"DDL" for SQL DDL (xml file prefix: SL)
    '"INC" for C Header (xml file prefix: SI)

#Region "Get XML structure Imports"

    Function GetSQDumpXML(ByVal strFileToParse As String _
                   , Optional ByVal StructType As enumStructure = modDeclares.enumStructure.STRUCT_UNKNOWN _
                   , Optional ByVal par1 As String = "" _
                   , Optional ByVal par2 As String = "") As String

        Dim strCommandCode As String = ""
        Dim strFilePrefix As String = ""
        Dim strTempDir As String = GetAppTemp()
        Dim args As String
        Dim fsERR As System.IO.FileStream
        Dim objWriteERR As System.IO.StreamWriter
        Dim PathErr As String

        Try
            Select Case StructType
                Case modDeclares.enumStructure.STRUCT_COBOL
                    strFilePrefix = "SC"
                    strCommandCode = "COB"
                Case modDeclares.enumStructure.STRUCT_COBOL_IMS
                    strFilePrefix = "SS"
                    strCommandCode = "DBS"
                Case modDeclares.enumStructure.STRUCT_IMS
                    strFilePrefix = "SD"
                    strCommandCode = "DBO"
                Case modDeclares.enumStructure.STRUCT_XMLDTD
                    strFilePrefix = "ST"
                    strCommandCode = "DTD"
                Case modDeclares.enumStructure.STRUCT_REL_DDL
                    strFilePrefix = "SL"
                    strCommandCode = "DDL"
                Case modDeclares.enumStructure.STRUCT_C
                    strFilePrefix = "SH"
                    strCommandCode = "INC"
                Case modDeclares.enumStructure.STRUCT_REL_DML_FILE
                    strFilePrefix = "SR"
                    strCommandCode = "DML"
            End Select

            If StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                GetSQDumpXML = strTempDir & "\" & strFilePrefix & par1 & ".xml"   '& "\" 
            Else
                GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"   '& "\" & "\" 
            End If

            '//delete previous temp file
            If IO.File.Exists(GetSQDumpXML) = True Then
                IO.File.Delete(GetSQDumpXML)
                GetSQDumpXML = ""
            End If

            '//delete previous log file
            'If IO.File.Exists(IO.Path.Combine(GetAppTemp(), "sqduiimp.log")) Then
            '    IO.File.Delete(IO.Path.Combine(GetAppTemp(), "sqduiimp.log"))
            'End If


            args = Quote(strFileToParse, """") & " " & strCommandCode & " " & Quote(strTempDir, """")

            If strCommandCode = "DBS" Then
                'par1=> segment name and par2=cob file name
                args = args & " " & par1 & " " & par2
            Else
                args = args
            End If

            'Debug.Write("sqduiimp.exe" & args)
            'Log("sqduiimp.exe " & args)

            Try
                '//delete previous log file
                If System.IO.File.Exists(IO.Path.Combine(GetAppLog(), "sqduiimp.ERR")) Then
                    System.IO.File.Delete(IO.Path.Combine(GetAppLog(), "sqduiimp.ERR"))
                End If
                '// create new error file stream
                fsERR = System.IO.File.Create(IO.Path.Combine(GetAppLog(), "sqduiimp.ERR"))
                PathErr = fsERR.Name
                objWriteERR = New System.IO.StreamWriter(fsERR)

                '//run out exe with command line args so it produces meta data in XML format
                Dim si As New ProcessStartInfo()

                si.FileName = GetAppPath() & "sqduiimp.exe" '
                si.WorkingDirectory = GetAppTemp()
                si.Arguments = args
                si.UseShellExecute = False
                si.CreateNoWindow = True

                '// Redirect standard error. Let standard output go where it is supposed to go.
                si.RedirectStandardOutput = False
                si.RedirectStandardError = True

                '//Create a new process to parse input file and dump to XML
                Using myProcess As New System.Diagnostics.Process()
                    myProcess.StartInfo = si

                    Log("Importer Run Started : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                    Log(si.FileName & " " & args)

                    '// Create a new process to asynchronously run sqduiimp.exe
                    myProcess.Start()

                    Dim OutStr As String = ""
                    Dim ErrStr As String = ""

                    '/// split output into multiple threads to capture each stream to a string
                    '/// OutStr stays as "" because StdOut is NOT redirected
                    OutputToEnd(myProcess, OutStr, ErrStr)

                    '//wait until task is done
                    myProcess.WaitForExit()

                    Log("Importer Run Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                    Log("Importer Returned Code : " & myProcess.ExitCode)
                    'Just for debugging
                    'MsgBox("exit code >>> " & myProcess.ExitCode, MsgBoxStyle.Information)

                    objWriteERR.Write(ErrStr)
                    objWriteERR.Close()
                    fsERR.Close()

                    If myProcess.ExitCode <> 0 Then
                        If MsgBox("Error occurred while parsing the file [" & strFileToParse & "]." & vbCrLf & _
                                  "Do you want to see the log?", _
                                  MsgBoxStyle.Critical Or MsgBoxStyle.YesNoCancel, _
                                  MsgTitle) = MsgBoxResult.Yes Then
                            If IO.File.Exists(PathErr) Then
                                Process.Start(PathErr)
                            End If
                        End If
                        Return ""
                    Else
                        '//Now return output file path : 
                        '<temp folder>\<fileprefix><input filename no extension>.XML
                        If StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                            GetSQDumpXML = strTempDir & "\" & strFilePrefix & par1 & ".xml"    '& "\"
                        Else
                            GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"  '& "\" & "\"
                        End If
                    End If

                    Log("Importer Report file saved at : " & PathErr)
                    Log("********* Importer Return Code = " & myProcess.ExitCode & " *********")
                    myProcess.Close()

                End Using

            Catch ex As Exception
                LogError(ex, "modXML GetSQDumpXML .. interior process")
                Return ""
            End Try

        Catch ex As Exception
            LogError(ex, "modXML GetSQDumpXML")
            Return ""
        End Try

    End Function

#End Region

#Region "XML Message Processing"

    '/// *** Created by TK   November 2011
    '/// *** Creates XML DTD files from XML messages
    '/// Main function:
    '/// -- Function ProcessXMLmessage(ByVal InfilePath As String, ByVal DirOut As String) As String
    '/// Arguments: 
    '/// -- InfilePath --> Path of the XML Message file to convert to DTD
    '/// -- DirOut     --> Path of Output Directory to Put new DTD file into
    '/// Return Value:
    '/// -- String Value of the complete filepath of the New DTD file 
    '/// -- or "" if no new file wanted or no new file was created, by user choice or by Error

    Private InFilePath As String                   '** Input XML Message File Path
    Private OutFilePath As String                  '** Output DTD XML Description File
    Private OutDir As String                       '** Output Directory for the new DTD file
    Private xml_Indoc As New XmlDocument           '** Windows XML doc that XML message is read into

    Private ArrAllElements As New ArrayList        '** Array of all elements in Document
    Private ArrParentNodes As New ArrayList        '** Array of all elements that are parents of other elements
    Private ArrPrintedChildren As New ArrayList    '** Array of child elements that have children
    Private ArrCDataNodes As New ArrayList         '** Array of child elements that have NO children
    Private InputDir As String = ""                '** Input Directory of the XML Message
    Private sbldr As System.Text.StringBuilder     '** String Builder Object that is built to create DTD message

    Private Enum enumXMLActionType
        Failed = 0
        E_PrintwCldrn = 1
        E_PrintAsCdata = 2
        E_Ignore = 3
    End Enum

    Public Function ProcessXMLmessage(ByVal InfilePath As String, ByVal DirOut As String) As String

        Try
            ArrAllElements.Clear()
            ArrParentNodes.Clear()
            ArrPrintedChildren.Clear()
            ArrCDataNodes.Clear()

            OutDir = DirOut
            sbldr = New System.Text.StringBuilder

            'Convert xml_doc to new string builder text
            'EncodeXMLFile(InFilePath) '//encode some special characters like & to &amp;
            xml_Indoc.Load(InfilePath)

            '*** Create a new filepath of the DTD named the same as the XML Message
            '*** in the new output directory
            OutFilePath = OutDir & "\" & GetFileNameWithoutExtenstionFromPath(InfilePath) & ".dtd"

            '*** See if the DTD file already exists
            '*** If so ask user how to proceed
TryAgain:   If System.IO.File.Exists(OutFilePath) = True Then
                Dim result As MsgBoxResult
                result = MsgBox("The file: " & _
                                GetFileNameFromPath(OutFilePath) & _
                                " Already exists." & Chr(13) & Chr(13) & _
                                " (Yes) Create a new DTD with a new Name." & Chr(13) & _
                                " (No) Overwrite Existing File." & Chr(13) & _
                                " (Cancel) Use Existing File.", _
                                MsgBoxStyle.YesNoCancel, _
                                "DTD File Already Exists. Create New, Overwrite or Use Existing.")
                '*** Process result of MsgBox ***
                Select Case result
                    Case MsgBoxResult.Yes
                        '*** Get New name from InputBox and insert it into the OutFilePath
                        OutFilePath = GetDirFromPath(OutFilePath) & _
                        InputBox("Enter New Name (no extension) for .dtd File", _
                                 "Rename DTD file", _
                                 GetFileNameWithoutExtenstionFromPath(OutFilePath)) & _
                                 ".dtd"
                        GoTo TryAgain
                    Case MsgBoxResult.Cancel
                        '*** Don't build a new DTD
                        '*** Use the existing file, Return the Path, and EXIT
                        ProcessXMLmessage = OutFilePath
                        Exit Try
                    Case Else
                        'Overwrite will happen automatically
                End Select
            End If

            'Start processing each node in the Message
            If xml_Indoc.HasChildNodes = True Then
                For Each nd As XmlNode In xml_Indoc.ChildNodes
                    '*** Process each Node, if it is an element
                    '*** if it's not an element, ignore it
                    If nd.NodeType = XmlNodeType.Element Then
                        If processNode(nd) = False Then
                            ProcessXMLmessage = ""
                            Exit Function
                        End If
                    End If
                Next 'next doc child
            End If

            If printCData(ArrCDataNodes) = True Then
                If SaveTextFile(OutFilePath, sbldr.ToString) = True Then
                    '***** Success ***** return the dtd file Path
                    ProcessXMLmessage = OutFilePath
                Else
                    ProcessXMLmessage = ""
                End If
            Else
                ProcessXMLmessage = ""
            End If

        Catch ex As Exception
            LogError(ex, "modXML ProcessXMLmessage")
            ProcessXMLmessage = ""
        End Try

    End Function

    Private Function processNode(ByVal Node As XmlNode) As Boolean

        Try
            Dim Action As enumXMLActionType = GetNodeAction(Node)

            Select Case Action
                Case enumXMLActionType.E_PrintwCldrn

                    '*** Print the Element with it's children
                    printNodeWithChildren(Node)
                    '*** Now process the Children
                    For Each cld As XmlNode In Node.ChildNodes
                        processNode(cld)
                    Next

                Case enumXMLActionType.E_PrintAsCdata
                    Dim QualName As String = GetCDataName(Node)
                    ArrCDataNodes.Add(QualName)

                Case enumXMLActionType.E_Ignore
                    '*** Do Nothing except goto next sibling or parent

                Case enumXMLActionType.Failed
                    If MsgBox("Translation failed at Element: " & Node.Name & Chr(13) & _
                           "Would you like to continue processing?", _
                           MsgBoxStyle.YesNo, _
                           "Translation Failed") = MsgBoxResult.No Then
                        '*** abort here
                        Return False
                        Exit Function
                    End If

            End Select

            Return True

        Catch ex As Exception
            LogError(ex, "modXML processNode")
            Return False
        End Try

    End Function

    Private Function GetNodeAction(ByVal nd As XmlNode) As enumXMLActionType

        Try
            Dim IsParentElement As Boolean = False

            If nd.HasChildNodes = True Then
                Dim NumEle As Integer = 0
                For Each node As XmlNode In nd.ChildNodes
                    If node.NodeType = XmlNodeType.Element Then
                        If node.HasChildNodes = True Then
                            If node.FirstChild.NodeType = XmlNodeType.Element Then
                                IsParentElement = True
                            End If
                        End If
                        NumEle += 1
                        If NumEle = 2 Then
                            GetNodeAction = enumXMLActionType.E_PrintwCldrn
                            Exit Function
                        End If
                    End If
                Next
                If IsParentElement = True Then
                    GetNodeAction = enumXMLActionType.E_PrintwCldrn
                Else
                    GetNodeAction = enumXMLActionType.E_PrintAsCdata
                End If
            Else
                If nd.NodeType = XmlNodeType.Element Then
                    GetNodeAction = enumXMLActionType.E_PrintAsCdata
                Else
                    GetNodeAction = enumXMLActionType.E_Ignore
                    Log("XML node:  " & nd.LocalName & "  was Ignored")
                End If
            End If


        Catch ex As Exception
            LogError(ex, "modXML GetNodeAction")
            GetNodeAction = enumXMLActionType.Failed
            sbldr.AppendLine("   --- Node action Failed ::" & nd.Name)
            sbldr.AppendLine("   --- Check XML Message for Syntax Errors")
        End Try

    End Function

    Private Function GetElementName(ByVal nd As XmlNode) As String

        Try
            Dim NewName As String = nd.LocalName
            Dim origName As String = NewName
            Dim count As Integer = 0

TryAgain:   If ArrParentNodes.Contains(NewName) = True Then
                count += 1
                NewName = origName & count.ToString
                GoTo TryAgain
            End If

            GetElementName = NewName

        Catch ex As Exception
            LogError(ex, "modXML GetNameToAdd")
            GetElementName = ""
        End Try

    End Function

    Private Function GetChildName(ByVal nd As XmlNode) As String

        Try
            Dim NewName As String = nd.LocalName
            Dim origName As String = NewName
            Dim count As Integer = 0

TryAgain:   If ArrAllElements.Contains(NewName) = True Then
                count += 1
                NewName = origName & count.ToString
                GoTo TryAgain
            End If

            GetChildName = NewName

        Catch ex As Exception
            LogError(ex, "modXML GetChildName")
            GetChildName = ""
        End Try

    End Function

    Private Function GetCDataName(ByVal nd As XmlNode) As String

        Try
            Dim NewName As String = nd.LocalName
            Dim origName As String = NewName
            Dim count As Integer = 0

TryAgain:   If ArrParentNodes.Contains(NewName) = True Then
                count += 1
                NewName = origName & count.ToString
                GoTo TryAgain
            End If
            If ArrPrintedChildren.Contains(NewName) = True Then
                ArrPrintedChildren.Remove(NewName)
                GetCDataName = NewName
            Else
                count += 1
                NewName = origName & count.ToString
                GoTo TryAgain
            End If

        Catch ex As Exception
            LogError(ex, "modXML GetCDataNameToAdd")
            GetCDataName = ""
        End Try

    End Function

    Private Function printNodeWithChildren(ByVal node As XmlNode) As Boolean

        Try
            Dim QualName As String = GetElementName(node)

            If QualName <> "" Then
                Dim ChildElementName As String = ""
                Dim Prefix As String = "     "
                Dim Line1 As String = String.Format("{0}{1}", "<!ELEMENT ", QualName)
                Dim Line2 As String = String.Format("{0}", "(")
                Dim LastLine As String = String.Format("{0}", ")>")

                '*** Add to the all elements list for comparison BEFORE printing
                '*** So child can't have parent's name
                ArrAllElements.Add(QualName)
                ArrParentNodes.Add(QualName)

                '*** Print element and it's children
                sbldr.AppendLine(Line1)
                sbldr.AppendLine(Line2)
                For Each cld As XmlNode In node.ChildNodes
                    If cld.NodeType = XmlNodeType.Element Then
                        ChildElementName = GetChildName(cld)
                        Dim LineChild As String = String.Format("{0}{1}", Prefix, ChildElementName)
                        sbldr.AppendLine(LineChild)
                        Prefix = "    ,"
                        '*** Add children to the child array
                        ArrAllElements.Add(ChildElementName)
                        ArrPrintedChildren.Add(ChildElementName)
                    End If
                Next
                sbldr.AppendLine(LastLine)
                sbldr.AppendLine()

                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            LogError(ex, "modXML printNode")
            Return False
        End Try

    End Function

    Private Function printCData(ByVal Arr As ArrayList) As Boolean

        Try
            Dim FORfld As String

            For Each name As String In Arr
                FORfld = String.Format("{0}{1,-40}{2}", "<!ELEMENT ", name, " (#PCDATA)>")
                sbldr.AppendLine(FORfld)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "modXML printCData")
            Return False
        End Try

    End Function

#End Region

#Region "Main Program XML"

    Function SaveStudioToXML(Optional ByVal Newpath As Boolean = False) As Boolean

        Try
            '*** Path to Studio XML file
            If Newpath = True Then
                AppDataPath = My.Computer.FileSystem.SpecialDirectories.MyDocuments()
                If Right(AppDataPath, 1) <> "\" Then
                    AppDataPath = AppDataPath & "\"
                End If
                AppDataPath = AppDataPath & "Design Studio\"
                If System.IO.Directory.Exists(AppDataPath) = False Then
                    System.IO.Directory.CreateDirectory(AppDataPath)
                End If
            End If

            Dim StudioXMLFullPath As String = System.Windows.Forms.Application.LocalUserAppDataPath() & _
            "\DesignStudio.settings.xml"

            '*** New XML writer for XML file
            Dim XMLwrite As New Xml.XmlTextWriter(StudioXMLFullPath, System.Text.Encoding.UTF8)

            '*** define doctype and formatting and Open XML file
            XMLwrite.Formatting = Formatting.Indented
            XMLwrite.WriteStartDocument()
            XMLwrite.WriteStartElement("DesignStudio", "DesignStudio", StudioXMLFullPath)

            '*** write Data
            XMLwrite.WriteElementString("Winstate", WinState.ToString())
            XMLwrite.WriteElementString("AppDataPath", AppDataPath)
            XMLwrite.WriteElementString("CCurl", CCurl)

            '*** write closing element and close file
            XMLwrite.WriteEndElement()
            XMLwrite.WriteEndDocument()
            XMLwrite.Close()

            Return True

        Catch ex As Exception
            LogError(ex, "modXML SaveStudioToXML")
            Return False
        End Try

    End Function

    '// New 11/2011 to get rid of Project saving to Registry
    Function RetrieveStudioFromXML() As Boolean

        Try
            Dim curNode As XmlNode
            '*** Path to Project XML file
            Dim StudioXMLFullPath As String = System.Windows.Forms.Application.LocalUserAppDataPath() & _
            "\DesignStudio.settings.xml"

            If System.IO.File.Exists(StudioXMLFullPath) = False Then
                'AppDataPath = System.Windows.Forms.Application.LocalUserAppDataPath()
                RetrieveStudioFromXML = SaveStudioToXML(True)
            End If

            '*** New XML Doc for XML file
            Dim XMLDoc As New Xml.XmlDocument
            XMLDoc.Load(StudioXMLFullPath)

            If XMLDoc.HasChildNodes = True Then
                curNode = XMLDoc.LastChild
                Dim TempStr As String = ""
                For Each nd As XmlNode In curNode.ChildNodes
                    If nd.InnerText <> "" Then
                        TempStr = nd.InnerText
                    Else
                        TempStr = ""
                    End If
                    Select Case nd.Name
                        Case "Winstate"
                            Select Case TempStr
                                Case "Maximized"
                                    WinState = FormWindowState.Maximized
                                Case "Normal"
                                    WinState = FormWindowState.Normal
                                Case Else
                                    WinState = FormWindowState.Normal
                            End Select
                        Case "AppDataPath"
                            AppDataPath = TempStr
                        Case "CCurl"
                            CCurl = TempStr

                    End Select
                Next
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modXML RetrieveStudioFromXML")
            Return False
        End Try

    End Function

#End Region

End Module