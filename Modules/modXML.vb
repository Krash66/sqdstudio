Imports System.Data
Imports System.Xml
Imports System.Xml.XPath

Public Module modXML

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

                    objFld.SetFieldAtt(xml_node.Attributes("Att").Value)
                    objFld.FieldName = xml_node.Attributes("ID").Value
                    objFld.OrgName = IIf(xml_node.Attributes("OrgName") Is Nothing, "", xml_node.Attributes("OrgName").Value)
                    objFld.ParentName = xml_node.Attributes("Parent").Value
                    Return objFld
                    Exit Function

                Case modDeclares.enumXMLType.XMLFOR_SQFUNCTIONS
                    If xml_node.Attributes("NODETYPE").Value <> "Heading" Then
                        Dim objFun As New clsSQFunction

                        If xml_node.Attributes("DESC") Is Nothing Then
                            objFun.SQFunctionDescription = ""
                        Else
                            objFun.SQFunctionDescription = xml_node.Attributes("DESC").Value
                        End If

                        objFun.SQFunctionSyntax = xml_node.Attributes("SYNTAX").Value
                        objFun.ParaCount = xml_node.Attributes("NPARM").Value
                        objFun.SQFunctionName = xml_node.Name

                        If xml_node.Attributes("NODETYPE").Value = "Template" Then
                            objFun.IsTemplate = True
                            Return objFun
                        ElseIf xml_node.Attributes("NODETYPE").Value = "Func" Then
                            Return objFun
                        End If
                    Else
                        Dim objFolder As New clsFolderNode(xml_node.Name, IIf(xml_node.Name = "Templates", NODE_FO_TEMPLATE, NODE_FO_FUNCTION))
                        Return objFolder
                        Exit Function
                    End If

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
                GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"   '& "\" 
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
                If System.IO.File.Exists(IO.Path.Combine(GetAppTemp(), "sqduiimp.ERR")) Then
                    System.IO.File.Delete(IO.Path.Combine(GetAppTemp(), "sqduiimp.ERR"))
                End If
                '// create new error file stream
                fsERR = System.IO.File.Create(IO.Path.Combine(GetAppTemp(), "sqduiimp.ERR"))
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
                        '//Now return output file path : <temp folder>\<fileprefix><input filename no extension>.XML
                        If StructType = modDeclares.enumStructure.STRUCT_COBOL_IMS Then
                            GetSQDumpXML = strTempDir & "\" & strFilePrefix & par1 & ".xml"    '& "\"
                        Else
                            GetSQDumpXML = strTempDir & "\" & strFilePrefix & GetFileNameWithoutExtenstionFromPath(strFileToParse) & ".xml"  '& "\"
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

    '// UnUsed???
    Function EncodeXMLFile(ByVal file As String) As Boolean
        Try
            Dim s As String

            s = LoadTextFile(file)
            s = s.Replace("&", "&amp;")

            SaveTextFile(file, s)
        Catch ex As Exception
            LogError(ex)
        End Try
    End Function

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

End Module


