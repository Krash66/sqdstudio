Public Class frmXMLconv

    Dim InFilePath As String
    Dim OutFilePath As String
    Dim xml_Indoc As New XmlDocument
    Dim xml_OutDoc As New XmlDocument
    Dim ArrAllElements As New ArrayList
    Dim ArrParentNodes As New ArrayList
    Dim ArrPrintedChildren As New ArrayList
    Dim ArrCDataNodes As New ArrayList

    Dim sb As New System.Text.StringBuilder

    Enum enumXMLActionType
        Failed = 0
        E_PrintwCldrn = 1
        E_PrintAsCdata = 2
        E_Ignore = 3
    End Enum

    Public Function OpenForm() As Boolean

        Try
            cmdOk.Enabled = False
            btnConv.Enabled = False
            btnbrowseOut.Enabled = False

            Me.Show()

        Catch ex As Exception
            LogError(ex, "frmXMLconv OpenForm")
            Return False
        End Try

    End Function

#Region "Load Input XML Message"

    Private Sub btnbrowseIn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbrowseIn.Click

        Try
            OFD1.Title = "XML Message"
            OFD1.ShowDialog()

        Catch ex As Exception
            LogError(ex, "frmXMLconv btnbrowseIn_Click")
        End Try
    End Sub

    Private Sub OFD1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OFD1.FileOk

        Try
            InFilePath = OFD1.FileName
            If InFilePath <> "" Then
                txtInPath.Text = InFilePath
                LoadInText()
                btnConv.Enabled = True
            End If

        Catch ex As Exception
            LogError(ex, "frmXMLconv OFD1_FileOk")
        End Try

    End Sub

    Sub LoadInText()

        ' Load the XML document in to the Text Window
        Try
            txtInMessage.Text = LoadTextFile(InFilePath)

        Catch ex As Exception
            LogError(ex, "frmXMLconv LoadInText")
        End Try

    End Sub

#End Region

#Region "Convert to DTD"

    Private Sub btnConv_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConv.Click

        Try
            ArrAllElements.Clear()
            ArrCDataNodes.Clear()

            'Convert xml_doc to new string builder text
            'EncodeXMLFile(InFilePath) '//encode some special characters like & to &amp;
            xml_Indoc.Load(InFilePath)

            'Start processing each node in the Message
            If xml_Indoc.HasChildNodes = True Then
                For Each nd As XmlNode In xml_Indoc.ChildNodes
                    '*** Process each Node, if it is an element
                    '*** if it's not an element, ignore it
                    If nd.NodeType = XmlNodeType.Element Then
                        processNode(nd)
                    End If
                Next 'next doc child
            End If

            If printCData(ArrCDataNodes) = True Then
                txtXMLout.Text = sb.ToString
                btnbrowseOut.Enabled = True
            End If

        Catch ex As Exception
            LogError(ex, "frmXMLconv btnConv_Click")
        End Try

    End Sub

    Private Function processNode(ByVal Node As XmlNode) As Boolean

        Try
            Dim Action As enumXMLActionType = GetNodeAction(Node)

            Select Case Action
                Case enumXMLActionType.E_PrintwCldrn

                    '*** Print the Element with it's children
                    printNodeWithChildren(Node)
                    '*** Now process the Children
                    For Each cld As XmlNode In Node.ChildNodes
                        'sb.AppendLine()
                        'sb.AppendLine("   ***Child = " & cld.Name)
                        'sb.AppendLine("   ***Of Parent = " & Node.Name)
                        'sb.AppendLine("   ***Sent to processNode")
                        'sb.AppendLine()
                        processNode(cld)
                    Next

                Case enumXMLActionType.E_PrintAsCdata
                    Dim QualName As String = GetCDataName(Node)
                    ArrCDataNodes.Add(QualName)
                    'ArrAllElements.Add(QualName)

                Case enumXMLActionType.E_Ignore
                    '*** Do Nothing except goto next sibling or parent

                Case enumXMLActionType.Failed
                    If MsgBox("Translation failed at Element: " & Node.Name & Chr(13) & _
                           "Would you like to continue processing?", MsgBoxStyle.YesNo, "Translation Failed") = MsgBoxResult.No Then
                        '*** abort here
                        Exit Function
                    End If

            End Select

            'If Node.NextSibling IsNot Nothing Then
            '    Node = Node.NextSibling
            '    sb.AppendLine()
            '    sb.AppendLine("   ***Next Sibling = " & Node.Name)
            '    sb.AppendLine()
            '    processNode(Node)
            '    'Else
            '    '    If Node.ParentNode IsNot Nothing Then
            '    '        Node = Node.ParentNode
            '    '        sb.AppendLine()
            '    '        sb.AppendLine("   ***Parent Node = " & Node.Name)
            '    '        sb.AppendLine()
            '    '    End If

            '    '    If Node.NextSibling IsNot Nothing Then
            '    '        Node = Node.NextSibling
            '    '        sb.AppendLine()
            '    '        sb.AppendLine("   ***Next Parent Sibling = " & Node.Name)
            '    '        sb.AppendLine()
            '    '        processNode(Node)
            '    '    End If
            'End If

            Return True

        Catch ex As Exception
            LogError(ex, "frmXMLconv processNode")
            Return False
        End Try

    End Function

    Private Function GetNodeAction(ByVal nd As XmlNode) As enumXMLActionType

        Try
            'sb.Append("   *** Type = " & nd.NodeType.ToString)
            'sb.Append("   LocalName = " & nd.LocalName)
            'sb.Append("   Name = " & nd.Name)
            'sb.Append("   Value = " & nd.Value)
            'sb.AppendLine("   # of Children = " & nd.ChildNodes.Count.ToString)

            Dim IsParentElement As Boolean = False

            If nd.HasChildNodes = True Then
                Dim NumEle As Integer = 0
                For Each node As XmlNode In nd.ChildNodes
                    If node.NodeType = XmlNodeType.Element Then
                        If node.HasChildNodes Then
                            If node.FirstChild.NodeType = XmlNodeType.Element Then
                                IsParentElement = True
                            End If
                        End If
                        NumEle += 1
                        If NumEle = 2 Then
                            'sb.AppendLine("   *** Node action E_PrintwCldrn ::" & nd.Name)
                            'sb.AppendLine()
                            GetNodeAction = enumXMLActionType.E_PrintwCldrn
                            Exit Function
                        End If
                    End If
                Next
                If IsParentElement = True Then
                    'sb.AppendLine("   *** Node action E_PrintwCldrn ::" & nd.Name)
                    'sb.AppendLine()
                    GetNodeAction = enumXMLActionType.E_PrintwCldrn
                Else
                    'sb.AppendLine("   *** Node action E_PrintAsCdata ::" & nd.Name)
                    'sb.AppendLine()
                    GetNodeAction = enumXMLActionType.E_PrintAsCdata
                End If
            Else
                If nd.NodeType = XmlNodeType.Element Then
                    'sb.AppendLine("   *** Node action E_PrintAsCdata ::" & nd.Name)
                    'sb.AppendLine()
                    GetNodeAction = enumXMLActionType.E_PrintAsCdata
                Else
                    'sb.AppendLine("   --- Node action E_Ignore ::" & nd.Name)
                    'sb.AppendLine()
                    GetNodeAction = enumXMLActionType.E_Ignore
                End If
            End If


        Catch ex As Exception
            LogError(ex, "frmXMLconv GetNodeAction")
            GetNodeAction = enumXMLActionType.Failed
            sb.AppendLine("   --- Node action Failed ::" & nd.Name)
            sb.AppendLine()
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
            LogError(ex, "frmXMLconv GetNameToAdd")
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
            LogError(ex, "frmXMLconv GetChildName")
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
            LogError(ex, "frmXMLconv GetCDataNameToAdd")
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
                sb.AppendLine(Line1)
                sb.AppendLine(Line2)
                For Each cld As XmlNode In node.ChildNodes
                    If cld.NodeType = XmlNodeType.Element Then
                        ChildElementName = GetChildName(cld)
                        Dim LineChild As String = String.Format("{0}{1}", Prefix, ChildElementName)
                        sb.AppendLine(LineChild)
                        Prefix = "    ,"
                        '*** Add children to the child array
                        ArrAllElements.Add(ChildElementName)
                        ArrPrintedChildren.Add(ChildElementName)
                    End If
                Next
                sb.AppendLine(LastLine)
                sb.AppendLine()

                Return True
            Else
                Return False
            End If


        Catch ex As Exception
            LogError(ex, "frmXMLconv printNode")
            Return False
        End Try

    End Function

    Private Function printCData(ByVal Arr As ArrayList) As Boolean

        Try
            Dim FORfld As String

            For Each name As String In Arr
                FORfld = String.Format("{0}{1,-40}{2}", "<!ELEMENT ", name, " (#CDATA)>")
                sb.AppendLine(FORfld)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "frmXMLconv printCData")
            Return False
        End Try

    End Function

#End Region

#Region "Save Output DTD file"

    Private Sub btnbrowseOut_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbrowseOut.Click

        Try
            SFD1.Title = "XML DTD File"
            SFD1.ShowDialog()

        Catch ex As Exception
            LogError(ex, "frmXMLconv btnbrowseOut_Click")
        End Try

    End Sub

    Private Sub SFD1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles SFD1.FileOk

        Try
            OutFilePath = SFD1.FileName
            txtOutPath.Text = OutFilePath
            If OutFilePath <> "" Then
                If Save() = True Then
                    'MsgBox("Save was successful", MsgBoxStyle.OkOnly)
                    cmdOk.Enabled = True
                End If
            End If

        Catch ex As Exception
            LogError(ex, "frmXMLconv SFD1_FileOk")
        End Try

    End Sub

    Private Sub cmdOk_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        'Me.Close()
        'If Save() = True Then
        '    MsgBox("Save was successful", MsgBoxStyle.OkOnly)
        'End If
        Try
            If OutFilePath <> "" Then
                Dim OpenProcess As New System.Diagnostics.Process
                Process.Start(OutFilePath)

                'Shell("notepad.exe " & RetCode.SQDPath, AppWinStyle.NormalFocus)
            End If

        Catch ex As Exception
            LogError(ex, "frmXMLconv cmdOk_Click_1")
        End Try

    End Sub

    '*** Save XML DTD to File
    Function Save() As Boolean

        Try
            If SFD1.FileName <> "" Then
                Save = SaveTextFile(OutFilePath, txtXMLout.Text)
            Else
                MsgBox("Please enter a valid Output File Path", MsgBoxStyle.Information, "No Valid Output Path")
            End If

            Save = True

        Catch ex As Exception
            LogError(ex, "frmXMLconv Save")
            Save = False
        End Try

    End Function

#End Region


    Private Sub cmdCancel_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()

    End Sub

    Private Sub cmdHelp_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdHelp.Click

        ShowHelp(HHId.H_Welcome_to_SQData_Studio)

    End Sub

End Class
