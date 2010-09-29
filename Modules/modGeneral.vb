Imports Microsoft.Win32

Public Module modGeneral

    'Private m_blnDSNUpdated As Boolean = False

    'Public Property DSNUpdated() As Boolean
    '    Get
    '        Return m_blnDSNUpdated
    '    End Get
    '    Set(ByVal Value As Boolean)
    '        m_blnDSNUpdated = Value
    '    End Set
    'End Property

#Region "Mylist Class"

    ''' <summary>
    ''' Simple list of objects and their names
    ''' </summary>
    ''' <remarks>Useful for small collections, dropdown lists, etc...</remarks>
    Public Class Mylist

        Private sName As String
        ' You can also declare this as String,bitmap or almost anything. 
        ' If you change this delcaration you will also need to change the Sub New 
        ' to reflect any change. Also the ItemData Property will need to be updated. 
        Private iID As Object

        ''' <summary>
        ''' Default empty constructor. 
        ''' </summary>
        ''' <remarks>Defines new list of objects</remarks>
        Public Sub New()
            sName = ""
            ' This would need to be changed if you modified the declaration above. 
            iID = 0
        End Sub

        ''' <summary>
        ''' Contructor with item definition
        ''' </summary>
        ''' <param name="Name">Text Name</param>
        ''' <param name="ID">Object</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal Name As String, ByVal ID As Object)
            sName = Name
            iID = ID
        End Sub

        Public Property Name() As String
            Get
                Return sName
            End Get
            Set(ByVal sValue As String)
                sName = sValue
            End Set
        End Property

        ' This is the property that holds the extra data. 
        Public Property ItemData() As Object
            Get
                Return iID
            End Get
            Set(ByVal iValue As Object)
                iID = iValue
            End Set
        End Property

        ' This is neccessary because the ListBox and ComboBox rely 
        ' on this method when determining the text to display. 
        Public Overrides Function ToString() As String
            Return sName
        End Function

    End Class

#End Region

#Region "MyIntlist Class"

    Public Class MyIntlist

        '/// created by tk 6/07 for integer key based "MyList Class"
        '/// Used for dropdowns, etc. that need a number key to
        '/// use for array continuity
        Private sName As Integer
        ' You can also declare this as String,bitmap or almost anything. 
        ' If you change this delcaration you will also need to change the Sub New 
        ' to reflect any change. Also the ItemData Property will need to be updated. 
        Private iID As Object

        Public Sub New()
            sName = 0
            ' This would need to be changed if you modified the declaration above. 
            iID = 0
        End Sub

        Public Sub New(ByVal Idx As Integer, ByVal ID As Object)
            sName = Idx
            iID = ID
        End Sub

        Public Property Idx() As Integer
            Get
                Return sName
            End Get
            Set(ByVal sValue As Integer)
                sName = sValue
            End Set
        End Property

        ' This is the property that holds the extra data. 
        Public Property ItemData() As Object
            Get
                Return iID
            End Get
            Set(ByVal iValue As Object)
                iID = iValue
            End Set
        End Property

        ' This is neccessary because the ListBox and ComboBox rely 
        ' on this method when determining the text to display. 
        Public Function GetIdx() As Integer
            Return sName
        End Function

    End Class

#End Region

#Region "Main Program Functions"

    Function ShowStatusMessage(ByVal sMsg As String) As Boolean
        mainwindow.StatusBar1.Panels(0).Text = sMsg
        mainwindow.StatusBar1.Refresh()
    End Function

    '//When user drag item on the last visible item of listview then next item will become visible
    '//When user drag item on the first visible item of listview then previous item will become visible
    Function ScrollTONextOrPrevItem(ByVal lvItm As ListViewItem) As Boolean
        If lvItm.ListView.TopItem.Index = lvItm.Index And lvItm.Index > 0 Then
            lvItm.ListView.Items(lvItm.Index - 1).EnsureVisible()
        ElseIf lvItm.Index < lvItm.ListView.Items.Count - 1 Then
            lvItm.ListView.Items(lvItm.Index + 1).EnsureVisible()
        End If
    End Function

    '//Note: Only work id list item set set using MyList class
    '//i.e combo.Items.Add(New MyList("AAAA",1))
    Function SetListItemByValue(ByVal cmb As ComboBox, ByVal Value As Object, ByVal canEdit As Boolean) As Boolean

        Dim i As Integer
        Dim oldIsEventFromCode As Boolean

        Try
            oldIsEventFromCode = IsEventFromCode
            SetListItemByValue = False

            For i = 0 To cmb.Items.Count - 1
                If CType(cmb.Items(i), Mylist).ItemData = Value Then
                    IsEventFromCode = True
                    cmb.SelectedIndex = i
                    IsEventFromCode = False
                    SetListItemByValue = True
                    Exit For
                End If
            Next
            If SetListItemByValue = False Then
                If Value.ToString <> "" And canEdit = True Then
                    cmb.Items.Add(New Mylist(Value.ToString, Value))
                    SetListItemByValue(cmb, Value, False)
                    SetListItemByValue = True
                    Exit Try
                Else
                    cmb.SelectedIndex = 0
                    IsEventFromCode = oldIsEventFromCode
                    Exit Function
                End If
            End If
            IsEventFromCode = oldIsEventFromCode

            SetListItemByValue = True

        Catch ex As Exception
            LogError(ex, cmb.Name, "modGeneral SetListItemByValue")
            SetListItemByValue = False
        End Try

    End Function

    '//Get Registered Owner name on current machine
    Function GetOwner() As String

        Dim regOwner As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")

        GetOwner = regOwner.GetValue("RegisteredOwner")

    End Function

    '//Get Registered company name on current machine
    Function GetCompany() As String

        Dim regCompany As RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")

        GetCompany = regCompany.GetValue("RegisteredOrganization")

    End Function

    Public Function SaveTextFile(ByVal FilePath As String, ByVal FileContent As String, Optional ByVal Append As Boolean = False) As Boolean

        Dim sw As System.IO.StreamWriter = Nothing

        Try
            sw = New System.IO.StreamWriter(FilePath, Append)
            sw.Write(FileContent)
            Return True

        Catch ex As Exception
            LogError(ex)
            Return False
        Finally
            If Not sw Is Nothing Then sw.Close()
        End Try

    End Function

    Function LoadTextFile(ByVal FilePath As String) As String

        Dim sr As System.IO.StreamReader = Nothing

        Try
            sr = New System.IO.StreamReader(FilePath)
            LoadTextFile = sr.ReadToEnd()

        Finally
            If Not sr Is Nothing Then sr.Close()
        End Try

    End Function

    Public Function ImgIdxFromName(ByVal Name As String, Optional ByVal obj As INode = Nothing) As Integer

        Select Case Name
            Case NODE_FOLDEROPEN
                ImgIdxFromName = 1
            Case NODE_PROJECT
                ImgIdxFromName = 2
            Case NODE_ENVIRONMENT
                ImgIdxFromName = 3
            Case NODE_STRUCT
                ImgIdxFromName = 4
            Case NODE_STRUCT_SEL
                ImgIdxFromName = 29
            Case NODE_SYSTEM
                ImgIdxFromName = 5
            Case NODE_ENGINE
                ImgIdxFromName = 6
            Case NODE_CONNECTION
                ImgIdxFromName = 7
            Case NODE_GEN
                ImgIdxFromName = 8
            Case NODE_LOOKUP
                ImgIdxFromName = 9
            Case NODE_MAIN
                ImgIdxFromName = 10
            Case NODE_PROC
                ImgIdxFromName = 11
            Case NODE_SOURCEDATASTORE
                If obj IsNot Nothing Then
                    If CType(obj, clsDatastore).IsLookUp = True Then
                        ImgIdxFromName = 9
                    Else
                        ImgIdxFromName = 12
                    End If
                Else
                    ImgIdxFromName = 12
                End If
            Case NODE_SOURCEDSSEL
                ImgIdxFromName = 29
            Case NODE_TARGETDATASTORE
                ImgIdxFromName = 13
            Case NODE_TARGETDSSEL
                ImgIdxFromName = 29
            Case NODE_VARIABLE
                ImgIdxFromName = 14
            Case NODE_FUN
                ImgIdxFromName = 31
            Case NODE_STRUCT_FLD
                ImgIdxFromName = 32
            Case NODE_FO_TEMPLATE
                ImgIdxFromName = 33
            Case NODE_TEMPLATE
                ImgIdxFromName = 34
            Case NODE_FO_FUNCTION_RECENT
                ImgIdxFromName = 35
            Case NODE_CONST
                ImgIdxFromName = 36
            Case Else
                ImgIdxFromName = 0
        End Select

    End Function

    Function FixStr(Optional ByVal s As String = "") As String

        If s Is Nothing Then
            Return ""
        Else
            FixStr = s.Replace("'", "''")
        End If

    End Function

    Function Quote(ByVal s As String, Optional ByVal ch As String = "'") As String

        Quote = ch & s & ch

    End Function

    Function NChr(ByVal n As Integer, Optional ByVal ch As String = "-") As String

        Dim i As Integer

        NChr = ""
        For i = 1 To n
            NChr = NChr & ch
        Next

    End Function

    Public Sub ClearControls(ByVal cControls As Windows.Forms.Control.ControlCollection, _
        Optional ByVal ExcludeControlType1 As enumControlTypes = modDeclares.enumControlTypes.ControlUnknown, _
        Optional ByVal ExcludeControlType2 As enumControlTypes = modDeclares.enumControlTypes.ControlUnknown, _
        Optional ByVal ExcludeControlType3 As enumControlTypes = modDeclares.enumControlTypes.ControlUnknown, _
        Optional ByVal ExcludeControlType4 As enumControlTypes = modDeclares.enumControlTypes.ControlUnknown)

        For Each c As Control In cControls
            If TypeOf c Is TextBox Then

                If ExcludeControlType1 <> modDeclares.enumControlTypes.ControlTextBox And _
                     ExcludeControlType2 <> modDeclares.enumControlTypes.ControlTextBox And _
                      ExcludeControlType3 <> modDeclares.enumControlTypes.ControlTextBox And _
                       ExcludeControlType4 <> modDeclares.enumControlTypes.ControlTextBox Then

                    CType(c, TextBox).Text = ""
                End If

            ElseIf TypeOf c Is ComboBox Then

                If ExcludeControlType1 <> modDeclares.enumControlTypes.ControlComboBox And _
                    ExcludeControlType2 <> modDeclares.enumControlTypes.ControlComboBox And _
                    ExcludeControlType3 <> modDeclares.enumControlTypes.ControlComboBox And _
                    ExcludeControlType4 <> modDeclares.enumControlTypes.ControlComboBox Then

                    If CType(c, ComboBox).Items.Count > 0 Then CType(c, ComboBox).SelectedIndex = 0
                End If

            ElseIf TypeOf c Is ListView Then

                If ExcludeControlType1 <> modDeclares.enumControlTypes.ControlListView And _
                    ExcludeControlType2 <> modDeclares.enumControlTypes.ControlListView And _
                    ExcludeControlType3 <> modDeclares.enumControlTypes.ControlListView And _
                    ExcludeControlType4 <> modDeclares.enumControlTypes.ControlListView Then

                    CType(c, ListView).Items.Clear()
                End If

            ElseIf TypeOf c Is TreeView Then

                If ExcludeControlType1 <> modDeclares.enumControlTypes.ControlTreeView And _
                    ExcludeControlType2 <> modDeclares.enumControlTypes.ControlTreeView And _
                    ExcludeControlType3 <> modDeclares.enumControlTypes.ControlTreeView And _
                    ExcludeControlType4 <> modDeclares.enumControlTypes.ControlTreeView Then

                    CType(c, TreeView).Nodes.Clear()
                End If

            ElseIf TypeOf c Is CheckBox Then
                CType(c, CheckBox).Checked = False

            ElseIf TypeOf c Is ListBox Then

                If CType(c, ListBox).Items.Count > 0 Then CType(c, ListBox).SelectedIndex = 0
            End If

            If c.Controls.Count > 0 Then ClearControls(c.Controls)
        Next

    End Sub

    Function GetCharFromPos(ByVal txt As TextBox, ByVal pt As Point) As Integer
        '// Convert the point into a DWord with horizontal position
        '// in the loword and vertical position in the hiword:
        Dim xy As Integer = (pt.X And &HFFFF) + ((pt.Y And &HFFFF) << 16)
        '// Get the position from the text box.
        Dim res As Integer = SendMessage(txt.Handle, EM_CHARFROMPOS, 0, xy)
        '// the Platform SDK appears to be incorrect on this matter.
        '// the hiword is the line number and the loword is the index
        '// of the character on this line
        Dim lineNumber As Integer = ((res And &HFFFF) >> 16)
        Dim charIndex As Integer = (res And &HFFFF)

        '// Find the index of the first character on the line within 
        '// the control:
        Dim lineStartIndex As Integer = SendMessage(txt.Handle, EM_LINEINDEX, lineNumber, 0)
        '// Return the combined index:

        Return lineStartIndex + charIndex

    End Function

    'Function GetCharFromPosRTF(ByVal txt As RichTextBox, ByVal pt As Point) As Integer
    '    '// Convert the point into a DWord with horizontal position
    '    '// in the loword and vertical position in the hiword:
    '    Dim xy As Integer = (pt.X And &HFFFF) + ((pt.Y And &HFFFF) << 16)
    '    '// Get the position from the text box.
    '    Dim res As Integer = SendMessage(txt.Handle, EM_CHARFROMPOS, 0, xy)
    '    '// the Platform SDK appears to be incorrect on this matter.
    '    '// the hiword is the line number and the loword is the index
    '    '// of the character on this line
    '    Dim lineNumber As Integer = ((res And &HFFFF) >> 16)
    '    Dim charIndex As Integer = (res And &HFFFF)

    '    '// Find the index of the first character on the line within 
    '    '// the control:
    '    Dim lineStartIndex As Integer = SendMessage(txt.Handle, EM_LINEINDEX, lineNumber, 0)
    '    '// Return the combined index:

    '    Return lineStartIndex + charIndex

    'End Function

    Sub ShowCustomCaret(ByVal ctl As Windows.Forms.Control, _
                        Optional ByVal width As Integer = 3, _
                        Optional ByVal height As Integer = 16, _
                        Optional ByVal CaretBlinkTimeinMs As Integer = 400)

        On Error Resume Next
        With ctl
            CreateCaret(.Handle.ToInt32, 0, width, height)
            ShowCaret(.Handle.ToInt32)

            Debug.Write("Current blinktime : " & GetCaretBlinkTime)
            SetCaretBlinkTime(CaretBlinkTimeinMs)
        End With

    End Sub

#End Region

#Region "Generate Scripts & Models"

    'InType=STR,SEL 
    'OutType=DTD, H, DDL
    'Function ModelScript(ByVal MetaDSN As String, ByVal Id As String, ByVal FileName As String, Optional ByVal InType As String = "STR", Optional ByVal OutType As String = "DTD", Optional ByVal SavePath As String = "", Optional ByVal StrName As String = "", Optional ByVal prefix As String = "") As String

    '    Dim args As String
    '    Dim si As New System.Diagnostics.ProcessStartInfo
    '    Dim myProcess As System.Diagnostics.Process
    '    Dim ext As String
    '    Dim makeModel As Boolean
    '    Dim PhysScript As String
    '    Dim TempPath As String = GetAppTemp()


    '    If SavePath = "" Then SavePath = GetAppPath()

    '    Select Case OutType
    '        Case "DTD"
    '            ext = "dtd"
    '        Case "DDL"
    '            ext = "ddl"
    '        Case "H"
    '            ext = "h"
    '        Case "LOD"
    '            ext = "lod"
    '        Case "SQL"
    '            ext = "sql"
    '        Case "MSSQL"
    '            ext = "sql"
    '        Case Else
    '            ext = OutType
    '    End Select

    '    If Strings.Right(TempPath, 1) <> "\" Then
    '        PhysScript = TempPath & "\" & StrName & "." & ext
    '    Else
    '        PhysScript = TempPath & StrName & "." & ext
    '    End If

    '    If Strings.Right(SavePath, 1) <> "\" Then
    '        ModelScript = SavePath & "\" & FileName & "." & ext
    '    Else
    '        ModelScript = SavePath & FileName & "." & ext
    '    End If

    '    makeModel = True

    '    If IO.File.Exists(ModelScript) = True Then
    '        If MsgBox("Model file " & ModelScript & " already exists.  Would you like to replace it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.No Then
    '            makeModel = False
    '            ModelScript = ""
    '        End If
    '    End If

    '    If makeModel = True Then
    '        args = MetaDSN & " " & Id & " " & InType & " " & OutType & " " & Quote(TempPath, """") & " " & prefix
    '        Log("********* Model Script Generation Start *********")
    '        Log("SQDUMODL.EXE " & MetaDSN & " " & Id & " " & InType & " " & OutType & " " & Quote(TempPath, """") & " " & prefix)

    '        '//run out little exe with command line args so it produces meta data in XML format

    '        Try
    '            si.CreateNoWindow = True
    '            si.WindowStyle = ProcessWindowStyle.Hidden

    '            '#If CONFIG = "ETI" Then
    '            '                si.FileName = "ETIUMODL.EXE"
    '            '#Else
    '            si.FileName = "SQDUMODL.EXE"
    '            '#End If

    '            si.Arguments = args

    '            myProcess = System.Diagnostics.Process.Start(si)

    '            '//wait until task is done
    '            myProcess.WaitForExit()
    '            'Log("Parser Run Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")

    '            'Select Case myProcess.ExitCode
    '            '    Case 8
    '            '        
    '            '    Case 4
    '            '       

    '            '    Case 1
    '            '        

    '            '    Case Is > 0
    '            '        

    '            '    Case 0
    '            '       

    '            '    Case Else
    '            '        
    '            'End Select

    '            Log("********* Modeler Return Code = " & myProcess.ExitCode & " *********")
    '            myProcess.Close()

    '            '// New script comes back with original structure name 
    '            '// so we will change the file name here
    '            If IO.File.Exists(PhysScript) Then
    '                IO.File.Copy(PhysScript, ModelScript, True)
    '            Else
    '                ModelScript = ""
    '            End If

    '        Catch ex As Exception
    '            LogError(ex, "modGeneral ModelScript")
    '            Return ""
    '        End Try

    '        Log("********* Model Script Generation End *********")
    '    End If

    'End Function

    Function ModelScript(ByVal MetaDSN As String, ByVal Id As String, ByVal FileName As String, Optional ByVal InType As String = "STR", Optional ByVal OutType As String = "DTD", Optional ByVal SavePath As String = "", Optional ByVal StrName As String = "", Optional ByVal prefix As String = "") As String

        Try
            Dim args As String
            'Dim myProcess As System.Diagnostics.Process
            Dim ext As String
            Dim makeModel As Boolean
            Dim PhysScript As String
            Dim TempPath As String = GetAppTemp()
            Dim fsERR As System.IO.FileStream
            Dim objWriteERR As System.IO.StreamWriter
            Dim PathErr As String


            If SavePath = "" Then SavePath = GetAppTemp()

            Select Case OutType
                Case "DTD"
                    ext = "dtd"
                Case "DDL"
                    ext = "ddl"
                Case "H"
                    ext = "h"
                Case "LOD"
                    ext = "lod"
                Case "SQL"
                    ext = "sql"
                Case "MSSQL"
                    ext = "sql"
                Case Else
                    ext = OutType
            End Select

            If Strings.Right(TempPath, 1) <> "\" Then
                PhysScript = TempPath & "\" & StrName & "." & ext
            Else
                PhysScript = TempPath & StrName & "." & ext
            End If

            If Strings.Right(SavePath, 1) <> "\" Then
                ModelScript = SavePath & "\" & FileName & "." & ext
            Else
                ModelScript = SavePath & FileName & "." & ext
            End If

            makeModel = True

            If IO.File.Exists(ModelScript) = True Then
                If MsgBox("Model file " & ModelScript & " already exists.  Would you like to replace it?", MsgBoxStyle.Question Or MsgBoxStyle.YesNo, MsgTitle) = MsgBoxResult.No Then
                    makeModel = False
                    ModelScript = ""
                End If
            End If

            If makeModel = True Then
                args = MetaDSN & " " & Id & " " & InType & " " & OutType & " " & Quote(TempPath, """") & " " & prefix

                '//run out little exe with command line args so it produces meta data in XML format
                Try
                    '//delete previous log file
                    If System.IO.File.Exists(IO.Path.Combine(GetAppTemp(), "sqdumodl.ERR")) Then
                        System.IO.File.Delete(IO.Path.Combine(GetAppTemp(), "sqdumodl.ERR"))
                    End If
                    '// create new error file stream
                    fsERR = System.IO.File.Create(IO.Path.Combine(GetAppTemp(), "sqdumodl.ERR"))
                    PathErr = fsERR.Name
                    objWriteERR = New System.IO.StreamWriter(fsERR)

                    Dim si As New ProcessStartInfo()

                    si.FileName = "SQDUMODL.EXE "
                    si.Arguments = args
                    si.UseShellExecute = False
                    si.CreateNoWindow = True

                    '// Redirect Standard Error. Let standard Output go  
                    si.RedirectStandardOutput = False
                    si.RedirectStandardError = True

                    '// Create a new process to Model new Description Files
                    Using myProcess As New System.Diagnostics.Process()
                        myProcess.StartInfo = si

                        Log("********* Model Script Generation Start *********")
                        Log("Modeler Run Started : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                        Log(si.FileName & args)

                        myProcess.Start()

                        Dim OutStr As String = ""
                        Dim ErrStr As String = ""

                        '/// split output into multiple threads to capture each stream to a string
                        '/// OutStr stays as "" because StdOut is NOT redirected
                        OutputToEnd(myProcess, OutStr, ErrStr)

                        '//wait until task is done
                        myProcess.WaitForExit()

                        Log("Modeler Run Ended : " & Date.Now & " & " & Date.Now.Millisecond & " Milliseconds")
                        Log("Modeler Returned Code : " & myProcess.ExitCode)

                        objWriteERR.Write(ErrStr)
                        objWriteERR.Close()
                        fsERR.Close()

                        If myProcess.ExitCode <> 0 Then
                            If MsgBox("Error occurred while Modeling the Description [" & StrName & "]." & vbCrLf & _
                                      "Do you want to see the log?", _
                                      MsgBoxStyle.Critical Or MsgBoxStyle.YesNoCancel, _
                                      MsgTitle) = MsgBoxResult.Yes Then
                                If IO.File.Exists(PathErr) Then
                                    Process.Start(PathErr)
                                End If
                            End If
                            Return ""
                        End If

                        Log("Modeler Report file saved at : " & PathErr)
                        Log("********* Modeler Return Code = " & myProcess.ExitCode & " *********")
                        myProcess.Close()

                    End Using

                    '// New script comes back with original structure name 
                    '// so we will change the file name here
                    If IO.File.Exists(PhysScript) Then
                        IO.File.Copy(PhysScript, ModelScript, True)
                    Else
                        ModelScript = ""
                    End If

                Catch ex As Exception
                    LogError(ex, "modGeneral ModelScript .. interior process")
                    Return ""
                End Try
            End If

        Catch ex As Exception
            LogError(ex, "modGeneral ModelScript")
            Return ""
        End Try

    End Function

    '/// Now Unused
    'Function GenerateEngineScript(ByVal MetaDSN As String, ByVal Id As String, ByVal Name As String, Optional ByVal SavePath As String = "", Optional ByVal Prefix As String = "") As String
    '    Dim args As String
    '    'Dim p As String
    '    Dim logfilepath As String

    '    If SavePath = "" Then SavePath = GetAppTemp()

    '    args = MetaDSN & " " & Id & " " & Quote(SavePath, """") & " " & Prefix

    '    Log("********* Script Generation Start *********")
    '    Log("SQDGNSQD.EXE " & args)

    '    '//run out little exe with command line args so it produces meta data in XML format
    '    Dim si As New System.Diagnostics.ProcessStartInfo
    '    Dim myProcess As System.Diagnostics.Process

    '    Try
    '        Dim filepath As String

    '        If Strings.Right(SavePath, 1) <> "\" Then
    '            filepath = SavePath & "\" & Name & ".sqd"
    '            logfilepath = SavePath & "\" & Name & ".rpt"
    '        Else
    '            filepath = SavePath & Name & ".sqd"
    '            logfilepath = SavePath & Name & ".rpt"
    '        End If

    '        '//delete previous rpt, sqd files
    '        If IO.File.Exists(filepath) Then
    '            IO.File.Delete(filepath)
    '        End If

    '        If IO.File.Exists(logfilepath) Then
    '            IO.File.Delete(logfilepath)
    '        End If

    '        si.CreateNoWindow = True
    '        si.WindowStyle = ProcessWindowStyle.Hidden
    '        si.FileName = GetAppPath() & "SQDGNSQD.EXE"
    '        si.Arguments = args

    '        '//Create a new process to parse input file and dump to XML
    '        myProcess = System.Diagnostics.Process.Start(si)

    '        '//wait until task is done
    '        myProcess.WaitForExit()


    '        Select Case myProcess.ExitCode
    '            Case 8
    '                If MsgBox("Script generated with errors. Would you like to view the result?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "Script Generation error") = MsgBoxResult.Yes Then
    '                    Shell("notepad.exe " & Quote(logfilepath, """"), AppWinStyle.NormalFocus) '"notepad.exe " & 
    '                End If
    '            Case 4
    '                If MsgBox("Script generated with warnings.  Would you like to view the result?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "Script Generation error") = MsgBoxResult.Yes Then
    '                    Shell("notepad.exe " & Quote(logfilepath, """"), AppWinStyle.NormalFocus) '"notepad.exe " & 
    '                End If
    '            Case 1
    '                MsgBox("Script Not Generated.  There is a problem with the PATH !!", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Script Generation error")

    '            Case Is > 0
    '                If MsgBox("Script generated with errors, do you want to see the log?", MsgBoxStyle.Critical Or MsgBoxStyle.YesNo, "Script Generation error") = MsgBoxResult.Yes Then
    '                    Shell("notepad.exe " & Quote(GetAppPath() & "sqdgnsqd.log", """"), AppWinStyle.NormalFocus) '"notepad.exe " & 
    '                End If
    '            Case 0

    '                Log("Script file saved at : " & filepath)
    '                Log("********* Script Generation End [Return Code=" & myProcess.ExitCode & "]*********")
    '                Return filepath
    '            Case Else
    '                MsgBox("return code >> " & myProcess.ExitCode.ToString & Chr(13) & "return path >> " & filepath)

    '        End Select

    '        Log("Error log file saved at : " & logfilepath)
    '        Log("********* Script Generation End [Return Code=" & myProcess.ExitCode & "]*********")

    '    Catch ex As Exception
    '        LogError(ex)
    '    End Try
    'End Function

    '/// splits system IO process into multiple threads for Asynchronous output streams

#End Region

#Region "Asynchronous Stream Read and Write"

    '/// New object to buffer Asynchonous data streams
    Private Class OutputToEndData

        Public thread As Threading.Thread
        Public Stream As IO.StreamReader
        Public Output As String
        Public EXCEPT As Exception

    End Class

    Sub OutputToEnd(ByVal p As Diagnostics.Process, ByRef StandardOutput As String, ByRef StandardError As String)

        Dim outputData As New OutputToEndData
        Dim errorData As New OutputToEndData

        If p.StartInfo.RedirectStandardOutput = True Then
            outputData.Stream = p.StandardOutput
            outputData.thread = New Threading.Thread(AddressOf OutputToEndProc)
            outputData.thread.Start(outputData)
        End If

        If p.StartInfo.RedirectStandardError = True Then
            errorData.Stream = p.StandardError
            errorData.thread = New Threading.Thread(AddressOf OutputToEndProc)
            errorData.thread.Start(errorData)
        End If

        If p.StartInfo.RedirectStandardOutput = True Then
            outputData.thread.Join()
            StandardOutput = outputData.Output
        End If

        If p.StartInfo.RedirectStandardError = True Then
            errorData.thread.Join()
            StandardError = errorData.Output
        End If

        If outputData.EXCEPT IsNot Nothing Then
            Throw outputData.EXCEPT
        End If

        If errorData.EXCEPT IsNot Nothing Then
            Throw errorData.EXCEPT
        End If

    End Sub

    '/// Sub to capture data into outputToEndData class
    Private Sub OutputToEndProc(ByVal OutData As Object)
        Dim data = DirectCast(OutData, OutputToEndData)
        Try
            data.output = data.stream.readtoend
        Catch ex As Exception
            data.except = ex
        End Try
    End Sub


#End Region

    '/// All General Node and Tree Functions
#Region "Node Functions"

    '/// Added 4/12/07 by TK to HiLite Treenodes of Fields that have Field Descriptions
    Public Function HiLiteFieldDescNodes(ByVal hnode As TreeNode, Optional ByVal ChkTV As Boolean = False, Optional ByVal TV As TreeView = Nothing) As Boolean

        Dim i As Integer = 0

        Try
            If ChkTV = True And TV IsNot Nothing Then
                For i = 0 To TV.GetNodeCount(False) - 1
                    If CType(TV.Nodes(i).Tag, clsField).FieldDesc <> "" Then
                        TV.Nodes(i).ForeColor = Color.Blue
                    Else
                        TV.Nodes(i).ForeColor = Color.Black
                    End If
                Next
            End If

            For Each fNode As TreeNode In hnode.Nodes
                If CType(fNode.Tag, clsField).FieldDesc <> "" Then
                    fNode.ForeColor = Color.Blue
                Else
                    fNode.ForeColor = Color.Black
                End If
                HiLiteFieldDescNodes(fNode)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ModGeneral HiLiteFieldDescNodes")
            Return False
        End Try

    End Function

    '/// Added 5/07 by TK to HiLite Treenodes of Fields that have Foreign Keys
    Public Function HiLiteFieldFKeyNodes(ByVal hnode As TreeNode, Optional ByVal ChkTV As Boolean = False, Optional ByVal TV As TreeView = Nothing) As Boolean

        Dim i As Integer = 0
        Dim StartColor As Color

        Try
            If ChkTV = True And TV IsNot Nothing Then
                For i = 0 To TV.GetNodeCount(False) - 1
                    StartColor = TV.Nodes(i).ForeColor
                    If CType(TV.Nodes(i).Tag, clsField).GetFieldAttr(enumFieldAttributes.ATTR_FKEY) <> "" Then
                        TV.Nodes(i).ForeColor = Color.Red
                        TV.Nodes(i).ImageIndex = 0
                    Else
                        TV.Nodes(i).ForeColor = StartColor
                        TV.Nodes(i).ImageIndex = 1
                    End If
                Next
            End If

            For Each fNode As TreeNode In hnode.Nodes
                StartColor = fNode.ForeColor
                If CType(fNode.Tag, clsField).GetFieldAttr(enumFieldAttributes.ATTR_FKEY) <> "" Then
                    fNode.ForeColor = Color.Red
                    fNode.ImageIndex = 0
                Else
                    fNode.ForeColor = StartColor
                    fNode.ImageIndex = 1
                End If
                HiLiteFieldFKeyNodes(fNode)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ModGeneral HiLiteFieldFkeyNodes")
            Return False
        End Try

    End Function

    '/// Added 5/07 by TK to HiLite Treenodes of Fields that are Keys
    Public Function HiLiteFieldKeyNodes(ByVal hnode As TreeNode, Optional ByVal ChkTV As Boolean = False, Optional ByVal TV As TreeView = Nothing) As Boolean

        Dim i As Integer = 0

        Try
            If ChkTV = True And TV IsNot Nothing Then
                For i = 0 To TV.GetNodeCount(False) - 1
                    If CType(TV.Nodes(i).Tag, clsField).GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                        TV.Nodes(i).ImageIndex = 0
                    Else
                        TV.Nodes(i).ImageIndex = 1
                    End If
                Next
            End If

            For Each fNode As TreeNode In hnode.Nodes
                If CType(fNode.Tag, clsField).GetFieldAttr(enumFieldAttributes.ATTR_ISKEY) = "Yes" Then
                    fNode.ImageIndex = 0
                Else
                    fNode.ImageIndex = 1
                End If
                HiLiteFieldKeyNodes(fNode)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ModGeneral HiLiteFieldkeyNodes")
            Return False
        End Try

    End Function

    '/// Added 5/07 by TK to HiLite Treenodes of selections that are mapped in a procedure
    Public Function HiLiteMappedNodes(ByVal TV As TreeView, ByRef DS As clsDatastore) As Boolean

        Dim i As Integer = 0
        Dim ss, ssc As clsStructureSelection
        Dim flag As Boolean = False
        Dim cnodeObj As INode

        '/// Modified 9/26/2007
        Try
            For Each fnode As TreeNode In TV.Nodes
                For Each ssnode As TreeNode In fnode.Nodes
                    cnodeObj = CType(ssnode.Tag, INode)
                    If cnodeObj.Type = NODE_STRUCT_SEL Then
                        ss = CType(cnodeObj, clsStructureSelection)
                        For Each dssel As clsDSSelection In DS.ObjSelections
                            If dssel.IsMapped = True And dssel.SelectionName = ss.SelectionName Then
                                ssnode.ForeColor = Color.Green
                                flag = True
                                Exit For
                            End If
                        Next
                        If flag = False Then
                            ssnode.ForeColor = Color.Black
                        End If
                    End If
                    For Each ssCnode As TreeNode In ssnode.Nodes
                        cnodeObj = CType(ssCnode.Tag, INode)
                        If cnodeObj.Type = NODE_STRUCT_SEL Then
                            ssc = CType(cnodeObj, clsStructureSelection)
                            For Each dssel As clsDSSelection In DS.ObjSelections
                                If dssel.IsMapped = True And dssel.SelectionName = ssc.SelectionName Then
                                    ssCnode.ForeColor = Color.Green
                                    flag = True
                                    Exit For
                                End If
                            Next
                            If flag = False Then
                                ssCnode.ForeColor = Color.Black
                            End If
                        End If
                    Next
                Next
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ModGeneral HiLiteMappedNodes")
            Return False
        End Try

    End Function

    '/// Added 3/08 by TK to Hilite Last Mapped Field for each Procedure
    Public Function HiLiteLastSrcTgtFlds(ByVal hnode As TreeNode, ByVal MatchName As String, Optional ByVal ChkTV As Boolean = False, Optional ByVal TV As TreeView = Nothing) As Boolean

        Dim i As Integer = 0

        Try
            If ChkTV = True And TV IsNot Nothing Then
                For i = 0 To TV.GetNodeCount(False) - 1
                    'If CType(TV.Nodes(i).Tag, clsField).FieldName = MatchName Then
                    '    TV.Nodes(i).BackColor = Color.Yellow
                    '    Return True
                    '    Exit Function
                    '    'Else
                    HiLiteLastSrcTgtFlds(TV.Nodes(i), MatchName)
                    'End If
                Next
            End If

            For Each fNode As TreeNode In hnode.Nodes
                If CType(fNode.Tag, INode).Type = NODE_STRUCT_FLD Then
                    If CType(fNode.Tag, clsField).ParentName & "." & CType(fNode.Tag, clsField).FieldName = MatchName Then
                        fNode.BackColor = Color.LightBlue
                        fNode.EnsureVisible()
                        Return True
                        Exit Function
                    End If
                End If
                HiLiteLastSrcTgtFlds(fNode, MatchName)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "ModGeneral HiLiteFieldkeyNodes")
            Return False
        End Try

    End Function

    '/// Adds nodes to a treenode (tree node collection)
    Public Function AddNode(ByRef pNode As TreeNode, ByVal NodeType As String, ByVal obj As INode, Optional ByVal NodeVisible As Boolean = True, Optional ByVal NodeText As String = "") As TreeNode

        Try
            Dim n As New TreeNode

            '//Make folder node in bold all other in regular font
            If Strings.Left(NodeType, 3) = "FO_" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Bold)
            ElseIf NodeType <> "" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Regular)
            End If

            obj.ObjTreeNode = n
            If NodeText <> "" Then
                n.Text = NodeText
            Else
                n.Text = obj.Text
            End If

            n.Tag = obj

            n.ImageIndex = ImgIdxFromName(NodeType, obj)
            '//TODO: You can change this logic to show open folder for folder node
            n.SelectedImageIndex = n.ImageIndex
            pNode.Nodes.Add(n)

            '/// Update Parent Folder Item Count (Number of Children or members)
            If n.Parent IsNot Nothing Then
                If CType(n.Parent.Tag, INode).IsFolderNode = True Then
                    UpdateParentNodeLabel(n)
                End If
            End If

            If NodeVisible = True Then
                n.EnsureVisible()
            Else
                n.Collapse()
            End If

            AddNode = n

        Catch ex As Exception
            LogError(ex, "modGeneral AddNode")
            Return Nothing
        End Try

    End Function

    '/// Adds nodes to a treeview (tree node collection)
    Public Function AddTreeNode(ByRef pNode As TreeView, ByVal NodeType As String, ByVal obj As INode, Optional ByVal NodeVisible As Boolean = True, Optional ByVal NodeText As String = "") As TreeNode

        Try
            Dim n As New TreeNode

            '//Make folder node in bold all other in regular font
            If Strings.Left(NodeType, 3) = "FO_" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Bold)
            ElseIf NodeType <> "" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Regular)
            End If

            obj.ObjTreeNode = n
            If NodeText <> "" Then
                n.Text = NodeText
            Else
                n.Text = obj.Text
            End If

            n.Tag = obj

            n.ImageIndex = ImgIdxFromName(NodeType, obj)
            '//TODO: You can change this logic to show open folder for folder node
            n.SelectedImageIndex = n.ImageIndex
            pNode.Nodes.Add(n)

            '/// Update Parent Folder Item Count (Number of Children or members)
            If n.Parent IsNot Nothing Then
                UpdateParentNodeLabel(n)
            End If

            If NodeVisible = True Then
                n.EnsureVisible()
            Else
                n.Collapse()
            End If

            AddTreeNode = n

        Catch ex As Exception
            LogError(ex, "modGeneral AddTreeNode")
            Return Nothing
        End Try

    End Function

    '/// Adds nodes to a treenodecollection (tree node collection)
    Public Function AddNodeToCol(ByRef pNode As TreeNodeCollection, ByVal NodeType As String, ByVal obj As INode, Optional ByVal NodeVisible As Boolean = True, Optional ByVal NodeText As String = "") As TreeNode

        Try
            Dim n As New TreeNode

            '//Make folder node in bold all other in regular font
            If Strings.Left(NodeType, 3) = "FO_" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Bold)
            ElseIf NodeType <> "" Then
                n.NodeFont = New Font("Arial", 8, FontStyle.Regular)
            End If

            obj.ObjTreeNode = n
            If NodeText <> "" Then
                n.Text = NodeText
            Else
                n.Text = obj.Text
            End If

            n.Tag = obj

            n.ImageIndex = ImgIdxFromName(NodeType, obj)
            '//TODO: You can change this logic to show open folder for folder node
            n.SelectedImageIndex = n.ImageIndex
            pNode.Add(n)

            If NodeVisible = True Then
                n.EnsureVisible()
            Else
                n.Collapse()
            End If

            AddNodeToCol = n

        Catch ex As Exception
            LogError(ex, "modGeneral AddNodeToCol")
            Return Nothing
        End Try

    End Function

    Function UpdateNode(ByRef pNode As TreeNode, ByVal obj As INode) As Boolean

        Dim tempobj As INode
        Dim objDS As clsDatastore


        Try
            tempobj = pNode.Tag

            pNode.Text = obj.Text

            UpdateParentNodeLabel(pNode)

            '// test to see if it's a Datastore Node then add strSelections under that node
            If (tempobj.Type = NODE_SOURCEDATASTORE) Or (tempobj.Type = NODE_TARGETDATASTORE) Then
                objDS = CType(pNode.Tag, clsDatastore)
                pNode.Text = objDS.DsPhysicalSource
                objDS.SetDSselParents()
                AddDSstructuresToTree(objDS.ObjTreeNode, objDS)

            End If

            '// if DS selection then see if parent is a datastrore and not a structure
            '// then add selections under Datastore node
            If ((tempobj.Type = NODE_SOURCEDSSEL) Or (tempobj.Type = NODE_TARGETDSSEL)) And ((obj.Type = NODE_SOURCEDATASTORE) Or (obj.Type = NODE_TARGETDATASTORE)) Then
                objDS = CType(obj, clsDatastore)
                objDS.SetDSselParents()
                AddDSstructuresToTree(objDS.ObjTreeNode, objDS)
            End If

            pNode.Tag = obj

            Return True

        Catch ex As Exception
            LogError(ex, "UpdateNode=>" & ex.Message)
        End Try

    End Function

    Public Function UpdateParentNodeLabel(ByRef cNode As TreeNode) As Boolean

        Try
            If cNode IsNot Nothing Then
                If cNode.Parent IsNot Nothing Then
                    If cNode.Parent.Text.StartsWith("(") Then
                        '/// add 1 to node counter
                        Dim c As Integer
                        Dim str As String

                        c = cNode.Parent.Text.IndexOf(")")
                        str = cNode.Parent.Text
                        str = str.Remove(1, c - 1)
                        str = str.Insert(1, cNode.Parent.Nodes.Count.ToString)
                        cNode.Parent.Text = str
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral UpdateParentNodeLabel")
            Return False
        End Try

    End Function

    Public Function UpdateParentNodeCount(ByRef pNode As TreeNode) As Boolean

        Try
            If pNode IsNot Nothing Then
                If pNode.Text.StartsWith("(") Then
                    '/// add 1 to node counter
                    Dim c As Integer
                    Dim str As String

                    c = pNode.Text.IndexOf(")")
                    str = pNode.Text
                    str = str.Remove(1, c - 1)
                    str = str.Insert(1, pNode.Nodes.Count.ToString)
                    pNode.Text = str
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral UpdateParentNodeCount")
            Return False
        End Try

    End Function

    '//This function expand and selection first matching node in treeview
    '//This function can search Treeview by NodeText or by Node Key (Key is TreeNode->Tag->INode.key)
    '//Parameter : 
    '//     tv = treeview for which search should be performed
    '//     TextToSearch = If SearchForId is False then Search will be performed by node text each node will be compered to TextToSearch and match then node will be selected and expanded
    '//     SearchForId = If this is true then Search will be done by Node->Tag->INode.Key
    '//     IdToSearch = Node->Tag->INode.Key to search if SearchForId is true
    Public Function SelectFirstMatchingNode(ByVal tv As TreeView, ByVal colSkipNodes As ArrayList, Optional ByVal TextToSearch As String = "", Optional ByVal SearchForId As Boolean = False, Optional ByVal IdToSearch As Integer = 0) As Boolean

        Dim cnt As Integer = tv.GetNodeCount(True)

        If cnt <= 0 Then Exit Function
        Try
            If SearchTree(tv, colSkipNodes, False, tv.Nodes, TextToSearch, SearchForId, IdToSearch) Is Nothing Then
                colSkipNodes.Clear()
                Return False
            Else
                Return True
            End If

        Catch ex As Exception
            LogError(ex, "modGeneral SelectFirstMatchingNode")
            Return False
        End Try

    End Function

    Public Function SelectFirstMatchingNode(ByVal tv As TreeView, Optional ByVal TextToSearch As String = "", Optional ByVal SearchForId As Boolean = False, Optional ByVal IdToSearch As String = "", Optional ByVal ExactMatch As Boolean = False, Optional ByVal IgnoreSkipList As Boolean = True) As TreeNode

        Dim colSkipNodes As New ArrayList '//dummy array list
        Dim nd As TreeNode
        Dim cnt As Integer = tv.GetNodeCount(True)

        If cnt <= 0 Then
            Return Nothing
            Exit Function
        End If

        Try
            nd = SearchTree(tv, colSkipNodes, IgnoreSkipList, tv.Nodes, TextToSearch, SearchForId, IdToSearch, , ExactMatch)
            Return nd

        Catch ex As Exception
            LogError(ex, "modGeneral SlectFirstMatingNode-OL2")
            Return Nothing
        End Try

    End Function

    '//This function can be used for the following purposes
    '// 1) search first matching node by text
    '// 2) search first matching node by id which is stored in tag -> INode Object -> Key (Key is ID)
    '// Note: if used for second option then must pass SelectionId to search and SearchForCheckmark must be true
    Function SearchTree(ByVal tv As TreeView, ByVal colSkipNodes As ArrayList, ByVal IgnoreSkipList As Boolean, ByRef cNode As TreeNodeCollection, ByRef TextToSearch As String, Optional ByVal SearchForCheckmark As Boolean = False, Optional ByVal IdToSearch As String = "", Optional ByVal DoSelection As Boolean = True, Optional ByVal ExactMatch As Boolean = False) As TreeNode

        Try
            Dim nd As TreeNode
            Dim arrItm As Object
            Dim bSkip As Boolean
            Dim oldIsEventFromCode As Boolean = IsEventFromCode

            For Each nd In cNode
                If SearchForCheckmark = True Then
                    If CType(nd.Tag, INode).IsFolderNode = False Then
                        If CType(nd.Tag, INode).Key = IdToSearch Then
                            If DoSelection = True Then
                                '//This will check this node and parent node too
                                IsEventFromCode = True
                                nd.Checked = True
                                IsEventFromCode = False
                                nd.EnsureVisible()
                            End If
                            SearchTree = nd
                            Exit Function
                        End If
                    End If

                    '//Search all nodes under this node
                    nd = SearchTree(tv, colSkipNodes, IgnoreSkipList, nd.Nodes, TextToSearch, SearchForCheckmark, IdToSearch)
                    If Not (nd Is Nothing) Then
                        SearchTree = nd
                        Exit Function
                    End If
                Else
                    bSkip = False

                    If IgnoreSkipList = False Then
                        '//loop through skip list
                        For Each arrItm In colSkipNodes
                            If CType(arrItm.tag, INode).GUID = CType(nd.Tag, INode).GUID Then
                                bSkip = True
                            End If
                        Next
                    End If

                    If bSkip = False Then
                        If ExactMatch = True Then
                            If nd.Text = TextToSearch Then
                                If IgnoreSkipList = False Then
                                    colSkipNodes.Add(nd) '//add this node in skip list so next time when we search we dont compare nodes in skip list
                                End If

                                If DoSelection = True Then
                                    tv.SelectedNode = nd
                                    nd.EnsureVisible()
                                End If
                                SearchTree = nd
                                Exit Function
                            End If
                        Else
                            If Strings.Left(nd.Text, TextToSearch.Length).ToUpper = TextToSearch.ToUpper Then
                                If IgnoreSkipList = False Then
                                    colSkipNodes.Add(nd) '//add this node in skip list so next time when we search we dont compare nodes in skip list
                                End If

                                If DoSelection = True Then
                                    tv.SelectedNode = nd
                                    nd.EnsureVisible()
                                End If
                                SearchTree = nd
                                Exit Function
                            End If
                        End If

                    End If

                    '//Search all nodes under this node
                    nd = SearchTree(tv, colSkipNodes, IgnoreSkipList, nd.Nodes, TextToSearch, , , , True)
                    If Not (nd Is Nothing) Then
                        SearchTree = nd
                        Exit Function
                    End If
                End If
            Next

            IsEventFromCode = oldIsEventFromCode

            Return Nothing

        Catch ex As Exception
            LogError(ex, "modGeneral SearchTree")
            Return Nothing
        End Try

    End Function

    '//If onput node has parent then Check parent node
    '//if input node has childeren then check all
    '//Input: Node to be processed
    Function CheckUncheckNodes(ByVal cNode As TreeNode, Optional ByVal IsSingleSelect As Boolean = False) As Boolean

        Dim bState As Boolean = cNode.Checked
        Dim oldIsEventFromCode As Boolean
        oldIsEventFromCode = IsEventFromCode

        Try
            '//Process parent
            If Not (cNode.Parent Is Nothing) Then
                IsEventFromCode = True
                ChangeAllParentState(cNode, bState)
                IsEventFromCode = False
            End If

            '//Node we will take care all children
            CheckUnchecAllChildren(cNode.Nodes, bState)

        Catch ex As Exception
            LogError(ex, "modGeneral CheckUncheckNodes")
            '//TODO
        Finally
            IsEventFromCode = oldIsEventFromCode
        End Try

    End Function

    Private Function ChangeAllParentState(ByVal cNode As TreeNode, ByVal bState As Boolean) As Boolean

        Try
            If Not (cNode.Parent Is Nothing) Then
                If bState = True Then
                    cNode.Parent.Checked = bState
                    ChangeAllParentState(cNode.Parent, bState)
                Else
                    '//Only uncheck parent if all children unchecked
                    If IsAllChildrenSameState(cNode.Parent, bState) = True Then
                        cNode.Parent.Checked = bState
                        ChangeAllParentState(cNode.Parent, bState)
                    End If
                End If
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral ChangeAllParentState")
            Return False
        End Try

    End Function

    Private Function IsAllChildrenSameState(ByVal cNode As TreeNode, ByVal bState As Boolean) As Boolean

        Dim cnt As Long
        Dim nd As TreeNode

        Try
            For Each nd In cNode.Nodes
                If nd.Checked = bState Then
                    cnt = cnt + 1
                End If
            Next
            If cnt = cNode.GetNodeCount(False) Then
                IsAllChildrenSameState = True
            End If

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral IsAllChildrenSameState")
            Return False
        End Try

    End Function

    Private Function CheckUnchecAllChildren(ByRef cNode As TreeNodeCollection, ByVal bState As Boolean) As Boolean

        Dim c As TreeNode
        Dim oldIsEventFromCode As Boolean = IsEventFromCode

        Try
            '//Process children
            For Each c In cNode
                IsEventFromCode = True
                c.Checked = bState
                IsEventFromCode = False
                CheckUnchecAllChildren(c.Nodes, bState)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral CheckUncheckAllChildren")
            Return False
        Finally
            IsEventFromCode = oldIsEventFromCode
        End Try

    End Function

    Function GetNodeTypeFromSourceType(ByVal srcType As enumMappingType) As String

        Try
            Select Case srcType
                Case modDeclares.enumMappingType.MAPPING_TYPE_FIELD
                    Return NODE_STRUCT_FLD
                Case modDeclares.enumMappingType.MAPPING_TYPE_FUN
                    Return NODE_FUN
                Case modDeclares.enumMappingType.MAPPING_TYPE_JOIN
                    Return NODE_GEN
                Case modDeclares.enumMappingType.MAPPING_TYPE_LOOKUP
                    Return NODE_LOOKUP
                Case modDeclares.enumMappingType.MAPPING_TYPE_PROC
                    Return NODE_PROC
                Case modDeclares.enumMappingType.MAPPING_TYPE_VAR
                    Return NODE_VARIABLE
                Case modDeclares.enumMappingType.MAPPING_TYPE_WORKVAR
                    Return NODE_VARIABLE
                Case modDeclares.enumMappingType.MAPPING_TYPE_CONSTANT
                    Return NODE_CONST
                Case Else
                    Return NODE_STRUCT_FLD
            End Select

        Catch ex As Exception
            LogError(ex, "modGeneral GetNodeTypeFromSourceType")
            Return ""
        End Try

    End Function

    Function GetSourceTypeFromNodeType(ByVal ndType As String) As enumMappingType

        Try
            Select Case ndType
                Case NODE_STRUCT_FLD
                    Return modDeclares.enumMappingType.MAPPING_TYPE_FIELD
                Case NODE_FUN, NODE_TEMPLATE
                    Return modDeclares.enumMappingType.MAPPING_TYPE_FUN
                Case NODE_GEN
                    Return modDeclares.enumMappingType.MAPPING_TYPE_JOIN
                Case NODE_LOOKUP
                    Return modDeclares.enumMappingType.MAPPING_TYPE_LOOKUP
                Case NODE_PROC
                    Return modDeclares.enumMappingType.MAPPING_TYPE_PROC
                Case NODE_VARIABLE
                    Return modDeclares.enumMappingType.MAPPING_TYPE_VAR
                Case Else
                    Return modDeclares.enumMappingType.MAPPING_TYPE_NONE
            End Select

        Catch ex As Exception
            LogError(ex, "modGeneral GetSourceTypeFromNodeType")
            Return ""
        End Try

    End Function

    '/////////////////////////////////////////////////////////////////////////////////////////
    Function LoadFldArrFromTreeview(ByVal arrList As ArrayList, ByRef tv As TreeView, Optional ByVal OnlyAddCheckedItem As Boolean = True) As Boolean

        Try
            If tv.GetNodeCount(True) <= 0 Then Exit Function
            arrList.Clear()

            LoadFldArrFromTreeview = AddTVItemToArr(arrList, tv.Nodes, OnlyAddCheckedItem)

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral LoadFldArrFromTreeview")
            Return False
        End Try

    End Function

    Function AddTVItemToArr(ByRef arrList As ArrayList, ByVal cNode As TreeNodeCollection, Optional ByVal OnlyAddCheckedItem As Boolean = True) As Boolean

        Try
            Dim nd As TreeNode
            For Each nd In cNode
                If OnlyAddCheckedItem = True Then
                    If nd.Checked Then
                        arrList.Add(nd.Tag) '//Store tag object into arraylist
                    End If
                Else
                    arrList.Add(nd.Tag) '//Store tag object into arraylist
                End If
                '//Search all nodes under this node
                AddTVItemToArr(arrList, nd.Nodes, OnlyAddCheckedItem)
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral AddTVItemToArr")
            Return False
        End Try

    End Function

    '// Added By TKarasch March/April 2007 to add fields to any and all Treeviews in SQDStudio
    '// Fills TreeView OR TreeNode and their fields under selected item 
    '// selObj is structure, Structure Selection, or Datastore selection
    '// selNode is the treenode for the selobj that is to be filled - this is a parent node to fill under
    '// tv is the treeview to fill from selobj - this is an EMPTY tree to be filled with fields
    '// fromDS is true if from the Datastore form or user control so only selected fields are shown
    '// **********************
    '// >>>>>>> NOTE !!!!!!!!!
    '// ****************************************************************************************
    '// ******** Must have one OR the other of TV-treeview OR selnode-treenode BUT NOT BOTH ****
    '// ****************************************************************************************
    Public Function AddFieldsToTreeView(ByVal selObj As INode, Optional ByRef selnode As TreeNode = Nothing, Optional ByRef tv As TreeView = Nothing, Optional ByVal fromDS As Boolean = False) As Boolean

        '// First Initialize all Variables 
        '// There are many variables so that loop will function QUICKLY
        '// with NO recursion no matter how many fields there are.
        '// also because this function will handle treenodes, treeviews
        '// and any type to INode object that has a field arraylist
        '// on average less than 9 lines of code execute per loop cycle
        Try
            Dim InodeType As String = "" 'type of Inode passed in
            Dim Initfield As clsField = Nothing 'first field of the Inode's field arraylist
            Dim curField As clsField = Nothing 'field currently being looked at to add to the tree
            Dim FldCount As Integer = 0 'top end of arraylist count (i.e. number of fields -1)
            Dim firstFlag As Boolean = False 'true means fill tree at initial level until level goes up
            Dim ParentNode As TreeNode = Nothing 'either the node passed in to add to 
            'or firstnode added to empty tree
            Dim curNode As TreeNode = Nothing 'current node on the tree

            InodeType = selObj.Type
            Select Case InodeType
                Case NODE_STRUCT
                    If CType(selObj, clsStructure).ObjFields.Count = 0 Then
                        Exit Function
                    End If
                    Initfield = CType(selObj, clsStructure).ObjFields(0)
                    FldCount = CType(selObj, clsStructure).ObjFields.Count - 1
                    firstFlag = True
                Case NODE_STRUCT_SEL
                    If CType(selObj, clsStructureSelection).ObjSelectionFields.Count = 0 Then
                        Exit Function
                    End If
                    If fromDS = True Then
                        Initfield = CType(selObj, clsStructureSelection).ObjSelectionFields(0)
                        FldCount = CType(selObj, clsStructureSelection).ObjSelectionFields.Count - 1
                    Else
                        Initfield = CType(selObj, clsStructureSelection).ObjStructure.ObjFields(0)
                        FldCount = CType(selObj, clsStructureSelection).ObjStructure.ObjFields.Count - 1
                    End If
                    firstFlag = True
                Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                    If CType(selObj, clsDSSelection).DSSelectionFields.Count = 0 Then
                        Exit Function
                    End If
                    Initfield = CType(selObj, clsDSSelection).DSSelectionFields(0)
                    FldCount = CType(selObj, clsDSSelection).DSSelectionFields.Count - 1
                    If selnode Is Nothing Then
                        firstFlag = True
                    Else
                        ParentNode = selnode
                    End If
            End Select

            Dim InitLevel As Integer = Initfield.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL) 'level of 1st fld
            Dim i As Integer = 0 'loop counter
            Dim CurFieldLevel As Integer = InitLevel 'level of the current field to be added
            Dim LastNodeLevel As Integer = InitLevel 'level of the last field added
            '/// note *** these are set the same to syncronize the node level and the field level***
            Dim ParentCol As New Collection 'collection of the current parent of EVERY field level.
            ' this is very important. It changes dynamically as list of fields is put into the tree.
            ' it is always the most current parent tree node at each level based on the collection "key"
            ' Each node level is the collection "key" and the collection "object" is the parent node of 
            ' that level.

            '/// ***********************************************************************
            '/// ********************* Initialization Finished *************************
            '/// ***********************************************************************
            '/// finally all variables are initialized
            '/// Now start the loop through the fields
            For i = 0 To FldCount
                Select Case InodeType
                    Case NODE_STRUCT
                        curField = CType(selObj, clsStructure).ObjFields(i)
                    Case NODE_STRUCT_SEL
                        If fromDS = True Then
                            curField = CType(selObj, clsStructureSelection).ObjSelectionFields(i)
                        Else
                            curField = CType(selObj, clsStructureSelection).ObjStructure.ObjFields(i)
                        End If
                    Case NODE_SOURCEDSSEL, NODE_TARGETDSSEL
                        curField = CType(selObj, clsDSSelection).DSSelectionFields(i)
                End Select
                '// set the current field level from the field object
                CurFieldLevel = curField.GetFieldAttr(enumFieldAttributes.ATTR_LEVEL)
                '// if this level is less than the initial level then 
                '// the field level is invalid and it is skipped
                If CurFieldLevel >= InitLevel Then
                    If CurFieldLevel <> LastNodeLevel Then
                        If CurFieldLevel > LastNodeLevel Then
                            '// If the level goes Up on the next node, the last node (curNode) is it's 
                            '// parent(node) so, remove the parent node of the last level and replace 
                            '// it with the new parent, which is curNode 
                            If ParentCol.Contains(LastNodeLevel.ToString) Then
                                ParentCol.Remove(LastNodeLevel.ToString)
                            End If
                            ParentCol.Add(curNode, LastNodeLevel.ToString)
                            ParentNode = curNode
                            '// level went up so start adding nodes to the parent node instead of the treeview
                            firstFlag = False
                        Else
                            '// If the level goes down it may go down 1 or any number down
                            '// so we go to the arraylist of parents and assign the right parent
                            '// to add this node to, based on the current field level.
                            ParentNode = ParentCol.Item((CurFieldLevel - 1).ToString)
                        End If
                    End If

                    If firstFlag = True Then
                        '// add the first level of treenodes into the treeview
                        ParentNode = AddTreeNode(tv, curField.Type, curField, False, curField.FieldName)
                        '// first time through set the curNode
                        If curNode Is Nothing Then
                            curNode = ParentNode
                        End If
                        ParentNode.EnsureVisible()
                    Else
                        '// now the tree view has nodes OR there was an initial parent node sent in to
                        '// the function to add nodes to. Either way the "firstFlag" was set to false
                        curNode = AddNode(ParentNode, curField.Type, curField, False, curField.FieldName)
                        curNode.EnsureVisible()
                    End If
                    '// make the last node level, the current level for the next time through the loop
                    LastNodeLevel = CurFieldLevel
                Else
                    'MsgBox("The Field named >> " & curField.Text & ", has an invalid field level" & (Chr(13)) & "Please check the structure containing this field for validity.", MsgBoxStyle.Information, MsgTitle)
                    Log("The Field named >> " & curField.Text & ", has an invalid field level Please check the structure containing this field for validity.")
                End If
            Next

            Return True

        Catch ex As Exception
            LogError(ex, "modGeneral AddFieldsToTreeView")
            Return False
        End Try

    End Function

#End Region

#Region "FTP"

    Property FTPHostName() As String
        Get
            Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "FTPHostName")
        End Get
        Set(ByVal Value As String)
            Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "FTPHostName", Value)
        End Set
    End Property

    Property FTPPort() As String
        Get
            Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "FTPPort")
        End Get
        Set(ByVal Value As String)
            Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "FTPPort", Value)
        End Set
    End Property

    Property FTPUserID() As String
        Get
            Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "FTPUserID")
        End Get
        Set(ByVal Value As String)
            Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "FTPUserID", Value)
        End Set
    End Property

    Property FTPRemoteDir() As String
        Get
            Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "FTPRemoteDir")
        End Get
        Set(ByVal Value As String)
            Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "FTPRemoteDir", Value)
        End Set
    End Property

#End Region

#Region "Folder Paths"

    '//Creates a Tempfolder in the same location of Application exe file
    Function GetAppTemp(Optional ByVal TempFolderName As String = "Temp") As String

        Dim AppData As String = System.Windows.Forms.Application.LocalUserAppDataPath()
        Dim AppTemp As String = ""

        If Right(AppData, 1) <> "\" Then
            AppData = AppData & "\"
        End If

        AppTemp = AppData & TempFolderName

        If Right(AppTemp, 1) <> "\" Then
            AppTemp = AppTemp & "\"
        End If

        If System.IO.Directory.Exists(AppTemp) = False Then
            System.IO.Directory.CreateDirectory(AppTemp)
        End If

        GetAppTemp = AppTemp

    End Function

    '//Filename with extension
    Public Function GetFileNameFromPath(ByVal sPath As String) As String
        If sPath <> "" Then
            GetFileNameFromPath = sPath.Substring(sPath.LastIndexOf("\") + 1)
        Else
            GetFileNameFromPath = ""
        End If
    End Function

    Function GetFileNameOnly(ByVal fullPath As String) As String
        Dim st As String
        fullPath = Replace(fullPath, "/", "\")
        st = InStr(StrReverse(fullPath), "\")
        If st > 0 Then
            GetFileNameOnly = StrReverse(Mid(StrReverse(fullPath), 1, st - 1))
        Else
            GetFileNameOnly = fullPath
        End If
    End Function

    Function GetDirFromPath(ByVal filepath As String) As String
        Dim tempStr As String
        Dim idx As Integer

        idx = filepath.LastIndexOf("\")
        tempStr = Strings.Left(filepath, idx + 1)
        GetDirFromPath = tempStr

    End Function

    '//Filename without extension
    Public Function GetFileNameWithoutExtenstionFromPath(ByVal sPath As String) As String
        Dim sTemp1 As String
        sTemp1 = GetFileNameOnly(sPath)
        If sTemp1.Contains("%") = False Then
            If sTemp1.Contains(".") = True Then
                GetFileNameWithoutExtenstionFromPath = Left(sTemp1, sTemp1.IndexOf("."))
            Else
                GetFileNameWithoutExtenstionFromPath = sTemp1
            End If
        Else
            GetFileNameWithoutExtenstionFromPath = sPath
        End If
    End Function

    Public Function GetAppPath() As String
        GetAppPath = System.Windows.Forms.Application.StartupPath()
        If Right(GetAppPath, 1) <> "\" Then
            GetAppPath = GetAppPath & "\"
        End If
    End Function

    ' JDM - 0/28/2006 - Find parent folder
    Public Function GetParentPath() As String
        Dim strAppPath As String = GetAppPath()
        Dim astrParentPath() As String = strAppPath.Split("\")
        Dim strParentPath As String = ""
        For intLoop As Integer = 0 To astrParentPath.Length - 3
            strParentPath += astrParentPath.GetValue(intLoop) & "\"
        Next
        Return strParentPath
    End Function

    ' JDM - 0/28/2006 - Find Models folder
    Public Function GetModelPath() As String
        Return GetAppTemp() & "Models"
    End Function

    ' JDM - 0/28/2006 - Find Scripts folder
    Public Function GetScriptPath() As String
        Return GetAppTemp() & "Scripts"
    End Function

    ' JDM - 0/28/2006 - Find Help folder
    'Public Function GetHelpPath() As String
    '    Return GetParentPath() & "Help"
    'End Function

    '// new 12/4/2006 by TK
    '//Creates a Query folder in the Application directory
    Function QueryFolderPath(Optional ByVal QueryFolderName As String = "Queries") As String
        Dim AppPath As String = GetAppTemp()

        If System.IO.Directory.Exists(AppPath & QueryFolderName) = False Then
            System.IO.Directory.CreateDirectory(AppPath & QueryFolderName)
        End If
        QueryFolderPath = AppPath & QueryFolderName
    End Function

    Function ReportFolderPath(Optional ByVal ReportFolderName As String = "Temp") As String

        Dim AppPath As String = GetAppPath()
        Dim AppTemp As String = GetAppTemp()

        'If System.IO.Directory.Exists(AppPath & ReportFolderName) = False Then
        '    System.IO.Directory.CreateDirectory(AppPath & ReportFolderName)
        'End If
        'ReportFolderPath = AppPath & ReportFolderName
        ReportFolderPath = GetAppTemp()

    End Function

    '//Unix file names are case sensitive so get exact same name of file.
    Function GetCaseSensetivePath(ByVal sPath As String) As String
        Dim files() As String

        If IO.File.Exists(sPath) Then
            files = IO.Directory.GetFiles(IO.Path.GetDirectoryName(sPath), IO.Path.GetFileName(sPath))
            Return files(0) '//just read first name which will be exact same name we want
        Else
            Return sPath
        End If

    End Function

#End Region

#Region "Paths in Registry"

    '//new: 8/15/05
    'Property LocalDTDFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalDTDFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalDTDFolderPath", Value)
    '    End Set
    'End Property
    ''//new: 8/15/05

    'Property LocalDDLFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalDDLFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalDDLFolderPath", Value)
    '    End Set
    'End Property
    ''//new: 8/15/05
    'Property LocalCFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalCFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalCFolderPath", Value)
    '    End Set
    'End Property
    ''//new: 8/15/05
    'Property LocalCOBOLFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalCOBOLFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalCOBOLFolderPath", Value)
    '    End Set
    'End Property

    ''//new: 8/15/05
    'Property LocalProjectFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalProjectFolderPath", GetAppPath() & "PROJECTS")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalProjectFolderPath", Value)
    '    End Set
    'End Property

    ''//new: 8/15/05
    'Property LocalDBDFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalDBDFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalDBDFolderPath", Value)
    '    End Set
    'End Property

    'Property ScriptFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "ScriptFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "ScriptFolderPath", Value)
    '    End Set
    'End Property

    'Property ModelFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "ModelFolderPath", GetAppPath() & "Model")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "ModelFolderPath", Value)
    '    End Set
    'End Property

    'Property LocalSCRIPTFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalSCRIPTFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalSCRIPTFolderPath", Value)
    '    End Set
    'End Property

    'Property LocalModelFolderPath() As String
    '    Get
    '        Return Microsoft.VisualBasic.GetSetting("SQDStudio", "AppSettings", "LocalModelFolderPath", GetAppPath() & "Scripts")
    '    End Get
    '    Set(ByVal Value As String)
    '        Microsoft.VisualBasic.SaveSetting("SQDStudio", "AppSettings", "LocalModelFolderPath", Value)
    '    End Set
    'End Property

#End Region

#Region "Substitution Variables"

    Function GetVarNum(ByVal Instr As String) As String

        Try
            If SubstList.Contains(Instr) = False Then
                SubstList.Add(Instr)
            End If

            Return Instr

        Catch ex As Exception
            LogError(ex, "modGeneral GetVarNum")
            Return ""
        End Try

    End Function

    Function AddToSubstList(ByVal Instr As String) As String

        Try
            If SubstList.Contains(Instr) = False Then
                SubstList.Add(Instr)
            End If

            Return Instr

        Catch ex As Exception
            LogError(ex, "modGeneral AddToSubstList")
            Return -1
        End Try

    End Function

#End Region

End Module
