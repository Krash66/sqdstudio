Imports System.IO
Imports System.Text.RegularExpressions

Public Enum FTPDType
    Unix
    MVS
    MVS_PDS
    Windows
End Enum

Public Enum FTPFileType
    File
    Directory
    Pipe
    Symlink
    Other
End Enum

Public Enum FTPFileEntryType
    UNKNOWN
    UNIX
    MVS_Dir
    MVS_Partitioned
    MVS_JCL
End Enum

Public Class FTPFile

    Public Filename As String
    Public Type As FTPFileType
    Public Owner As String
    Public Group As String
    Public Size As String
    Public ModDate As String
    Public Mode As String
    Public UnParsed As String
    Public EntryType As FTPFileEntryType

    Public Function Clone() As FTPFile
        Clone = New FTPFile
        Clone.Filename = Me.Filename
        Clone.Type = Me.Type
        Clone.Owner = Me.Owner
        Clone.Group = Me.Group
        Clone.Size = Me.Size
        Clone.ModDate = Me.ModDate
        Clone.Mode = Me.Mode
        Clone.UnParsed = Me.UnParsed
        Clone.EntryType = Me.EntryType
    End Function

    Public ReadOnly Property EntryString() As String

        Get
            Dim strEntryType As String = ""
            Select Case Me.EntryType
                Case FTPFileEntryType.MVS_Dir
                    strEntryType = "MVS_Dir|"
                Case FTPFileEntryType.MVS_JCL
                    strEntryType = "MVS_JCL|"
                Case FTPFileEntryType.MVS_Partitioned
                    strEntryType = "MVS_Partitioned|"
                Case FTPFileEntryType.UNIX
                    strEntryType = "UNIX|"
                Case FTPFileEntryType.UNKNOWN
                    strEntryType = ""
            End Select
            Return strEntryType & Me.UnParsed
        End Get

    End Property

    Public Function Fill(ByVal Listing As String, Optional ByVal ListingType As FTPFileEntryType = FTPFileEntryType.UNKNOWN) As Boolean

        If ListingType = FTPFileEntryType.UNKNOWN Then
            ListingType = Me.EntryType
        Else
            Me.EntryType = ListingType
        End If

        Select Case ListingType
            Case FTPFileEntryType.MVS_Dir
                Me.LIST_Parse_MVS_dir(Listing)
            Case FTPFileEntryType.MVS_JCL
                'Me.LIST_Parse_MVS_jcl(Listing)
            Case FTPFileEntryType.MVS_Partitioned
                Me.LIST_Parse_MVS_partitioned(Listing)
            Case FTPFileEntryType.UNIX
                Me.LIST_Parse_Unix(Listing)
        End Select

    End Function

    Public Function FillFromEntryString(ByVal EntryString As String) As Boolean

        Dim strEntryType As String
        Dim Entry As String

        strEntryType = Mid(EntryString, 1, InStr(EntryString, "|"))
        Entry = Mid(EntryString, InStr(EntryString, "|") + 1)

        Select Case strEntryType
            Case "MVS_Dir|"
                Me.Fill(Entry, FTPFileEntryType.MVS_Dir)
            Case "MVS_JCL|"
                Me.Fill(Entry, FTPFileEntryType.MVS_JCL)
            Case "MVS_Partitioned|"
                Me.Fill(Entry, FTPFileEntryType.MVS_Partitioned)
            Case "UNIX|"
                Me.Fill(Entry, FTPFileEntryType.UNIX)
        End Select

    End Function

    Private Sub LIST_Parse_Unix(ByVal line As String)

        Dim re As New Regex("([^ ]+) +[0-9]+ +(\w+) +(\w+) +([0-9]+) +((?:[a-zA-Z]{3} +)?[^ ]+ +[^ ]+) +(.*)")
        Dim matches As Match

        Me.EntryType = FTPFileEntryType.UNIX
        Me.UnParsed = line
        matches = re.Match(line)
        Me.Mode = matches.Groups(1).ToString

        Select Case Me.Mode.Chars(0)
            Case "-"
                Me.Type = FTPFileType.File
            Case "d"
                Me.Type = FTPFileType.Directory
            Case "l"
                Me.Type = FTPFileType.Symlink
            Case "p"
                Me.Type = FTPFileType.Pipe
            Case Else
                Me.Type = FTPFileType.Other
        End Select

        Me.Owner = matches.Groups(2).ToString
        Me.Group = matches.Groups(3).ToString
        Me.Size = matches.Groups(4).ToString
        Me.ModDate = matches.Groups(5).ToString
        Me.Filename = matches.Groups(6).ToString

        If Me.Type = FTPFileType.Symlink Then
            Me.Filename = Mid(Me.Filename, 1, InStr(Me.Filename, " -> ") - 1)
        End If

    End Sub

    Private Sub LIST_Parse_MVS_dir(ByVal line As String)

        Dim matches As Match

        Me.EntryType = FTPFileEntryType.MVS_Dir
        Me.UnParsed = line
        matches = Regex.Match(line, "Error determining attributes +([^ ]+)")
        If matches.Success Then
            Me.Type = FTPFileType.File
            Me.Filename = matches.Groups(1).ToString
        Else
            matches = Regex.Match(line, ".* +([^ ]+) +([^ ]+) *")
            If matches.Groups(1).ToString = "PO" Then
                Me.Type = FTPFileType.Directory
            Else
                Me.Type = FTPFileType.File
            End If
            Me.Filename = matches.Groups(2).ToString
        End If

    End Sub

    Private Sub LIST_Parse_MVS_partitioned(ByVal line As String)

        Dim matches As Match
        ' Name     VV.MM   Created       Changed      Size  Init   Mod   Id

        Me.EntryType = FTPFileEntryType.MVS_Partitioned
        Me.UnParsed = line

        '                             file    ?      create    mod            size     ??      ??      owner
        matches = Regex.Match(line, "([^ ]+)(?: +[^ ]+ +([^ ]+) +([^ ]+ +[^ ]+) +([^ ]+) +([^ ]+) +[^ ]+ +([^ ]+))?")

        Me.Filename = matches.Groups(1).ToString
        Me.Owner = matches.Groups(6).ToString
        Me.ModDate = matches.Groups(3).ToString
        Me.Size = matches.Groups(4).ToString

    End Sub

End Class

Public Class FTPClient

    Private CWD_Sync As Boolean
    Private Hostname As String
    Private Port As String
    Private Username As String
    Shared oldHostname As String
    Shared oldPort As String = "21"
    Shared oldUsername As String
    Shared olddir As String
    Private Password As String
    Private ScriptFile As String = "r.ftp"
    Private TempFile As String = "ftp.out"
    Public fScript As StreamWriter

    Public Files As Collection
    Public SystemType As FTPDType
    Public CurrentWorkingDirectory As String
    Public ftpForm As Windows.Forms.Form = frmFTPClient
    Public OutText As TextBox

    Property hostNameValue() As String
        Get
            Return oldHostname
        End Get
        Set(ByVal Value As String)
            oldHostname = Value
        End Set
    End Property

    Property userNameValue() As String
        Get
            Return oldUsername
        End Get
        Set(ByVal Value As String)
            oldUsername = Value
        End Set
    End Property

    Property portValue() As String
        Get
            Return oldPort
        End Get
        Set(ByVal Value As String)
            oldPort = Value
        End Set
    End Property

    Property dirValue() As String
        Get
            Return olddir
        End Get
        Set(ByVal Value As String)
            olddir = Value
        End Set
    End Property

    Public Sub New(Optional ByVal Hostname As String = "", Optional ByVal Port As String = "21", _
                   Optional ByVal Username As String = "", Optional ByVal Password As String = "", _
                   Optional ByVal InitialDir As String = "", Optional ByRef formftp As frmFTPClient = Nothing)

        Me.Hostname = Hostname
        Me.Port = Port
        Me.Username = Username
        Me.Password = Password
        If formftp IsNot Nothing Then
            Me.ftpForm = formftp
            OutText = formftp.txtFTPout
        End If

        If Me.Hostname.Length > 0 Then
            Me.CurrentWorkingDirectory = ""
            If InitialDir.Length > 0 Then
                Me.CWD(InitialDir)
            Else
                Me.CurrentWorkingDirectory = Me.PWD()
            End If
        End If

    End Sub

    Public Function PWD() As String

        If Me.CWD_Sync Then
            PWD = Me.CurrentWorkingDirectory
        Else
            Dim fScript As StreamWriter = Init_FTP_Script()
            Dim Reply As String
            Dim re As New Regex("""(.*)""")
            Dim matches As Match

            fScript.WriteLine("pwd")
            OutText.AppendText("pwd" & Chr(10))
            Exec_FTP_Script(fScript, Me.TempFile)
            Reply = Find_Reply_To("pwd", Me.TempFile)

            GetSystype(Reply)
            matches = re.Match(Reply)
            Me.CurrentWorkingDirectory = matches.Groups(1).ToString
            Me.CWD_Sync = True
            PWD = Me.CurrentWorkingDirectory
        End If

    End Function

    Public Sub InitFTPScript()

        fScript = Init_FTP_Script()

    End Sub

    Public Function AddGetCmd(ByVal RemotePath As String, ByVal LocalPath As String) As Boolean

        fScript.WriteLine("get """ & RemotePath & """ """ & LocalPath & """")
        OutText.AppendText("get """ & RemotePath & """ """ & LocalPath & """" & Chr(10))

    End Function

    Public Function AddPutCmd(ByVal LocalFile As String, ByVal RemoteFile As String) As Boolean

        fScript.WriteLine("put """ & LocalFile & """ """ & RemoteFile & """")
        OutText.AppendText("put """ & LocalFile & """ """ & RemoteFile & """" & Chr(10))

    End Function

    Public Function GETFILE() As Boolean

        Exec_FTP_Script(fScript, Me.TempFile)

    End Function

    Public Function PUTFILE() As Boolean

        Exec_FTP_Script(fScript, Me.TempFile)

    End Function

    Public Function CWD(ByVal NewPath As String) As String

        Me.CWD_Sync = False

        Dim fScript As StreamWriter = Init_FTP_Script()
        Dim Reply As String
        Dim re As New Regex("""(.*)""")
        Dim matches As Match

        If NewPath.Length > 0 Then
            fScript.WriteLine("cd """ & NewPath & """")
            OutText.AppendText("cd """ & NewPath & """" & Chr(10))
        End If
        fScript.WriteLine("pwd")
        OutText.AppendText("pwd" & Chr(10))
        Exec_FTP_Script(fScript, Me.TempFile)
        Reply = Find_Reply_To("pwd", Me.TempFile)

        GetSystype(Reply)
        matches = re.Match(Reply)
        Me.CurrentWorkingDirectory = matches.Groups(1).ToString
        Me.CWD_Sync = True

        CWD = Me.CurrentWorkingDirectory

    End Function

    Public Function UP() As String

        UP = Me.CWD("..")

    End Function

    Public Function CDLS(ByVal NewPath As String) As Collection

        Dim fScript As StreamWriter = Init_FTP_Script()
        Dim Listing As Collection
        Dim line As String
        Dim file As New FTPFile

        Dim Reply As String
        Dim matches As Match

        If NewPath.Length > 0 Then
            fScript.WriteLine("cd """ & NewPath & """")
            OutText.AppendText("cd """ & NewPath & """" & Chr(10))
        End If
        fScript.WriteLine("pwd")
        OutText.AppendText("pwd" & Chr(10))
        fScript.WriteLine("dir")
        OutText.AppendText("dir" & Chr(10))

        Exec_FTP_Script(fScript, Me.TempFile)

        ' New Dir
        Reply = Find_Reply_To("pwd", Me.TempFile)

        GetSystype(Reply)

        matches = Regex.Match(Reply, """(.*)""")
        Me.CurrentWorkingDirectory = matches.Groups(1).ToString
        Me.CWD_Sync = True

        ' LIST
        Listing = LIST_seperate(Me.TempFile)

        CDLS = Nothing
        If (Listing.Count > 0) Then
            file.EntryType = LIST_type(Listing(1))
            Listing.Remove(1)

            CDLS = New Collection
            For Each line In Listing
                file.Fill(line)
                CDLS.Add(file.Clone)
            Next
        End If
        Me.Files = CDLS

    End Function

    Public Function LIST() As Collection

        Dim fScript As StreamWriter = Init_FTP_Script()
        Dim Listing As Collection
        Dim line As String
        Dim file As New FTPFile

        fScript.WriteLine("dir")
        OutText.AppendText("dir" & Chr(13))
        Exec_FTP_Script(fScript, Me.TempFile)

        Listing = LIST_seperate(Me.TempFile)

        LIST = Nothing
        If (Listing.Count > 0) Then
            file.EntryType = LIST_type(Listing(1))
            Listing.Remove(1)

            LIST = New Collection
            For Each line In Listing
                file.Fill(line)
                LIST.Add(file.Clone)
            Next
        End If
        Me.Files = LIST

    End Function

    Private Function LIST_type(ByVal Headers As String) As FTPFileEntryType

        Select Case Headers.Trim
            Case "Volume Unit    Referred Ext Used Recfm Lrecl BlkSz Dsorg Dsname"
                LIST_type = FTPFileEntryType.MVS_Dir
                SystemType = FTPDType.MVS
            Case "Name     VV.MM   Created       Changed      Size  Init   Mod   Id"
                LIST_type = FTPFileEntryType.MVS_Partitioned
                SystemType = FTPDType.MVS_PDS
            Case Else
                LIST_type = FTPFileEntryType.UNIX
                SystemType = FTPDType.Unix

                'If Headers.StartsWith("total ") Then
                '   LIST_type = FTPFileEntryType.UNIX
                '   SystemType = FTPDType.Unix
                'Else
                '    LIST_type = Nothing
                '    SystemType = FTPDType.Windows
                'End If
        End Select

    End Function

    Private Sub GetSystype(ByVal Headers As String)

        Dim id As String

        id = Strings.Mid(Headers.Trim, 6, 1)
        If (id = "'" And Strings.InStr(Headers, "partitioned data set") > 0) Then
            SystemType = FTPDType.MVS_PDS
        ElseIf id = "'" Then
            SystemType = FTPDType.MVS
        ElseIf id = "/" Then
            SystemType = FTPDType.Unix
        Else
            SystemType = FTPDType.Windows
        End If

    End Sub

    Private Function LIST_seperate(ByVal LogFileName As String) As Collection

        Dim fLog As New StreamReader(LogFileName)
        Dim listing As New Collection
        Dim line As String
        Dim i As Integer = 0

        Find_Reply_To("dir", fLog)
        line = StreamReader_Readline_SkipBlank(fLog) 'Eat PORT reply
        line = StreamReader_Readline_SkipBlank(fLog) 'LIST command status reply

        If line.Substring(0, 4) = "125 " Or line.Substring(0, 4) = "150 " Then
            line = StreamReader_Readline_SkipBlank(fLog)

            While line <> ""
                If line.Length > 0 Then
                    i = i + 1
                    listing.Add(line, i)
                End If
                line = StreamReader_Readline_SkipBlank(fLog)
            End While

        End If
        fLog.Close()

        If i > 0 Then
            'Remove in reverse until line begins with ftp:
            line = listing(i)
            While Not i < 1 And Not line.Substring(0, 5) = "ftp: "
                listing.Remove(i)
                i = i - 1
                If i > 0 Then
                    line = listing(i)
                End If
            End While

            'Remove the "End of LIST" message
            If i > 0 Then
                listing.Remove(i)
                If (i - 1 > 0) Then
                    listing.Remove(i - 1)
                End If
            End If
        End If

        LIST_seperate = New Collection
        i = 0
        For Each line In listing
            i = i + 1
            LIST_seperate.Add(line, i)
        Next

    End Function

    Private Function Init_FTP_Script() As StreamWriter

        Dim fOut As New StreamWriter(Me.ScriptFile)
        fOut.WriteLine("open " & Hostname & " " & Port)
        OutText.AppendText("open " & Hostname & " " & Port & Chr(10))
        fOut.WriteLine(Username)
        OutText.AppendText(Username & Chr(10))
        fOut.WriteLine(Password)
        'OutText.AppendText(Password & Chr(10))
        fOut.WriteLine("ascii")
        OutText.AppendText("ascii" & Chr(10))
        If Me.CurrentWorkingDirectory.Length > 0 Then
            fOut.WriteLine("cd """ & Me.CurrentWorkingDirectory & """")
            OutText.AppendText("cd """ & Me.CurrentWorkingDirectory & """" & Chr(10))
        End If
        Init_FTP_Script = fOut

    End Function

    Private Sub Exec_FTP_Script(ByRef fScript As StreamWriter, ByVal LogFileName As String)

        fScript.WriteLine("close")
        OutText.AppendText("close" & Chr(10))
        fScript.WriteLine("quit")
        OutText.AppendText("quit" & Chr(10))
        fScript.Close()
        If LogFileName.Length > 0 Then
            LogFileName = " > " & """" & LogFileName & """"
        End If
        Debug.WriteLine("ftp -s:""" & Me.ScriptFile & """" & LogFileName)
        Shell(Environ$("comspec") & " /c" & "ftp -s:""" & Me.ScriptFile & """" & LogFileName, AppWinStyle.Hide, True, -1)

    End Sub

    Private Function Find_Reply_To(ByVal Command As String, ByVal LogFileName As String) As String

        Dim fLog As New StreamReader(LogFileName)
        Dim found As Boolean = False

        found = Find_Reply_To(Command, fLog)

        If found Then
            Find_Reply_To = StreamReader_Readline_SkipBlank(fLog)
            OutText.AppendText(Find_Reply_To & Chr(10))
        Else
            Find_Reply_To = Nothing
        End If
        fLog.Close()

    End Function

    Private Function Find_Reply_To(ByVal Command As String, ByRef fLog As StreamReader) As Boolean

        Dim line As String
        Dim found As Boolean = False

        line = StreamReader_Readline_SkipBlank(fLog)
        While Not found And line IsNot Nothing
            If line.Length = "ftp> ".Length + Command.Length Then
                If line.Substring("ftp> ".Length, Command.Length) = Command Then
                    found = True
                End If
            End If

            If Not found Then
                line = StreamReader_Readline_SkipBlank(fLog)
                OutText.AppendText(line & Chr(10))
            End If
        End While

        If found Then
            Find_Reply_To = True
        Else
            Find_Reply_To = False
        End If

    End Function

    Private Function StreamReader_Readline_SkipBlank(ByVal Stream As StreamReader) As String

        Dim found As Boolean = False
        Dim Line As String

        Line = Stream.ReadLine()
        While Not found And Not Line Is Nothing
            If Line.Length = 0 Then
                Line = Stream.ReadLine()
            Else
                found = True
            End If
        End While

        If found Then
            StreamReader_Readline_SkipBlank = Line
        Else
            StreamReader_Readline_SkipBlank = Nothing
        End If

    End Function

    Public Sub DelWorkFiles()

        System.IO.File.Delete(Me.TempFile)
        System.IO.File.Delete(Me.ScriptFile)

    End Sub

    Public Function TransRc() As Integer

        Dim fLog As New StreamReader(Me.TempFile)
        Dim line As String

        TransRc = 0
        line = StreamReader_Readline_SkipBlank(fLog)
        While Not line <> ""
            If line.Length > 0 Then
                If (Strings.Left(line, 20) = "451 Transfer aborted") Then
                    TransRc = 8
                    fLog.Close()
                    Exit Function
                End If
            End If
            line = StreamReader_Readline_SkipBlank(fLog)
        End While

        fLog.Close()

    End Function

    Public Sub ShowLog()

        Shell("notepad.exe " & Me.TempFile, AppWinStyle.NormalFocus)

    End Sub

End Class