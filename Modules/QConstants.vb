Public Module QConstants

    Public Class SQD_IQRY_ERR
        Inherits System.ApplicationException
        Public Sub New()
            MyBase.New()
            MyBase.Source = "IQRY"
        End Sub
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
            MyBase.Source = "IQRY"
        End Sub
        Public Sub New(ByVal Source As String, ByVal Message As String)
            MyBase.New(Message)
            MyBase.Source = "IQRY." & Source
        End Sub
    End Class

    Public Class SQD_IQRY_EOF
        Inherits QConstants.SQD_IQRY_ERR
        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
        Public Sub New(ByVal Source As String, ByVal Message As String)
            MyBase.New(Source, Message)
        End Sub
    End Class

    Public Class SQD_IQRY_INVALID
        Inherits QConstants.SQD_IQRY_ERR
        Public Sub New()
            MyBase.New()
        End Sub
        Public Sub New(ByVal Message As String)
            MyBase.New(Message)
        End Sub
        Public Sub New(ByVal Source As String, ByVal Message As String)
            MyBase.New(Source, Message)
        End Sub
    End Class

    'QNulTerminatedString
    Public Function QNTS(ByVal Bytes() As Byte) As String
        Dim Last As Integer
        Dim Nul As Byte
        Nul = 0
        Last = Array.IndexOf(Bytes, Nul)
        Return System.Text.Encoding.ASCII.GetString(Bytes, 0, Last)
    End Function

    Private Enum SQD_DS_Types
        DSBIN = 1
        DSTEXT = 2
        DSDELIM = 3
        DSXML = 4
        DSDB2L = 5
        DSVK = 6
        DSVR = 7
        DSVS = 8
        DSODBC = 9
        DSDYN = 10
        DSIMS = 11
        DSTVLEN = 12
        DSBVLEN = 13
        DSIMSCDC = 14
        DSVSAMCDC = 15
        DSXMLCDC = 16
        DSEDI86702 = 17
        DSIMSUNLD = 18
        DSDB2CDC = 19
        DSSTREAM = 20
        DSORA = 21
        DSHSSUNLD = 22
        DSVSAM = 23
        DSIBMEVENT = 24
        DSIBMEVENTBC = 25
        DSSQDCDCODBC = 26
        DSSQDCDCORA = 27
        DSSQDCDCDB2 = 28
        DSSQDCDCMQS = 29
        DSSQDCDC = 30
        DSSQLCDC = 31
    End Enum

    Public Enum QDS_Types
        BINARY = SQD_DS_Types.DSBIN
        DELIMITED = SQD_DS_Types.DSDELIM
        TEXT = SQD_DS_Types.DSTEXT
        UNKNOWN = 0
        XML = SQD_DS_Types.DSXML
        RELATIONAL = 5
        DB2LOAD = SQD_DS_Types.DSDB2L
        HSSUNLOAD = SQD_DS_Types.DSHSSUNLD
        IMSDB = SQD_DS_Types.DSIMS
        VSAM = SQD_DS_Types.DSVSAM
        IMSCDC = SQD_DS_Types.DSIMSCDC
        DB2CDC = SQD_DS_Types.DSDB2CDC
        VSAMCDC = SQD_DS_Types.DSVSAMCDC
        XMLCDC = SQD_DS_Types.DSXMLCDC
        IBMEVENT = SQD_DS_Types.DSIBMEVENT
        TRBCDC = 15
    End Enum

    Public Enum QFLD_Types
        FCHAR = 1
        TEXTNUM = 2
        SQD_DATE = 701
        YEAR = 711
        MONTH = 712
        DAY = 713
        WEEKDAY = 714
        TIME = 702
        TIMESTAMP = 703
        HOUR = 721
        MINUTE = 722
        SECOND = 723
        LOGICAL = 724
        VCHAR2 = 725
        SQDCLOB = 726
        PACKED = 3
        FULLW = 4
        HALFW = 5
        NUMBER = 6
        ZONE = 7
        DATETIME = 9
        DATE_TIME = 11
        VCHAR = 12
        VAR = 14
        BINARY = 15
        CDCBIN = 16
        MMDDYY = 20
        DDMMYY = 21
        YYDDMM = 22
        YYMMDD = 23
        YYYYDDMM = 24
        YYYYMMDD = 25
        MMDDYYYY = 26
        CIYYMMDD = 27
        ALPHA = 28
        ALNUM = 29
        HHMMSS = 30
        XMLCDATA = 96
        XMLPCDATA = 97
        XMLATTC = 98
        XMLATTPC = 99
        GROUPITEM = 100
        STCKTYPE = 800
        AUTONUM = 1000
    End Enum

    Private Enum SQD_DS_Transports
        NETFILE = 0
        NETTCP = 1
        NETMQS = 2
        NETFTP = 3
    End Enum

    Public Enum QDS_Transports
        FILE = SQD_DS_Transports.NETFILE
        TCP = SQD_DS_Transports.NETTCP
        MQS = SQD_DS_Transports.NETMQS
        FTP = SQD_DS_Transports.NETFTP
    End Enum

    Private Enum SQD_DS_Encodings
        COTH = 0
        CASC = 1
        CEBC = 2
    End Enum

    Public Enum QDS_Encodings
        LENDIAN_ASCII = SQD_DS_Encodings.COTH
        BENDIAN_ASCII = SQD_DS_Encodings.CASC
        BENDIAN_EBCDIC = SQD_DS_Encodings.CEBC
    End Enum

    Public Enum QRC
        RC_OK = 0
        RC_INVALID = 3
        RC_WARN = 4
        RC_ERR = 8
        RC_EOF = 100
    End Enum

End Module
