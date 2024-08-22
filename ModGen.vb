Option Explicit On

Module ModGen


    Public blnErrorLog As Boolean
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String

    Public strInputFolderPath As String
    Public gstrInputFilePMT As String
    Public gstrInputFileADV As String
    Public gstrInputFolder As String

    Public gstrOutputFile_Name As String

    Public strOutputFolderPath As String

    ''Res
    Public gstrResOutputfile As String
    Public gstrResponseInputFolder As String
    Public gstrResponseInputFile As String

    Public strResponseFolderPath As String             ' Response folder path
    Public strReverseResponseFolderPath As String            ' RevResponse folder path
    Public strRemoveRows As String            ' RemoveRows from Input
    ''Archive
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String
    ''''''''''''''''''
    Public strReportFolderPath As String
    Public strMasterFolderPath As String
    Public strProceed As String
    Public strInvalidTrans As String
    Public FileCounter As String

    Public strValidationPath As String
    Public strTransactionNo As String

    ''Encryption
    Public strEncryptFolderPath As String            ' Output folder path
    Public strEncrypt As String
    Public strBatchFilePath As String
    Public strPICKDIRpath As String
    Public strDROPDIRPath As String

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String

    Public strFileFormat As String
     
End Module


