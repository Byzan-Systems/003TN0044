Imports System
Imports System.IO

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 

    Private DtValidation As DataTable       ''need to be dispose
    Private DtMaster As DataTable       ''need to be dispose
    Private dtValidation_Opt As DataTable
    Private DtSpCharValidation As DataTable       ''need to be dispose
    ''---
    Public DtMasterHouseBank As DataTable       ''need to be dispose
    Public DtInput_PMT As DataTable             ''need to be dispose

    Public DtUnSucInput As DataTable        ''need to be dispose
    Public DtInput_ADV As DataTable             ''need to be dispose
    'Public DtSucc_Output As DataTable                     ''need to be dispose
    Public DtUnSucc_Output As DataTable                ''need to be dispose
    Public Dt_OutputPMT As DataTable                     ''need to be dispose
    Public Dt_OutputADV As DataTable                     ''need to be dispose
    Public DtSuccResp As DataTable                     ''need to be dispose
    Public DtUnSucResp As DataTable                ''need to be dispose
    Public DtRespHeader As DataTable                     ''need to be dispose

    Private DtTemp_PMT As DataTable             ''need to be dispose
    Private DtTemp_ADV As DataTable             ''need to be dispose
    Private StrFilePath As String
    Private SpCharValidationPath As String
    ''---

    Public StrSettingPath As String

    Public ErrorMessage As String
    Private ValidationPath As String

    Private DtStatus As DataTable       ''need to be dispose
    Public HeaderUploadDate As String       ''need to be dispose

    Private strFileNameColNo As String = ""
    Private strFileNo As String = ""
    Private strTxnColNo As String = ""
    Private strLineNoColNo As String = ""
    Private strReasonColNo As String = ""


    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String)

        StrFilePath = _strFilePath
        StrSettingPath = _SettINIPath

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)

            ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)

            DtInput_PMT = New DataTable("DT_OutputPMT")
            DtMasterHouseBank = New DataTable("DtMasterHouseBank")
            DefineColumnForOutput_PMT(DtInput_PMT)
            DtInput_ADV = New DataTable("Dt_OutputADV")
            DefineColumnForOutput_ADV(DtInput_ADV)
            DtUnSucInput = New DataTable("DtUnSucInput")
            DefineColumnForOutput_PMT(DtUnSucInput)
            DtUnSucInput.Columns.Add("Reason")

            DtUnSucc_Output = New DataTable("DtUnSucc_Output")
            DefineColumnForOutput_PMT(DtUnSucc_Output)

            DtSuccResp = New DataTable("DtSuccResp")
            DefineColumnForRevResponse(DtSuccResp)
            DtUnSucResp = New DataTable("DtUnSucResp")
            DefineColumnForRevResponse(DtUnSucResp)

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")
        End Try

    End Sub
    Private Sub DefineColumnForOutput_PMT(ByRef DtInput As DataTable)
        Try
            'DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(strValidationPath, "Sheet1")
            'RemoveBlankRow(DtValidation)
            'For Each drRBI As DataRow In DtValidation.Select()
            '    DtInput.Columns.Add(New DataColumn(drRBI(1).ToString()))
            'Next
            DtInput.Columns.Add(New DataColumn("First Payt Doc")) ''0
            DtInput.Columns.Add(New DataColumn("Run Date")) ''1
            DtInput.Columns.Add(New DataColumn("Additional identification"))   ''2
            DtInput.Columns.Add(New DataColumn("Paying company code"))    ''3
            DtInput.Columns.Add(New DataColumn("Payment document no.")) ''4
            DtInput.Columns.Add(New DataColumn("Amount"))  ''5
            DtInput.Columns.Add(New DataColumn("Currency"))   ''6
            DtInput.Columns.Add(New DataColumn("Payment method"))   ''7
            DtInput.Columns.Add(New DataColumn("Vendor Account No/Code"))   ''8
            DtInput.Columns.Add(New DataColumn("Title of the Vendor"))   ''9
            DtInput.Columns.Add(New DataColumn("Name of the payee1"))    ''10
            DtInput.Columns.Add(New DataColumn("Name of the payee2"))    ''11
            DtInput.Columns.Add(New DataColumn("Name of the payee3"))    ''12
            DtInput.Columns.Add(New DataColumn("Postal code"))    ''13
            DtInput.Columns.Add(New DataColumn("City")) ''14
            DtInput.Columns.Add(New DataColumn("Street/house number"))  ''15
            DtInput.Columns.Add(New DataColumn("HouseBankID"))    ''16
            DtInput.Columns.Add(New DataColumn("AccountID")) ''17
            DtInput.Columns.Add(New DataColumn("Posting Date"))   ''18
            DtInput.Columns.Add(New DataColumn("System Date"))   '19
            DtInput.Columns.Add(New DataColumn("Delivery Mode"))   ''20
            DtInput.Columns.Add(New DataColumn("Cheque No"))   ''21
            DtInput.Columns.Add(New DataColumn("DD location"))   ''22
            DtInput.Columns.Add(New DataColumn("Beneficiary Account Number"))    ''23
            DtInput.Columns.Add(New DataColumn("IFSC Code"))    ''24
            DtInput.Columns.Add(New DataColumn("Beneficiary Bank Name"))    ''25
            DtInput.Columns.Add(New DataColumn("Beneficiary email id"))   ''26
            DtInput.Columns.Add(New DataColumn("Column1"))   ''27
            DtInput.Columns.Add(New DataColumn("Column2"))    ''28
            DtInput.Columns.Add(New DataColumn("Column3"))    ''29
            DtInput.Columns.Add(New DataColumn("Column4"))    ''30
            DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))   '31
            DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   '32
            'DtInput.Columns.Add(New DataColumn("SUBTXN_NO", System.Type.GetType("System.Int32")))   '33
            DtInput.Columns.Add(New DataColumn("REASON"))   '33

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "DefineColumnForOutput")
        End Try

    End Sub
    Private Sub DefineColumnForOutput_ADV(ByRef DtInput As DataTable)
        Try

            DtInput.Columns.Add(New DataColumn("Client Code"))   ''1
            DtInput.Columns.Add(New DataColumn("Payment Document No."))    ''2
            DtInput.Columns.Add(New DataColumn("Document Date")) ''3
            DtInput.Columns.Add(New DataColumn("~D"))  ''4
            DtInput.Columns.Add(New DataColumn("Accounting document number"))   ''5
            DtInput.Columns.Add(New DataColumn("document Posting date"))   ''6
            DtInput.Columns.Add(New DataColumn("Reference document number"))   ''7
            DtInput.Columns.Add(New DataColumn("Document date in document"))   ''8
            DtInput.Columns.Add(New DataColumn("Amount"))    ''9
            DtInput.Columns.Add(New DataColumn("Item text"))    ''10
            DtInput.Columns.Add(New DataColumn("Assignment number"))    ''11
            DtInput.Columns.Add(New DataColumn("Column1"))    ''12
            DtInput.Columns.Add(New DataColumn("Column2"))    ''13
            DtInput.Columns.Add(New DataColumn("Column3"))    ''14
            DtInput.Columns.Add(New DataColumn("Column4"))    ''15
            DtInput.Columns.Add(New DataColumn("TXN_NO", System.Type.GetType("System.Int32")))   '16
            DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   '17
            'DtInput.Columns.Add(New DataColumn("SUBTXN_NO", System.Type.GetType("System.Int32")))   '18
            DtInput.Columns.Add(New DataColumn("REASON"))   '19

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "DefineColumnForOutput")
        End Try

    End Sub

    Private Sub DefineColumnForRevResponse(ByRef DtInput As DataTable)
        Try
            'DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(strValidationPath, "Response")
            'RemoveBlankRow(DtValidation)
            'For Each drRBI As DataRow In DtValidation.Select()
            '    DtInput.Columns.Add(New DataColumn(drRBI(1).ToString()))
            'Next
            DtInput.Columns.Add(New DataColumn("PROGRAM RUN DATE")) ''0
            DtInput.Columns.Add(New DataColumn("RUN IDENTIFICATION")) ''1
            DtInput.Columns.Add(New DataColumn("Vendor Code"))   ''2
            DtInput.Columns.Add(New DataColumn("Voucher No."))    ''3
            DtInput.Columns.Add(New DataColumn("Value Date")) ''4
            DtInput.Columns.Add(New DataColumn("Bank Ref No"))  ''5
            DtInput.Columns.Add(New DataColumn("CHEQUE NUMBER"))   ''6
            DtInput.Columns.Add(New DataColumn("DD NUMBER"))   ''7
            DtInput.Columns.Add(New DataColumn("Amount"))   ''8
            DtInput.Columns.Add(New DataColumn("Transaction Status"))   ''9
            DtInput.Columns.Add(New DataColumn("UTR /NEFT Ref No"))    ''10
            DtInput.Columns.Add(New DataColumn("Reject Reason"))    ''11

            DtInput.Columns.Add(New DataColumn("TXN_NO"))    '12
            DtInput.Columns.Add(New DataColumn("File_NO", System.Type.GetType("System.Int32")))    '13
            DtInput.Columns.Add(New DataColumn("Line_NO", System.Type.GetType("System.Int32")))   ''14
            DtInput.Columns.Add(New DataColumn("Reason"))    '15

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "DefineColumnForRevResponse")
        End Try

    End Sub

    Private Sub FixedColPostion(ByVal dt_ValTable As DataTable)
        Try

            strFileNameColNo = ""
            strTxnColNo = ""
            strFileNo = ""
            strLineNoColNo = ""
            strReasonColNo = ""

            strFileNameColNo = dt_ValTable.Columns.Count - 5
            strTxnColNo = strFileNameColNo + 1
            strFileNo = strFileNameColNo + 2
            strLineNoColNo = strFileNameColNo + 3
            strReasonColNo = strFileNameColNo + 4
        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "FixedColPostion")
        End Try

    End Sub

    Public Function CheckValidateFile(ByVal gstrInputFile As String) As Boolean

        Try
            If Not File.Exists(gstrInputFile) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input file path is incorrect or not file found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If

            'If Not File.Exists(strValidationPath) Then
            '    Call ObjBaseClass.Handle_Error(New ApplicationException("Mapping Validation File path is incorrect or File not found. [" & strValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
            '    CheckValidateFile = False
            '    Exit Function
            'End If

            'If File.Exists(ValidationPath) Then
            CheckValidateFile = Validate()
            'Else
            'Call ObjBaseClass.Handle_Error(New ApplicationException("Validation File path is incorrect. [" & ValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
            'End If


        Catch ex As Exception
            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function


    Public Function CheckResponseValidateFile(ByVal strResFileName As String) As Boolean

        Try

            CheckResponseValidateFile = Validate_response(strResFileName)


        Catch ex As Exception
            CheckResponseValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckResponseValidateFile")
        End Try

    End Function
    Private Function Validate() As Boolean

        Validate = False

        Try
            DtMaster = ObjBaseClass.GetDataTable_ExcelSheet(strMasterFolderPath & "\" & "Master.xlsx", "Sheet1")

            'DrValidOptColumn = DtValidation.Select("[SRNO] <> 0  ", "[SRNO]")
            'RemoveBlankRow(DtValidation)
            Dim strFileName = Path.GetFileNameWithoutExtension(gstrInputFolder & "\" & gstrInputFilePMT)

            DtInput_PMT = ObjBaseClass.GetDatatable_PMT(gstrInputFolder & "\" & strFileName & ".pmt", DtInput_PMT, DtUnSucInput)
            If File.Exists(gstrInputFolder & "\" & strFileName & ".adv") Then
                DtInput_ADV = ObjBaseClass.GetDatatable_ADV(gstrInputFolder & "\" & strFileName & ".adv", DtInput_ADV, DtUnSucInput)
            End If

            RemoveBlankRow(DtInput_PMT)
            RemoveBlankRow(DtInput_ADV)

            Dim strquery = " ([HouseBankID] = '" & DtInput_PMT.Rows(0)(16).ToString().Trim & "')  "
            strquery = strquery & " and ( [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(17)).ToString().Trim & "' ) "



            'Dim strquery = " ([HouseBankID] = '" & DtInput_PMT.Rows(0)(11).ToString().Trim & "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(12).ToString().Trim &
            '     "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(13).ToString().Trim & "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(14).ToString().Trim & "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(15).ToString().Trim &
            '     "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(16).ToString().Trim & "' or [HouseBankID] = '" & DtInput_PMT.Rows(0)(17).ToString().Trim & "' )  "

            'strquery = strquery & " and ( [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(11)).ToString().Trim & "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(12)).ToString().Trim &
            '     "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(13)).ToString().Trim & "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(14)).ToString().Trim & "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(15)).ToString().Trim &
            '     "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(16)).ToString().Trim & "' or [AccountID] = '" & CStr(DtInput_PMT.Rows(0)(17)).ToString().Trim & "' ) "

            Dim rows = DtMaster.Select(strquery)

            If (rows.Count = 0) Then
                Validate = False
                DtUnSucInput.Rows.Add()
                DtUnSucInput.Rows(0)("Reason") = "House Bank ID and Account ID is not maintained in Master."
            Else
                DtMasterHouseBank = rows.CopyToDataTable()
                Validate = True
            End If


        Catch ex As Exception
            Validate = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")
        Finally
            ObjBaseClass.ObjectDispose(DtValidation)
            ObjBaseClass.ObjectDispose(DtMaster)
            ObjBaseClass.ObjectDispose(dtValidation_Opt)
        End Try

    End Function

    Private Function Validate_response(ByVal strRespFileName As String) As Boolean
        Try
            'DtValidation = ObjBaseClass.GetDataTable_ExcelSheet(strValidationPath, "Sheet1")
            Dim DtMapping = ObjBaseClass.GetDataTable_ExcelSheet(strMasterFolderPath & "\" & "Mapping.xlsx", "Sheet1")
            'DrValidOptColumn = DtValidation.Select("[SRNO] <> 0  ", "[SRNO]")
            RemoveBlankRow(DtMapping)
            Dim strFileName = Path.GetFileNameWithoutExtension(strRespFileName)
            Dim filenameNew = strFileName & ".txt"
            'File.Move(strRespFileName, strResponseFolderPath & "\" & filenameNew)
            'DtSuccResp = ObjBaseClass.GetDataTable_ExcelNumber(strRespFileName, 1)
            'For i = 0 To DtSuccResp.Rows.Count - 1
            '    If (Not DtSuccResp.Rows(i)(0) = "Company Code") Then
            '        DtSuccResp.Rows.RemoveAt(0)
            '    Else
            '        Exit For
            '    End If
            'Next

            DtSuccResp = ObjBaseClass.GetDatatable_Text(strRespFileName, DtSuccResp)
            RemoveBlankRow(DtSuccResp)
            For i = 0 To DtMapping.Rows.Count - 1
                Dim clMnm = DtMapping.Rows(i)(0).ToString
                'Dim row = DtMapping.Select("[Field]='" & clMnm.ToString & "'", "[OutputPosition]")
                Dim nR = Convert.ToInt16(DtMapping.Rows(i)(2)) - 1
                DtSuccResp.Columns(clMnm).SetOrdinal(nR)

            Next
            Validate_response = True
        Catch ex As Exception
            'Validate_response = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Validate")
        Finally
            ObjBaseClass.ObjectDispose(DtValidation)
            ObjBaseClass.ObjectDispose(DtMaster)
            ObjBaseClass.ObjectDispose(dtValidation_Opt)
        End Try

    End Function

    Private Function GetInArrayByComma(ByVal pStrValue As String) As String()

        Try

            Dim Tmpstr As String = ""
            Dim Index_S, Index_E, TmpIndex As Integer


            Index_E = InStr(pStrValue, Chr(34))

            If Index_E > 0 Then

                Index_S = 0
                Tmpstr = ""
                While True

                    Index_E = InStr(Index_S + 1, pStrValue, Chr(34))

                    If Index_E > 0 Then

                        Tmpstr += pStrValue.Substring(Index_S, Index_E - Index_S - 1).Replace(",", "|")
                        Index_S = Index_E
                        Index_E = InStr(Index_E + 1, pStrValue, Chr(34))
                        Tmpstr += pStrValue.Substring(Index_S, (Index_E - Index_S) - 1)
                        Index_S = Index_E

                    Else
                        Tmpstr += pStrValue.Substring(Index_S, pStrValue.Length - Index_S).Replace(",", "|")
                        GetInArrayByComma = Tmpstr.Split("|")
                        Exit While
                    End If

                End While

            Else
                GetInArrayByComma = pStrValue.Split(",")

            End If

        Catch ex As Exception

        End Try

    End Function

    Public Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean
        Dim delIndexStr As String = ""
        Dim DelIndex() As String
        Try

            For i As Integer = 0 To _DtTemp.Rows.Count - 1
                blnRowBlank = True
                Dim vRow As DataRow = _DtTemp.Rows(i)
                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    'DtTemp1.Rows(i).Delete()
                    delIndexStr = delIndexStr & i & ","
                End If

            Next

            If delIndexStr <> "" Then
                delIndexStr = Left(delIndexStr, delIndexStr.Length - 1)
                DelIndex = delIndexStr.Split(",")
                For j As Integer = 0 To DelIndex.Length - 1

                    If DelIndex(j) <> "" Then

                        If j = 0 Then
                            _DtTemp.Rows(DelIndex(j)).Delete()
                        Else
                            _DtTemp.Rows(DelIndex(j) - j).Delete()
                        End If
                        _DtTemp.AcceptChanges()

                    End If
                Next


            End If

            '------------------End Here

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try

    End Function

    Private Sub ClearArray(ByRef pArr() As String)
        Try
            For I As Int16 = 0 To pArr.Length - 1
                pArr(I) = ""
            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Sub ClearArraySplit(ByRef pArr() As String, ByVal inputPos As Integer)
        Try

            For I As Int16 = 0 To pArr.Length - 1
                If inputPos <> 0 And inputPos <> 10 And inputPos <> 11 Then
                    pArr(I) = ""
                End If

            Next

        Catch ex As Exception

        End Try

    End Sub

    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~Error~"
                Else
                    'GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring = "~Error~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")
        End Try

    End Function

    Private Function GetValidateDate(ByRef pStrDate As String) As Boolean

        Try
            ''Commented and Added by Lakshmi dtd 08-05-12
            'strInputDateFormat = strInputDateFormat.ToUpper()


            strInputDateFormat = strInputDateFormat.ToUpper()
            ''-

            Dim TmpstrInputDateFormat() As String
            Dim TempStrDateValue() As String = pStrDate.Split(" ")

            If InStr(TempStrDateValue(0), "/") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("/")
                TmpstrInputDateFormat = strInputDateFormat.Split("/")
            ElseIf InStr(TempStrDateValue(0), "-") > 0 Then
                TempStrDateValue = TempStrDateValue(0).Split("-")
                If strInputDateFormat.Contains("-") Then
                    TmpstrInputDateFormat = strInputDateFormat.Split("-")
                Else
                    TmpstrInputDateFormat = strInputDateFormat.Split("/")
                End If

                '   TmpstrInputDateFormat = strInputDateFormat.Split("-")
            End If

            Dim HsUserDate As New Hashtable
            Dim HsSystemDate As New Hashtable
            Dim StrFinalDate As String

            If TempStrDateValue.Length = 3 Then
                For IntStr As Integer = 0 To TempStrDateValue.Length - 1
                    HsUserDate.Add(GetShort(TmpstrInputDateFormat(IntStr)), TempStrDateValue(IntStr))
                Next
                Dim SysDate() As String
                Dim dtSys As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
                If InStr(dtSys, "/") > 0 Then
                    SysDate = dtSys.Split("/")
                ElseIf InStr(dtSys, "-") > 0 Then
                    SysDate = dtSys.Split("-")
                End If

                StrFinalDate = ""
                For IntStr As Integer = 0 To SysDate.Length - 1
                    If StrFinalDate = "" Then
                        StrFinalDate += HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    Else
                        StrFinalDate += "/" & HsUserDate(GetShort(SysDate(IntStr))).ToString().Trim()
                    End If
                Next

                Try
                    ''pStrDate = Format(CDate(StrFinalDate), "dd/MM/yyyy")
                    pStrDate = CDate(StrFinalDate)
                    'InputDate = CDate(StrFinalDate)
                    GetValidateDate = True

                Catch ex As Exception
                    GetValidateDate = False

                End Try
            Else
                GetValidateDate = False
            End If

        Catch ex As Exception
            GetValidateDate = False

        End Try
    End Function
    Private Function GetShort(ByVal pStr As String) As String

        pStr = pStr.ToUpper

        If InStr(pStr, "D") > 0 Then
            GetShort = "D"
        ElseIf InStr(pStr, "M") > 0 Then
            GetShort = "M"
        ElseIf InStr(pStr, "Y") > 0 Then
            GetShort = "Y"
        End If

    End Function

    Private Sub AddRowsToDataTable(ByVal pNotValid As Boolean, ByVal Data() As String)
        Try
            If Data Is Nothing Then Exit Sub

            If pNotValid = True Then
                DtUnSucInput.Rows.Add(Data)
            Else
                DtInput_PMT.Rows.Add(Data)
            End If


        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "AddRowsToDataTable")
        End Try
    End Sub

    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String

        Try
            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
            Else
                GetValueFormArray = "~ERROR~"
            End If

        Catch ex As Exception

            GetValueFormArray = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")

        End Try

    End Function

    Public Function IsJustAlpha(ByVal sText As String, ByVal num As Integer, ByVal ReplaceSpace As String, Optional ByVal ShowMsgYN As String = "") As String
        Try
            Dim SpecialCharReplace() As DataRow = Nothing
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""


            'If sText <> "" Then
            For n = 1 To iTextLen
                sChar = Mid(sText, n, 1)
                If ChkText(sChar, num) Then
                    IsJustAlpha = IsJustAlpha + sChar
                Else

                    If ShowMsgYN = "Y" Then
                        IsJustAlpha = "Y"
                        Exit Function
                    Else
                        If ReplaceSpace = "Y" Then
                            IsJustAlpha = IsJustAlpha + " "
                        End If

                    End If

                End If
            Next
            'End If

            If Not IsJustAlpha Is Nothing Then
                Return IsJustAlpha
            Else
                Return ""
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
        End Try
    End Function
    'Public Function IsJustAlpha(ByVal sText As String, ByVal num As Integer, ByVal ReplaceWithSpace As String) As String
    '    Try
    '        Dim iTextLen As Integer = Len(sText)
    '        Dim n As Integer
    '        Dim sChar As String = ""

    '        'If sText <> "" Then
    '        For n = 1 To iTextLen
    '            sChar = Mid(sText, n, 1)
    '            If ChkText(sChar, num) Then
    '                IsJustAlpha = IsJustAlpha + sChar
    '            Else
    '                If (ReplaceWithSpace = "Y") Then
    '                    IsJustAlpha = IsJustAlpha + " "
    '                End If

    '            End If
    '        Next
    '        'End If

    '        If Not IsJustAlpha Is Nothing Then
    '            Return IsJustAlpha
    '        Else
    '            Return ""
    '        End If


    '    Catch ex As Exception
    '        Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
    '    End Try
    'End Function

    Private Function ChkText(ByVal sChr As String, ByVal num As Integer) As Boolean

        Try
            Select Case num
                Case 1
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 2
                    '- amount field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[.]" 'Or sChr Like "[,]"
                    'ChkText = True
                Case 3
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[ ]" Or sChr Like "[.]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[:]"
                    'ChkText = True
                Case 4
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 5
                    '- number field
                    ChkText = sChr Like "[0-9]"
                    'ChkText = True
                Case 6
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]"
                    'ChkText = True
                Case 7
                    '- Date field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[:]" Or sChr Like "[/]" Or sChr Like "[\]" Or sChr Like "[-]" Or sChr Like "[.]"
                    'ChkText = True
                Case 8
                    '- alhpa numeric field & All Characters on Keyboard
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[+]" Or sChr Like "[/]" Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[_]" Or sChr Like "[&]" Or sChr Like "[$]" Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[\]" Or sChr Like "[[]" Or sChr Like "[]]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[']" Or sChr Like "[ ]" Or sChr Like "[;]" Or sChr Like "[#]" Or sChr Like "[%]" Or sChr Like "[^]" Or sChr Like "[*]" Or sChr Like "[=]" Or sChr Like "[|]"
                    'ChkText = True
                Case 9
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case 10
                    '- alhpa and numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[-]" Or sChr Like "[ ]" Or sChr Like "[_]"

                Case 11
                    '- alhpa numeric field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[,]" Or sChr Like "[ ]" Or sChr Like "[.]"
                Case 12
                    '- address field
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" Or sChr Like "[{]" Or sChr Like "[}]" Or sChr Like "[|]" Or sChr Like "[!]" Or sChr Like "[#]" Or sChr Like "[@]" Or sChr Like "[-]" Or sChr Like "[?]" Or sChr Like "[:]" Or sChr Like "[%]" Or sChr Like "[ ]"
                    'ChkText = True
                Case 13
                    '- name field 
                    ChkText = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[ ]"
                Case 14
                    '- Bene ID
                    ChkText = sChr Like "[0-9]" Or sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[_]" Or sChr Like "[-]" Or sChr Like "[/]"
                Case 15
                    '- PayDate
                    ChkText = sChr Like "[0-9]" Or sChr Like "[/]" Or sChr Like "[|]" Or sChr Like "[~]"
                Case 16  ''''If amount in (-) minus
                    '- amount field
                    ChkText = sChr Like "[0-9]" Or sChr Like "[.]" Or sChr Like "[-]"
                    'ChkText = True
                Case Else
                    ChkText = False
            End Select

            Return ChkText

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "ChkText")
        End Try
    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "RemoveJunk")

        End Try
    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        ''Added By Jaiwant dtd  03-Dec-2010  ''To remove Junk Characters

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" _
        Or sChr Like "[ ]" 'commented dtd 03-06-2011

    End Function


    Public Function SpCharValidation(ByVal StringValue As String, ByRef _dtSpChar As DataTable) As String

        ''Added by Jaiwant dtd  03-Dec-2010
        Dim ArrSpChar(0) As String
        Dim intSpCharRow As Integer
        ''---
        ClearArray(ArrSpChar)
        Array.Resize(ArrSpChar, _dtSpChar.Select.Length)
        intSpCharRow = 0

        For Each SVRow As DataRow In _dtSpChar.Rows
            ArrSpChar(intSpCharRow) = SVRow(0).ToString
            intSpCharRow += 1
        Next

        ''Added By Jaiwant dtd  03-Dec-2010 ''For All Special Characters
        Dim StrOriginalValue As String = ""
        Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&", " "}

        Try
            ''To remove special chars from array which need to ignore.
            For iIChar As Int16 = 0 To ArrSpChar.Length - 1
                For iSChar As Int16 = 0 To arrSpecialChar.Length - 1
                    If ArrSpChar(iIChar) = arrSpecialChar(iSChar) Then
                        arrSpecialChar(iSChar) = Nothing
                    End If
                Next
            Next
            SpCharValidation = ""
            Dim i As Integer
            For i = 0 To arrSpecialChar.Length - 1
                If InStr(StringValue, arrSpecialChar(i), CompareMethod.Binary) <> 0 Then
                    SpCharValidation = SpCharValidation & arrSpecialChar(i)
                End If
            Next

            Return SpCharValidation

        Catch ex As Exception

            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "SpCharValidation")

        End Try
    End Function
    Public Function RemoveSplChar(ByVal sText As String, ByVal intType As Integer) As String
        ''-To remove Junk Characters-
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsSplChar(sChar, intType) = True Then
                        RemoveSplChar = RemoveSplChar & sChar
                    Else
                        RemoveSplChar = RemoveSplChar & " "
                    End If
                Next
            Else
                RemoveSplChar = ""
            End If

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveSplChar")

        End Try

    End Function

    Private Function IsSplChar(ByVal strChar As String, ByVal intType As Integer) As Boolean


        Select Case intType

            Case 1

                IsSplChar = strChar Like "[0-9]"

            Case 2

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]"

            Case 3

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[/]" _
                            Or strChar Like "[:]" Or strChar Like "[-]" Or strChar Like "[?]" Or strChar Like "[+]" _
                            Or strChar Like "[(]" Or strChar Like "[)]" Or strChar Like "[.]" Or strChar Like "[,]"
            Case 4

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[/]" Or strChar Like "[-]"

            Case 5

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[.]"

            Case 6
                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[/]" _
                            Or strChar Like "[:]" Or strChar Like "[(]" Or strChar Like "[)]" Or strChar Like "[.]" Or strChar Like "[,]"

            Case 7
                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" _
                             Or strChar Like "[.]" Or strChar Like "[_]" Or strChar Like "[@]"
            Case 8

                IsSplChar = strChar Like "[0-9]" Or strChar Like "[a-z]" Or strChar Like "[A-Z]" Or strChar Like "[.]"
        End Select
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call ObjBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function

#Region " IDisposable Support "

    Public Sub Dispose() Implements IDisposable.Dispose
        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtMaster Is Nothing Then DtMaster.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        ''Added by Jaiwant dtd  03-Dec-2010
        If Not DtSpCharValidation Is Nothing Then DtSpCharValidation.Dispose()
        ''----
        If Not DtInput_ADV Is Nothing Then DtInput_ADV.Dispose()
        If Not DtInput_PMT Is Nothing Then DtInput_PMT.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp_PMT Is Nothing Then DtTemp_PMT.Dispose()
        If Not DtTemp_ADV Is Nothing Then DtTemp_ADV.Dispose()
        DtMaster = Nothing
        ObjBaseClass = Nothing
        DtValidation = Nothing
        ''Added by Jaiwant dtd  03-Dec-2010
        DtSpCharValidation = Nothing
        ''----
        DtInput_ADV = Nothing
        DtInput_PMT = Nothing
        DtUnSucInput = Nothing
        DtTemp_PMT = Nothing
        DtTemp_ADV = Nothing
        GC.SuppressFinalize(Me)
    End Sub
    Private Function GetSubstring1(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring1 = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring1 = "~Error~"
                Else
                    'GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring1 = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring1 = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring1 = "~Error~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")
        End Try

    End Function




    Function RupeesToWord(ByVal MyNumber)
        Dim Temp
        Dim Rupees, Paisa As String
        Dim DecimalPlace, iCount
        Dim Hundreds, Words As String
        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Arab "
        place(8) = " Kharab "
        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                'RupeesToWord = "Rupees " & Words & Paisa & " Only"
                RupeesToWord = Words & Paisa

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    'RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    RupeesToWord = Words & Paisa
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundreds = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or
                Trim(Words) = "Lakh  Thousand" Or
                Trim(Words) = "Lakh" Or
                Trim(Words) = "Crore" Or
                Trim(Words) = "Crore  Lakh  Thousand" Or
                Trim(Words) = "Arab  Crore  Lakh  Thousand" Or
                Trim(Words) = "Arab" Or
                Trim(Words) = "Kharab  Arab  Crore  Lakh  Thousand" Or
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or
                   Trim(Words) = "Lakh  Thousand" Or
                   Trim(Words) = "Lakh" Or
                   Trim(Words) = "Crore" Or
                   Trim(Words) = "Crore  Lakh  Thousand" Or
                   Trim(Words) = "Arab  Crore  Lakh  Thousand" Or
                   Trim(Words) = "Arab" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or
                       Trim(ConvertTens(Temp) & place(iCount)) = "Arab" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        'RupeesToWord = "Rupees " & Words & Hundreds & Paisa & " Only"
        RupeesToWord = Words & Hundreds & Paisa
    End Function

    Private Function ConvertDigit(ByVal MyDigit)
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function

    Private Function ConvertTens(ByVal MyTens)
        Dim Result As String

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertHundreds(ByVal MyNumber)
        Dim Result As String

        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then Exit Function

        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred And "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Public Function IsJustAlpha1(ByVal sText As String, ByVal num As Integer, ByVal ReplaceSpace As String, Optional ByVal ShowMsgYN As String = "") As String
        Try
            Dim SpecialCharReplace() As DataRow = Nothing
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""


            'If sText <> "" Then
            For n = 1 To iTextLen
                sChar = Mid(sText, n, 1)
                If ChkText(sChar, num) Then
                    IsJustAlpha1 = IsJustAlpha1 + sChar
                Else

                    If ShowMsgYN = "Y" Then
                        IsJustAlpha1 = "Y"
                        Exit Function
                    Else
                        If ReplaceSpace = "Y" Then
                            IsJustAlpha1 = IsJustAlpha1 + " "
                        End If

                    End If

                End If
            Next
            'End If

            If Not IsJustAlpha1 Is Nothing Then
                Return IsJustAlpha1
            Else
                Return ""
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "IsJustAlpha")
        End Try
    End Function
#End Region

End Class
