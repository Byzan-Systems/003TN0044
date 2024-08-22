Imports System.IO
Imports System.Text
Imports System.Drawing
Imports System
Imports System.Data
Imports Microsoft.Office.Interop

Module GenrateOutput

    Dim objLogCls As New ClsErrLog
    Dim objGetSetINI As ClsShared
    Dim objBaseClass As ClsBase
    Dim objValidationClass As ClsValidation
    Dim SumOfAmount As Double = 0

    Public Function GenerateOutPutFile(ByRef dtOutput_PMT As DataTable, ByRef dtOutput_ADV As DataTable, ByVal strFileName As String) As Boolean
        Dim gstrA2Afile As String = String.Empty
        Dim objStrmWriter As StreamWriter
        Dim strMethodCalForEpay As Boolean = False

        Try
            objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            objValidationClass = New ClsValidation(strFileName, objBaseClass.gstrIniPath)
            objStrmWriter = New StreamWriter(strOutputFolderPath & "\" & strFileName & ".txt")
            'FileCounter = objBaseClass.GetINISettings("General", "File Counter", My.Application.Info.DirectoryPath & "\settings.ini")
            'FileCounter = FileCounter + 1

            For i = 0 To dtOutput_PMT.Rows.Count - 1
                Dim Row_adv = dtOutput_ADV.Select("[Payment Document No.]='" & dtOutput_PMT.Rows(i)("Payment document no.").ToString & "'")

                Dim strOutPutLine = ""
                'Payment Writing
                For IntA As Int32 = 0 To dtOutput_PMT.Rows(i).ItemArray.Length - 1
                    strOutPutLine = strOutPutLine & (dtOutput_PMT.Rows(i)(IntA).ToString()) & "|"
                Next
                For IntA As Int32 = 0 To strOutPutLine.Length - 1
                    If (strOutPutLine.Substring(strOutPutLine.Length - 2, 2)).Contains("||") Then
                        strOutPutLine = strOutPutLine.Substring(0, strOutPutLine.Length - 2)
                    Else
                        Exit For
                    End If
                Next
                If (Not strOutPutLine.Substring(strOutPutLine.Length - 1, 1).Contains("|")) Then
                    strOutPutLine = strOutPutLine & "|"
                End If
                objStrmWriter.WriteLine(strOutPutLine, strFileName)
                strOutPutLine = "!"
                'Advice Writing
                For IntA As Int32 = 0 To Row_adv.Length - 1
                    strOutPutLine = "!"
                    For IntAA As Int32 = 0 To Row_adv(IntA).ItemArray.Length - 1
                        strOutPutLine = strOutPutLine & (Row_adv(IntA).ItemArray(IntAA).ToString()) & "|"
                    Next
                    For IntA1 As Int32 = 0 To strOutPutLine.Length - 1
                        If (strOutPutLine.Substring(strOutPutLine.Length - 2, 2)).Contains("||") Then
                            strOutPutLine = strOutPutLine.Substring(0, strOutPutLine.Length - 2)
                        Else
                            Exit For
                        End If
                    Next
                    If (Not strOutPutLine.Substring(strOutPutLine.Length - 1, 1).Contains("|")) Then
                        strOutPutLine = strOutPutLine & "|"
                    Else
                        Exit For
                    End If
                    'strFileName = strFileName & System.DateTime.Now.ToString("yyyyMMddHHmmss") & ".txt"
                    objStrmWriter.WriteLine(strOutPutLine, strFileName)
                Next
            Next
            objStrmWriter.Close()
            objStrmWriter.Dispose()
            GenerateOutPutFile = True

        Catch ex As Exception
            GenerateOutPutFile = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "GenerateOutPutFile")
            objStrmWriter.Close()
            objStrmWriter.Dispose()
        End Try


    End Function


    Public Function Generate_Output_Response(ByRef _dtRes As DataTable, ByRef _dtHeader As DataTable, ByVal strRespFileName As String) As Boolean
        Dim strOutPutLine As String
        Dim objStrmWriter As StreamWriter
        Dim strRevResp_OptFileName As String = ""
        Try
            If _dtRes.Rows.Count > 0 Then

                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
                strRevResp_OptFileName = Path.GetFileNameWithoutExtension(strRespFileName) & ".dat"

                objStrmWriter = New StreamWriter(strReverseResponseFolderPath & "\" & strRevResp_OptFileName)
                objBaseClass.LogEntry("Reverse Output File generating process Started...")

                '''''Header section
                strOutPutLine = ""
                For Inti As Int32 = 0 To _dtRes.Columns.Count - 1
                    strOutPutLine = strOutPutLine & (_dtRes.Columns(Inti).ColumnName.ToString()) & "" & vbTab & ""
                Next

                'strOutPutLine = Left(strOutPutLine, strOutPutLine.Length - 1)
                'objStrmWriter.WriteLine(strOutPutLine, strRevResp_OptFileName)

                '''''Detail section
                For Each drRow As DataRow In _dtRes.Rows
                    strOutPutLine = ""
                    If (Not drRow.ItemArray(0).ToString() = "") Then
                        For Inti As Int32 = 0 To drRow.ItemArray.Length - 1
                            strOutPutLine = strOutPutLine & (drRow.ItemArray(Inti).ToString()) & "" & vbTab & ""
                        Next

                        'strOutPutLine = Left(strOutPutLine, strOutPutLine.Length - 1)
                        objStrmWriter.WriteLine(strOutPutLine, strRevResp_OptFileName)

                    End If
                Next

                If Not objStrmWriter Is Nothing Then
                    objStrmWriter.Close()
                    objStrmWriter.Dispose()

                End If
                objBaseClass.LogEntry("Reverse Response Output File [" & strRevResp_OptFileName & "] is generated successfully", False)

                Generate_Output_Response = True
            Else
                objBaseClass.LogEntry("Reverse Response Record Not Found")
            End If


        Catch ex As Exception
            Generate_Output_Response = False
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "GenerateOutput", "Generate_Output_Response")
        Finally
            If Not objStrmWriter Is Nothing Then
                objStrmWriter.Close()
                objStrmWriter.Dispose()

            End If
        End Try
    End Function

    Public Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then

                ' Check_Comma = Chr(34) & strTemp & Chr(34) & ","
                Check_Comma = strTemp
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "Check_Comma")

        End Try
    End Function

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            Call objBaseClass.Handle_Error(ex, "frmGenericRBI", Err.Number, "Pad_Length")

        End Try
    End Function


    Function RemoveCharacter(ByVal stringToCleanUp As String)
        Dim characterToRemove As String = ""
        characterToRemove = Chr(34) + "=~^!#$%&'()*+,-@`/\:{}[]"

        Dim firstThree As Char() = characterToRemove.Take(30).ToArray()
        For index = 0 To firstThree.Length - 1
            stringToCleanUp = stringToCleanUp.ToString.Replace(firstThree(index), "")
        Next
        Return stringToCleanUp
    End Function
End Module
