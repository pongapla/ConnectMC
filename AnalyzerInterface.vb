Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports ValueSoft.DALManage
Imports DeviceControlApp.ComPortMonitor
Imports System.Threading
Imports System.Threading.Thread

Public Class AnalyzerInterface
    'Dim regis As Core.LIS.Registration
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    'Tui add 2013-09-18
    Public Property builder As StringBuilder
    Public Property parm As IDbDataParameter()
    Public Property SqlProvider As ValueSoft.DALManage.SqlDataProvider
    Dim dtTestResult As New DataTable
    Dim dtTestResultToUpdate As New DataTable
    Dim conString As String = MyGlobal.myConnectionString
    Dim resultDate As DateTime = DateTime.Now()
    Dim analyzerDate As DateTime = DateTime.Now()
    Dim comPort As String
    Dim analyzerSkey As Int32
    Dim analyzerModel As String
    Dim clsMon As ComPortMonitor

    Sub New(ByVal argComport As String, ByVal argAnalyzerSkey As Int32, ByVal argAnalyzerModel As String)
        'regis = New Core.LIS.Registration(_lis)
        comPort = argComport
        analyzerSkey = argAnalyzerSkey
        analyzerModel = argAnalyzerModel

        If My.Settings.TrackError Then
            log.Info(comPort & ", " & analyzerSkey & ", " & analyzerModel)
        End If
    End Sub

    Public Function ExtractResult(ByVal DeviceStr As String) As Boolean

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim ResultStatus As Boolean = False

        Try

            If AppSetting.TestLisConnection = False Then
                If My.Settings.TrackError Then
                    log.Info(comPort & "Cannot connect database")
                End If
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ExtractResult: Cannot connect database !")
                Return True
            End If
            If analyzerModel = "AS300" Or analyzerModel = "AS720" Or analyzerModel = "H-500" Or analyzerModel = "Uriscan" Or analyzerModel = "U120" Then
                'Golf add H-500, U120 2016-04-20
                Return ExtractResult_Micro(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "RxLyte" Then
                Return ExtractResult_RxLyte(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Daytona" Or analyzerModel = "DaytonaPlus" Or analyzerModel = "Imola" Or analyzerModel = "Suzuka" Or analyzerModel = "DX300" Or analyzerModel = "XS-Serie" Or analyzerModel = "Liaison" Or analyzerModel = "COBAS_e411" Or analyzerModel = "SUZUKA" Or analyzerModel = "RXMODENA" Or analyzerModel = "Modena" Or analyzerModel = "COBAS_C311" Or analyzerModel = "astmsim" Then
                'If My.Settings.TrackError Then
                '    log.Info(comPort & "before ExtractResult_Chem")
                'End If
                Return ExtractResult_Chem(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "LD-500" Then
                Return ExtractResult_HbA1c(DeviceStr)
                Exit Function
                'ElseIf analyzerModel = "BF-5180" Then
            ElseIf analyzerModel = "BC-5600" Then
                'Return ExtractResult_BF5180(analyzerModel, DeviceStr)
                'Exit Function 
            ElseIf analyzerModel = "ISE6000" Then
                Return ExtractResult_ISE6000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Dimension" Then
                Return ExtractResult_Dimension(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "XS" Then
                Return ExtractResult_Chem(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Architecti1000" Then
                Return ExtractResult_Architecti1000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Architecti2000" Then
                Return ExtractResult_Architecti2000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "C4000" Then
                Return ExtractResult_C4000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "PATHFAST" Then
                Return ExtractResult_PATHFAST(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "ABXPentraXL80" Then
                Return ExtractResult_ABXPentraXL80(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "C8000" Then
                Return ExtractResult_C8000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "IQ200" Then 'Golf IQ200 2020-10-15
                Return ExtractResult_IQ200(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Echo" Then
                Return ExtractResult_Echo(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "HA8180V" Then 'PLOY ADD 2020.10.05
                Return ExtractResult_HA8180V(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Stago" Then 'PLOY ADD 2020.11.20
                Return ExtractResult_Stago(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "DxC700AU" Then 'Pong ADD 2021.02
                Return ExtractResult_DxC700AU(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "CoagACL" Then
                Return ExtractResult_CoagACL(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "GEM4000" Then
                Return ExtractResult_GEM4000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "V3600" Then 'Pong ADD 2021.08
                Return ExtractResult_V3600(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "XN550" Then
                Return ExtractResult_XN550(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "XN3000" Then 'Pong ADD 2021.09
                Return ExtractResult_XN3000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "HCLAB" Then 'Pong ADD 2021.11
                Return ExtractResult_HCLAB(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "XN1000" Then 'Pong ADD 2021.11
                Return ExtractResult_XN1000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "AU5800" Then 'Pong ADD 2022.01.02
                Return ExtractResult_AU5800(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "AU480" Then 'Pong ADD 2022.03.14
                Return ExtractResult_AU480(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "UriscanPRO" Then 'Pong ADD 2022.01.11
                Return ExtractResult_UriscanPRO(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "BACTEC9000" Then 'Pong ADD 2022.02.
                Return ExtractResult_BACTEC9000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "bloodgasRapidlab348" Then 'paramet add
                Return ExtractResult_bloodgas(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Rapidlab348EX1.00" Then 'paramet add (Have EOT)
                Return ExtractResult_Rapidlab348EX100(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Rapidlab348EX1.32" Then 'paramet add (None EOT)
                Return ExtractResult_Rapidlab348EX132(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "D-10" Or analyzerModel = "D10" Then 'Arm add 2021.11.23  Compile By Pong 2022.02.24
                Return ExtractResult_D10(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "CA620" Then 'Arm add 2021.11.23 Compile By Pong 2022.02.25
                Return ExtractResult_CA620(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "ElyteISE6000" Then 'Arm add 2021.11.23 Compile By Pong 2022.02.25
                Return ExtractResult_ElyteISE6000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "ElitePro" Then 'Pong add 2022-03-16
                Return ExtractResult_ElitePro(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Cobas6000" Then 'Pong add 2022-03-17
                Return ExtractResult_Cobas6000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "CaretiumXI-1021B" Then 'Pong add 2022-03-25
                Return ExtractResult_CaretiumXI(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Access2" Then 'Pong add 2022-06-07
                Return ExtractResult_Access2(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "A1000" Then 'Pong add 2022-06-21
                Return ExtractResult_A1000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "SF8050" Then 'Pong add 2022-06-27
                Return ExtractResult_SF8050(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "DxI800" Then 'Pong add 2022-06-27
                Return ExtractResult_DxI800(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "XL1000" Then 'Pong add 2022-10-04
                Return ExtractResult_XL1000(DeviceStr)
                Exit Function
            ElseIf analyzerModel = "Q4lyte" Then 'Pong add 2022-10-10
                Return ExtractResult_Q4lyte(DeviceStr)
                Exit Function
            Else


                'Mythic
            End If

            'If DeviceStr.IndexOf("END_RESULT") = -1 Then
            '    Return False
            'End If

            Dim i, j As Integer

            LineArr = DeviceStr.Split(Environment.NewLine)
            j = LineArr.Count

            '##1## Get Lab No.
            For i = 0 To j - 1
                ItemStr = LineArr(i)
                ItemArr = ItemStr.Split(";")

                If ItemArr.Count <= 0 Then
                    Continue For
                End If

                If ItemArr(0) = "SID" Then
                    SpecimenId = ItemArr(1).Trim
                    'SpecimenId = SpecimenId.Substring(0, SpecimenId.Length - 1)
                    Exit For
                End If
            Next

            ' Mark ไว้ก่อน กรณีไม่เจอ Lab No.
            'If SpecimenId = "" Or SpecimenId.Length < 9 Then
            '    Return True
            '    Exit Function
            'End If

            'Tui Add
            'ดึงค่า Result ทั้งหมดที่ Order ไว้
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "N").Copy
            Dim dvTestResult As New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            For i = 0 To j - 1

                ItemStr = LineArr(i)
                ItemArr = ItemStr.Split(";")

                If ItemArr.Count <= 0 Then
                    Continue For
                End If

                If ItemArr(0) = "PID" Then
                    Continue For
                End If

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(0) & "'"
                If dvTestResult.Count = 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = ItemArr(1)
                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = resultDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)
                End If

            Next
            'Tui Add

            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()

            'Tui add 2015-09-22  auto update formula
            UpdateFormula(SpecimenId)

        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ExtractResult: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return True
    End Function

    Public Function ExtractResult_Micro(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemStrBLD As String
        Dim ItemArr() As String
        Dim ItemArrBLD() As String
        Dim ItemArrGLU() As String
        Dim SpecimenId As String = ""
        Dim Sex As String = ""
        Dim ResultStatus As Boolean = False
        Dim resultCode As String = ""
        Dim resultValue As String = ""
        Dim isFoundHbA1c As Boolean = False
        Dim dtGetH500Check As DataTable
        dtGetH500Check = GetUA_LIS_FORMULA_IN_INTERFACE_ANALYZER()
        If analyzerModel = "AS300" Or analyzerModel = "AS720" Then
            If DeviceStr.IndexOf("LOT") = -1 Then
                Return False
            End If
        ElseIf analyzerModel = "H-500" Then

            'dtGetH500Check = GetUA_LIS_FORMULA_IN_INTERFACE_ANALYZER()

            'If DeviceStr.IndexOf(Chr(5)) <> -1 Then
            '    Dim fm As New frmMain
            '    fm = DirectCast(Application.OpenForms("frmMain"), frmMain)
            '    fm.StopMonitor()
            '    fm.Start()
            '    Return False
            'End If
            ' Golf 2018-11-12
            If DeviceStr.IndexOf(Chr(5)) <> -1 Then
                Return False
            End If
            If DeviceStr.IndexOf(Chr(3)) = -1 Then
                Return False
            End If


            'Tui Add Uriscan, Mar 26, 2015
        ElseIf analyzerModel = "Uriscan" Then
            If DeviceStr.IndexOf("CLA") = -1 Then
                Return False
            End If
            'End Tui Add Uriscan, Mar 26, 2015
            'Golf add U120 2016-04-19
        ElseIf analyzerModel = "U120" Then
            'If DeviceStr.IndexOf("CRE") = -1 Then
            If DeviceStr.IndexOf(Chr(3)) = -1 Then
                Return False
            End If
            'Golf add U120 2016-04-19
        End If

        Dim i, j, k As Integer

        LineArr = DeviceStr.Split(Environment.NewLine)
        j = LineArr.Count

        'Get QC Date
        Dim tempQcDate As String = String.Empty
        Dim tempQcDateArray() As String = Nothing
        'Get QC Date

        '##1## Get Lab No.
        For i = 0 To j - 1
            If analyzerModel = "AS300" Or analyzerModel = "AS720" Then

                'Get QC Date
                If i = 0 Then
                    tempQcDate = LineArr(i)
                    tempQcDate = tempQcDate.Replace("~", String.Empty)
                    tempQcDate = tempQcDate.Trim()
                    If tempQcDate.Length = 20 Then
                        DateTime.TryParse(tempQcDate, analyzerDate)
                        Continue For
                    End If
                End If

                'End Get QC Date

                If LineArr(i).IndexOf("ID") > -1 Then
                    LineArr(i) = LineArr(i).Replace("ID(", String.Empty)
                    LineArr(i) = LineArr(i).Replace(")", String.Empty)
                    LineArr(i) = LineArr(i).Trim
                    SpecimenId = LineArr(i)
                    'SpecimenId = SpecimenId.Substring(0, SpecimenId.Length - 1)
                    Continue For
                End If

            ElseIf analyzerModel = "H-500" Then
                If LineArr(i).IndexOf("ID") > -1 Then
                    LineArr(i) = LineArr(i).Replace("ID:", String.Empty)
                    LineArr(i) = LineArr(i).Replace("-", String.Empty)
                    LineArr(i) = LineArr(i).Trim
                    SpecimenId = LineArr(i)
                    'Get QC Date 
                    tempQcDate = LineArr(i - 2).Trim.Replace("Date:", "") ' Golf 2016-04-04
                    'tempQcDate = Now() ' Golf 2019-07-26
                    analyzerDate = DateTime.Parse(tempQcDate)
                    'End Get QC Date

                    Continue For
                End If
                'Golf add U120 2016-04-19
            ElseIf analyzerModel = "U120" Then
                If LineArr(i).IndexOf("ID") > -1 Then
                    LineArr(i) = LineArr(i).Replace("ID:", String.Empty)
                    LineArr(i) = LineArr(i).Trim
                    SpecimenId = LineArr(i).Replace(Chr(2), String.Empty).Trim

                    'Get QC Date  
                    tempQcDate = LineArr(i + 1).Trim.Replace("Date:", "") ' Golf 2016-04-04
                    Dim tempQcDateDay As String = tempQcDate.Substring(0, 2).Trim
                    Dim tempQcDateMonth As String = tempQcDate.Substring(3, 2).Trim
                    Dim tempQcDateYear As String = tempQcDate.Substring(6, 4).Trim
                    Dim tempQcDateTime As String = tempQcDate.Substring(11, 5).Trim
                    tempQcDate = tempQcDateYear + "-" + tempQcDateMonth + "-" + tempQcDateDay + " " + tempQcDateTime
                    analyzerDate = DateTime.Parse(tempQcDate)
                    'End Get QC Date

                    Continue For
                End If
                'Golf add U120 2016-04-19
                'Tui Add Uriscan, Mar 26, 2015 
            ElseIf analyzerModel = "Uriscan" Then

                'Get QC Date
                If LineArr(i).IndexOf("Date") > -1 Then

                    tempQcDate = LineArr(i)
                    tempQcDate = tempQcDate.Replace("Date :", "|")
                    tempQcDateArray = tempQcDate.Split("|")
                    tempQcDate = tempQcDateArray(1)

                    analyzerDate = DateTime.Parse(tempQcDate)
                    Continue For
                End If
                'End Get QC Date

                If LineArr(i).IndexOf("ID") > -1 Then
                    LineArr(i) = LineArr(i).Substring(12, Len(LineArr(i)) - 12)
                    'LineArr(i) = LineArr(i).Replace("ID_NO:0012-", String.Empty)
                    LineArr(i) = LineArr(i).Trim
                    SpecimenId = LineArr(i)
                    'SpecimenId = SpecimenId.Substring(0, SpecimenId.Length - 1)
                    Continue For
                End If
                'End Tui Add Uriscan, Mar 26, 2015

            End If
        Next

        If SpecimenId = "" Then
            Return True
            Exit Function
        End If

        'Debug.WriteLine("SpecimenId = " & SpecimenId)

        'Tui Add
        'ดึงค่า Result ทั้งหมดที่ Order ไว้
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        '##2## Get test result by test code
        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemStr = LineArr(i).TrimStart
            ItemStr = ItemStr.TrimEnd
            ItemStrBLD = LineArr(i)
            ItemStrBLD = LineArr(i).TrimStart
            ItemStrBLD = ItemStrBLD.TrimEnd
            'Golf 2016-04-04 H-500 Remove value+
            If analyzerModel = "H-500" Then
                If System.Text.RegularExpressions.Regex.Replace(ItemStr, "\s+", " ").Split(" ").Length > 2 And (dtGetH500Check.Rows.Count = 0 Or dtGetH500Check.Rows(0).Item("parm_value").ToString() = "N") Then
                    If ItemStr.IndexOf("+") > -1 Then 'Golf mark
                        ItemStr = ItemStr.Replace(ItemStr.Substring(ItemStr.IndexOf(" "), ItemStr.IndexOf("+")).Trim, " ")
                    End If
                    If ItemStr.IndexOf("Normal") > -1 Then 'Golf mark 2016-11-24
                        ItemStr = ItemStr.Replace("Normal", " ")
                    End If
                End If

            Else
                'Tui add replace + - sign
                ItemStr = ItemStr.Replace("+-", "")
                ItemStr = ItemStr.Replace("+", "")
            End If
            'Golf 2016-04-04""

            'Tui Edit 2015-05-26, Exception Case Result from Uriscan
            If analyzerModel = "Uriscan" And Not ItemStr.Contains("URO") Then
                ItemStr = ItemStr.Replace("norm", "")
            End If

            If analyzerModel = "Uriscan" And ItemStr.Contains("URO") Then
                ItemStr = ItemStr.Replace("norm", "0")
            End If

            If analyzerModel <> "Uriscan" Then
                ItemStr = ItemStr.Replace("norm", "")
            End If
            'Tui Edit 2015-05-26, Exception Case Result from Uriscan

            'Tui Edit 2016-02-10, Exception Case Result from Uriscan
            If analyzerModel = "Uriscan" And ItemStr.Contains("NIT") Then
                ItemStr = ItemStr.Replace("pos", "+")
            ElseIf analyzerModel = "H-500" Then
                ItemStr = ItemStr.Replace("pos", "Pos")
            Else
                ItemStr = ItemStr.Replace("pos", "")
            End If
            'Tui Edit 2016-02-10, Exception Case Result from Uriscan

            If analyzerModel = "H-500" Or analyzerModel = "U120" Then
                If ItemStr.IndexOf("A:C") <> -1 Then
                    If ItemStr.IndexOf("<") <> -1 Then
                        ItemStr = "A:C      <30   mg/g"
                    End If
                    If ItemStr.IndexOf(">") <> -1 Then
                        ItemStr = "A:C      >300   mg/g"
                    End If
                    If ItemStr.IndexOf("-") <> -1 Then
                        ItemStr = "A:C      30-300   mg/g"
                    End If
                End If
            End If
            ItemArr = System.Text.RegularExpressions.Regex.Replace(ItemStr, "\s+", " ").Split(" ")
            ItemArrBLD = System.Text.RegularExpressions.Regex.Replace(ItemStrBLD, "\s+", " ").Split(" ")
            If ItemArr.Count <= 1 Then
                Continue For
            End If

            'remove * infront of TestCode 
            resultCode = ItemArr(0).Replace("*", String.Empty).Trim

            k = 1
            Do

                If ItemArr(k) <> "" Then
                    Exit Do
                End If

                If k = ItemArr.Count - 1 Then
                    Exit Do
                End If

                k += 1
            Loop
            Select Case analyzerModel
                Case "H-500"
                    '"H-500" 
                    If dtGetH500Check.Rows.Count >= 0 Then
                        If dtGetH500Check.Rows(0).Item("parm_value").ToString() = "Y" Then
                            Select Case resultCode
                                Case "PRO", "*PRO"
                                    Select Case ItemArr(k)
                                        Case "Neg", "0"
                                            resultValue = "Negative"
                                        Case "Trace", "0.5", "+-"
                                            resultValue = "Trace"
                                        Case "30", "1"
                                            resultValue = "1+"
                                        Case "100", "2"
                                            resultValue = "2+"
                                        Case "300", "3"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else
                                            If ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                resultValue = GetResultMicro(ItemArr(k))
                                            End If
                                    End Select
                                Case "GLU", "*GLU"
                                    Select Case ItemArr(k)
                                        Case "Neg", "0"
                                            resultValue = "Negative"
                                        Case "Trace", "+-" 'Golf add "+-" 2017-09-11
                                            resultValue = "Trace"
                                        Case "100", "1"
                                            resultValue = "1+"
                                        Case "250", "2"
                                            resultValue = "2+"
                                        Case "500", "3"
                                            resultValue = "3+"
                                        Case "1000", "4"
                                            resultValue = "4+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else
                                            If ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 1000 Then
                                                        resultValue = "4+"
                                                    Else
                                                        resultValue = GetResultMicro(ItemArr(k))
                                                    End If
                                                Else
                                                    resultValue = GetResultMicro(ItemArr(k))
                                                End If

                                            End If
                                    End Select
                                Case "KET", "*KET"
                                    Select Case ItemArr(k)
                                        Case "Neg", "0"
                                            resultValue = "Negative"
                                        Case "Trace", "0.5", "5", "+-"
                                            resultValue = "Trace"
                                        Case "1", "15"
                                            resultValue = "1+"
                                        Case "2", "40"
                                            resultValue = "2+"
                                        Case "3", "80"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else
                                            If ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 80 Then
                                                        resultValue = "3+"
                                                    Else
                                                        resultValue = GetResultMicro(ItemArr(k))
                                                    End If
                                                Else
                                                    resultValue = GetResultMicro(ItemArr(k))
                                                End If

                                            End If
                                    End Select
                                Case "UBG", "*UBG"
                                    Select Case ItemArr(k)
                                        Case "0.2", "1", "Normal"
                                            resultValue = "Normal"
                                        Case "2",
                                            resultValue = "1+"
                                        Case "4"
                                            resultValue = "2+"
                                        Case "8"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else
                                            If ItemArr(k).IndexOf("Normal") <> -1 Then
                                                resultValue = "Normal"
                                            ElseIf ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 8 Then
                                                        resultValue = "3+"
                                                    Else
                                                        resultValue = GetResultMicro(ItemArr(k))
                                                    End If
                                                Else
                                                    resultValue = GetResultMicro(ItemArr(k))
                                                End If

                                            End If
                                    End Select
                                Case "NIT", "*NIT"
                                    Select Case ItemArr(k)
                                        Case "1", "pos", "Pos"
                                            resultValue = "Positive"
                                        Case "error"
                                            resultValue = ""
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "BIL", "*BIL"
                                    Select Case ItemArr(k)
                                        Case "1"
                                            resultValue = "1+"
                                        Case "3"
                                            resultValue = "2+"
                                        Case "6"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else

                                            If ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 6 Then
                                                        resultValue = "3+"
                                                    Else
                                                        resultValue = "Negative"
                                                    End If
                                                Else
                                                    resultValue = "Negative"
                                                End If

                                            End If
                                    End Select
                                Case "BLD", "*BLD"
                                    Select Case ItemArr(k)
                                        Case "10", "+-"
                                            resultValue = "Trace"
                                        Case "25"
                                            resultValue = "1+"
                                        Case "80"
                                            resultValue = "2+"
                                        Case "200"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else

                                            If ItemArr(k).IndexOf("+") <> -1 Then

                                                If ItemArr(k).IndexOf("1+") <> -1 Then
                                                    resultValue = "1+"
                                                ElseIf ItemArr(k).IndexOf("2+") <> -1 Then
                                                    resultValue = "2+"
                                                ElseIf ItemArr(k).IndexOf("3+") <> -1 Then
                                                    resultValue = "3+"
                                                Else
                                                    resultValue = ItemArr(k)
                                                End If

                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 200 Then
                                                        resultValue = "3+"
                                                    Else
                                                        resultValue = "Negative"
                                                    End If
                                                Else
                                                    resultValue = "Negative"
                                                End If

                                            End If
                                    End Select
                                Case "LEU", "*LEU"
                                    Select Case ItemArr(k)
                                        Case "0", "Neg", "neg"
                                            resultValue = "Negative"
                                        Case "15", "Trace", "+-"
                                            resultValue = "Trace"
                                        Case "70"
                                            resultValue = "1+"
                                        Case "125"
                                            resultValue = "2+"
                                        Case "500"
                                            resultValue = "3+"
                                        Case "error"
                                            resultValue = ""
                                        Case Else

                                            If ItemArr(k).IndexOf("+") <> -1 Then
                                                resultValue = ItemArr(k)
                                            Else
                                                If IsNumeric(ItemArr(k)) Then
                                                    If ItemArr(k) > 500 Then
                                                        resultValue = "3+"
                                                    Else
                                                        resultValue = "Negative"
                                                    End If
                                                Else
                                                    resultValue = "Negative"
                                                End If

                                            End If
                                    End Select
                                Case "SG", "PH", "pH"

                                    If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                        Select Case ItemArr(k)
                                            Case "Pos", "pos"
                                                resultValue = "Positive"
                                            Case "Neg", "neg"
                                                resultValue = "Negative"
                                            Case "error"
                                                resultValue = ""
                                            Case Else
                                                resultValue = ItemArr(k)
                                        End Select
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                        If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                            If ItemArr(k).IndexOf("=") > -1 Then
                                                resultValue = "<=" + resultValue
                                            Else
                                                resultValue = "<" + resultValue
                                            End If
                                        ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                            If ItemArr(k).IndexOf("=") > -1 Then
                                                resultValue = ">=" + resultValue
                                            Else
                                                resultValue = ">" + resultValue
                                            End If
                                        End If
                                    End If
                                Case "MALB", "ALB", "*ALB"
                                    If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                        Select Case ItemArr(k)
                                            Case "Pos", "pos"
                                                resultValue = "Positive"
                                            Case "Neg", "neg"
                                                resultValue = "Negative"
                                            Case "error"
                                                resultValue = ""
                                            Case Else
                                                resultValue = ItemArr(k)
                                        End Select
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                        If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                            If ItemArr(k).IndexOf("=") > -1 Then
                                                resultValue = "<=" + resultValue
                                            Else
                                                resultValue = "<" + resultValue
                                            End If
                                        ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                            If ItemArr(k).IndexOf("=") > -1 Then
                                                resultValue = ">=" + resultValue
                                            Else
                                                resultValue = ">" + resultValue
                                            End If
                                        End If
                                    End If
                                Case "A:C"
                                    If ItemArr(k) = "30-300" Then
                                        resultValue = ItemArr(k) 'GetResultMicro(ItemArr(k))
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If

                                    If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = "<=" + resultValue
                                        Else
                                            resultValue = "<" + resultValue
                                        End If
                                    ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = ">=" + resultValue
                                        Else
                                            resultValue = ">" + resultValue
                                        End If
                                    End If


                                Case Else
                                    If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                        Select Case ItemArr(k)
                                            Case "Pos", "pos"
                                                resultValue = "Positive"
                                            Case "Neg", "neg"
                                                resultValue = "Negative"
                                            Case "error"
                                                resultValue = ""
                                            Case Else
                                                resultValue = ItemArr(k)
                                        End Select
                                    Else
                                        If ItemArr(k).IndexOf("+") <> -1 Then
                                            resultValue = ItemArr(k)
                                        Else
                                            resultValue = GetResultMicro(ItemArr(k))
                                        End If

                                    End If
                            End Select
                            GoTo endResultUA
                        End If
                    End If

                    If System.Text.RegularExpressions.Regex.Replace(ItemStr, "\s+", " ").Split(" ").Length > 2 Then
                        Select Case resultCode
                            Case "SG", "PH", "pH", "MALB", "ALB"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                    If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = "<=" + resultValue
                                        Else
                                            resultValue = "<" + resultValue
                                        End If
                                    ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = ">=" + resultValue
                                        Else
                                            resultValue = ">" + resultValue
                                        End If
                                    End If
                                End If
                            Case "A:C"
                                If ItemArr(k) = "30-300" Then
                                    resultValue = ItemArr(k) 'GetResultMicro(ItemArr(k))
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                End If

                                If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                    If ItemArr(k).IndexOf("=") > -1 Then
                                        resultValue = "<=" + resultValue
                                    Else
                                        resultValue = "<" + resultValue
                                    End If
                                ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                    If ItemArr(k).IndexOf("=") > -1 Then
                                        resultValue = ">=" + resultValue
                                    Else
                                        resultValue = ">" + resultValue
                                    End If
                                End If

                            Case "NIT"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = ItemArr(k)
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                End If
                            Case "LEU", "*LEU"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    If ItemArr(k).IndexOf("+") <> -1 Then
                                        resultValue = ItemArr(k)
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If
                                End If

                            Case Else
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    If ItemArr(k).IndexOf("+") <> -1 Then
                                        resultValue = ItemArr(k)
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If
                                End If
                        End Select
                    Else
                        Select Case resultCode
                            Case "SG", "PH", "pH"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                    If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = "<=" + resultValue
                                        Else
                                            resultValue = "<" + resultValue
                                        End If
                                    ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = ">=" + resultValue
                                        Else
                                            resultValue = ">" + resultValue
                                        End If
                                    End If
                                End If
                            Case "MALB", "ALB", "*ALB"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = ItemArr(k)
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                    If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = "<=" + resultValue
                                        Else
                                            resultValue = "<" + resultValue
                                        End If
                                    ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                        If ItemArr(k).IndexOf("=") > -1 Then
                                            resultValue = ">=" + resultValue
                                        Else
                                            resultValue = ">" + resultValue
                                        End If
                                    End If
                                End If
                            Case "A:C"
                                If ItemArr(k) = "30-300" Then
                                    resultValue = ItemArr(k) 'GetResultMicro(ItemArr(k))
                                Else
                                    resultValue = GetResultMicro(ItemArr(k))
                                End If

                                If ItemArr(k).IndexOf("<") > -1 Then 'Golf mark
                                    If ItemArr(k).IndexOf("=") > -1 Then
                                        resultValue = "<=" + resultValue
                                    Else
                                        resultValue = "<" + resultValue
                                    End If
                                ElseIf ItemArr(k).IndexOf(">") > -1 Then 'Golf mark
                                    If ItemArr(k).IndexOf("=") > -1 Then
                                        resultValue = ">=" + resultValue
                                    Else
                                        resultValue = ">" + resultValue
                                    End If
                                End If

                            Case "NIT"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = ItemArr(k)
                                Else
                                    If ItemArr(k).IndexOf("+") <> -1 Then
                                        resultValue = ItemArr(k)
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If
                                End If

                            Case "LEU", "*LEU"
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    If ItemArr(k).IndexOf("+") <> -1 Then
                                        resultValue = ItemArr(k)
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If

                                End If
                            Case Else
                                If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                                    resultValue = """" + ItemArr(k) + """"
                                Else
                                    If ItemArr(k).IndexOf("+") <> -1 Then
                                        resultValue = ItemArr(k)
                                    Else
                                        resultValue = GetResultMicro(ItemArr(k))
                                    End If

                                End If
                        End Select
                    End If
endResultUA:
                    '"H-500"
                Case "AS720"
                    If dtGetH500Check.Rows.Count >= 0 Then
                        If dtGetH500Check.Rows(0).Item("parm_value").ToString() = "Y" Then
                            Select Case resultCode
                                Case "pH"
                                    resultValue = GetResultMicro(ItemArr(k))
                                Case "S.G"
                                    resultValue = GetResultMicro(ItemArr(k))
                                Case "URO"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    If IsNumeric(resultValue) Then
                                        If resultValue <= 1 Then
                                            resultValue = "Normal"
                                        End If
                                    End If
                                    Select Case resultValue
                                        Case "1"
                                            resultValue = "Normal"
                                        Case "2"
                                            resultValue = "1+"
                                        Case "4"
                                            resultValue = "2+"
                                        Case "8"
                                            resultValue = "3+"
                                        Case "12"
                                            resultValue = "4+"
                                        Case Else
                                            resultValue = GetResultMicro(ItemArr(k))
                                            If IsNumeric(resultValue) Then
                                                If resultValue <= 1 Then
                                                    resultValue = "Normal"
                                                End If
                                            End If
                                    End Select
                                Case "GLU"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "100"
                                            resultValue = "Trace"
                                        Case "250"
                                            resultValue = "1+"
                                        Case "500"
                                            resultValue = "2+"
                                        Case "1000"
                                            resultValue = "3+"
                                        Case Else
                                            If ItemArrBLD(1) = "+++" Then
                                                resultValue = "3+"
                                            Else
                                                resultValue = "Negative"
                                            End If
                                    End Select
                                Case "KET"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "5"
                                            resultValue = "Trace"
                                        Case "10"
                                            resultValue = "1+"
                                        Case "50"
                                            resultValue = "2+"
                                        Case "100"
                                            resultValue = "3+"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "BIL"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "0.5"
                                            resultValue = "1+"
                                        Case "2"
                                            resultValue = "2+"
                                        Case "3"
                                            resultValue = "3+"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "PRO"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "15"
                                            resultValue = "Trace"
                                        Case "30"
                                            resultValue = "1+"
                                        Case "100"
                                            resultValue = "2+"
                                        Case "300"
                                            resultValue = "3+"
                                        Case "1000"
                                            resultValue = "4+"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "NIT"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "0.05"
                                            resultValue = "Positive"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "BLD"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "5"
                                            resultValue = "Trace"
                                        Case "10"
                                            resultValue = "1+"
                                        Case "50"
                                            If ItemArrBLD(1) = "+++" Then
                                                resultValue = "3+"
                                            Else
                                                resultValue = "2+"
                                            End If

                                        Case "250"
                                            resultValue = "3+"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case "LEU"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "25"
                                            resultValue = "Trace"
                                        Case "75"
                                            resultValue = "1+"
                                        Case "250"
                                            resultValue = "2+"
                                        Case "500"
                                            resultValue = "3+"
                                        Case Else
                                            resultValue = "Negative"

                                    End Select
                                Case "VTC"
                                    resultValue = GetResultMicro(ItemArr(k))
                                    Select Case resultValue
                                        Case "0"
                                            resultValue = "Negative"
                                        Case "10"
                                            resultValue = "1+"
                                        Case "20"
                                            resultValue = "2+"
                                        Case "40"
                                            resultValue = "3+"
                                        Case Else
                                            resultValue = "Negative"
                                    End Select
                                Case Else
                                    resultValue = GetResultMicro(ItemArr(k))
                            End Select



                        Else
                            resultValue = GetResultMicro(ItemArr(k))
                        End If
                    Else
                        resultValue = GetResultMicro(ItemArr(k))
                    End If
                Case Else
                    resultValue = GetResultMicro(ItemArr(k))
            End Select


            If analyzerModel = "H-500" Then
                Select Case resultCode
                    Case "MALB", "ALB"
                        'dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
                        dvTestResult.RowFilter = "analyzer_ref_cd = 'MALB' or analyzer_ref_cd =  'ALB'"
                    Case Else
                        dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
                End Select
            Else
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
            End If


            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = resultValue
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If

        Next

        UpdateTestResultValue(dtTestResultToUpdate, False)
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        UpdateFormula(SpecimenId)

        Return True

    End Function
    Public Function ExtractResult_RxLyte(ByVal DeviceStr As String) As Boolean

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim Sex As String = ""
        Dim ResultStatus As Boolean = False

        'If DeviceStr.IndexOf("AG") = -1 Then
        '    Return False
        'End If

        Dim i, j As Integer
        Dim labTemp As Int32 = 0
        Dim running As Int32 = 0

        'IQC
        Dim rxlyteDate As String = String.Empty
        Dim rxlyteTime As String = String.Empty
        Dim rxlyteQc As String = String.Empty

        LineArr = DeviceStr.Split(Chr(10))
        j = LineArr.Count

        'ตุ้ย Mark ไว้ก่อน 
        '##1## Get Lab No.
        'ถ้าไม่ใช่ IQC เข้าเงื่อนไขแรก แต่ถ้าเข้า IQC ต้องเข้าเงื่อนไข 2
        For i = 0 To j - 1

            If LineArr(i).IndexOf("DATE") > -1 Then
                LineArr(i) = LineArr(i).Replace("DATE", String.Empty)
                LineArr(i) = LineArr(i).Trim
                rxlyteDate = LineArr(i)
            End If

            If LineArr(i).IndexOf("TIME") > -1 Then
                LineArr(i) = LineArr(i).Replace("TIME", String.Empty)
                LineArr(i) = LineArr(i).Trim
                rxlyteTime = LineArr(i)
            End If

            If LineArr(i).IndexOf("K") > -1 Then
                rxlyteQc = LineArr(i - 1).Replace("Test", String.Empty).Trim
            End If

            If LineArr(i).IndexOf("Sample No.:") > -1 Then
                LineArr(i) = LineArr(i).Replace("Sample No.:", String.Empty)
                LineArr(i) = LineArr(i).Trim
                If LineArr(i).IndexOf("(") > -1 Then
                    LineArr(i) = LineArr(i).Substring(0, LineArr(i).IndexOf("("))
                End If
                SpecimenId = LineArr(i)
                Continue For
            End If

        Next

        analyzerDate = DateTime.Parse(rxlyteDate & " " & rxlyteTime)

        'ตุ้ย Mark ไว้ก่อนกรณีไม่มี Specimen Id.
        If SpecimenId = "" Then
            SpecimenId = rxlyteQc
        End If

        'Debug.WriteLine("Rxlyte Lab No. " & SpecimenId)

        'Tui Add
        'ดึงค่า Result ทั้งหมดที่ Order ไว้
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        If dtTestResult.Rows.Count <= 0 Then
            Return True
            Exit Function
        End If

        '##2## Get test result by test code
        For i = 0 To j - 1
            'Example : K  = 4.13 mmol/L  ( 3.60- 5.00)
            ItemStr = LineArr(i)

            Debug.WriteLine("ItemStr =" & ItemStr)

            'Example : TCO2:<5.0 mmol/L  ( 3.60- 5.00)
            '-->       TCO2<5.0 mmol/L  ( 3.60- 5.00)
            If ItemStr.IndexOf(":") <> -1 Then
                ItemStr = ItemStr.Replace(":", "")
            End If

            If ItemStr.IndexOf("=*") <> -1 Then
                'K   =* 4.32 mmol/L ( 3.50- 5.50)
                ItemStr = ItemStr.Replace("=*", "#@")
            ElseIf ItemStr.IndexOf("=") <> -1 Then
                'K  #= 4.13 mmol/L  ( 3.60- 5.00)
                ItemStr = ItemStr.Replace("=", "#=")
            ElseIf ItemStr.IndexOf(">") <> -1 Then
                'K  #> 4.13 mmol/L  ( 3.60- 5.00)
                ItemStr = ItemStr.Replace(">", "#>")
            ElseIf ItemStr.IndexOf("<") <> -1 Then
                'K  #< 4.13 mmol/L  ( 3.60- 5.00)
                ItemStr = ItemStr.Replace("<", "#<")
            ElseIf ItemStr.IndexOf("---") <> -1 Then
                'K  #--- 4.13 mmol/L  ( 3.60- 5.00)
                ItemStr = ItemStr.Replace("---", "#---")
            ElseIf ItemStr.IndexOf("*") <> -1 Then
                'K   * 4.32 mmol/L ( 3.50- 5.50)
                ItemStr = ItemStr.Replace("*", "#*")
            End If

            ItemArr = ItemStr.Split("#")
            'ItemArr(0) --> K
            'ItemArr(1) --> = 4.13 mmol/L  ( 3.60- 5.00)

            If ItemArr.Count < 2 Then
                Continue For
            End If

            ItemArr(0) = ItemArr(0).Trim

            Debug.WriteLine("ItemArr(0) = " & ItemArr(0))

            dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(0) & "'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = get_result_rxlyte(ItemArr(1))
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            ElseIf dvTestResult.Count > 1 Then
                Dim ii As Integer = 0

                For ii = 0 To dvTestResult.Count - 1
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(ii)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(ii)("alias_id")
                    row("result_value") = get_result_rxlyte(ItemArr(1))
                    row("order_id") = dvTestResult.Item(ii)("order_id")
                    row("analyzer") = dvTestResult.Item(ii)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                    row("result_date") = analyzerDate
                    row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(ii)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)

                Next


            End If
        Next

        UpdateTestResultValue(dtTestResultToUpdate, False)
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        UpdateFormula(SpecimenId)

        Return True

    End Function
    Public Function ExtractResult_ISE6000(ByVal DeviceStr As String) As Boolean

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim Sex As String = ""
        Dim ResultStatus As Boolean = False

        'If DeviceStr.IndexOf("AG") = -1 Then
        '    Return False
        'End If

        Dim i, j As Integer
        Dim labTemp As Int32 = 0
        Dim running As Int32 = 0

        'IQC
        Dim rxlyteDate As String = String.Empty
        Dim rxlyteTime As String = String.Empty
        Dim rxlyteQc As String = String.Empty

        LineArr = DeviceStr.Split(Chr(10))
        j = LineArr.Count

        For i = 0 To j - 1

            ItemArr = System.Text.RegularExpressions.Regex.Replace(LineArr(i), "\s+", " ").Split(" ")


            If Not IsNumeric(ItemArr(1)) Then
                Continue For
            End If

            SpecimenId = Integer.Parse(ItemArr(1)).ToString()
            If SpecimenId = "" Then
                Continue For
            End If

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
            Dim dvTestResult As New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            If dtTestResult.Rows.Count <= 0 Then
                Continue For
            End If
            dvTestResult.RowFilter = "analyzer_ref_cd = 'K'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = ItemArr(3)
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If

            dvTestResult.RowFilter = "analyzer_ref_cd = 'Na'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = ItemArr(4)
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If

            dvTestResult.RowFilter = "analyzer_ref_cd = 'Cl'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = ItemArr(5)
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If

            dvTestResult.RowFilter = "analyzer_ref_cd = 'TCO2'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = ItemArr(8)
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If


            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()
            UpdateFormula(SpecimenId)


        Next

        Return True

    End Function
    Private Function get_result_rxlyte(ByVal str As String) As String 'For RxLyte Only!!!!

        Dim str_num As String

        str = str.Trim

        'ลบเครื่องหมาย = ออก แต่ถ้าเป็นเครื่องหมาย > หรือ < จะเก็บไว้
        '= 4.13 mmol/L  ( 3.60- 5.00) -->  4.13 mmol/L  ( 3.60- 5.00)
        '> 4.13 mmol/L  ( 3.60- 5.00) --> ไม่เปลี่ยนแปลง
        '= --- mmol/L  ( 3.60- 5.00)  -->  --- mmol/L  ( 3.60- 5.00)
        str = str.Replace("=", String.Empty)

        'ลบ Space ออก
        '  4.13 mmol/L  ( 3.60- 5.00) --> 4.13mmol/L(3.60-5.00)
        '> 4.13 mmol/L  ( 3.60- 5.00) --> >4.13mmol/L(3.60-5.00)
        '   --- mmol/L  ( 3.60- 5.00) --> ---mmol/L(3.60-5.00)
        str = str.Replace(" ", String.Empty)

        str = str.Trim

        For i As Integer = 0 To str.Length - 1
            If IsNumeric(str.Substring(i, 1)) Or str.Substring(i, 1) = "." Or str.Substring(i, 1) = ">" Or str.Substring(i, 1) = "<" Or str.Substring(i, 1) = "-" Or str.Substring(i, 1) = "*" Or str.Substring(i, 1) = "@" Then
                If str_num = "" Then
                    str_num = str.Substring(i, 1)
                Else
                    str_num = str_num & str.Substring(i, 1)
                End If
            Else
                Exit For
            End If
            str_num = str_num.Replace("@", "=*")
        Next

        'str_num = 4.13 , >4.13 , ---

        Return str_num

    End Function
    Public Function ExtractResult_Chem(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If





        Dim i, j As Integer
        'Tui Declare
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            If analyzerModel = "COBAS_C311" Then

                ' AppSetting.WriteErrorLog(comPort, "Error", "a " & i)

                If ItemStrOld.Length > 0 And LineArr(i).IndexOf(Chr(2)) <> -1 Then
                    ItemStr = ItemStrOld + LineArr(i).Substring(3) ' LineArr(i).Replace(Chr(2), "")
                    ItemStrOld = ""
                Else
                    ItemStr = LineArr(i)
                End If

                ItemArr = ItemStr.Split("|")

                If ItemArr(0).Substring(0, 1) = "R" Then
                    If ItemStr.IndexOf(Chr(23)) <> -1 Then
                        ItemStrOld = ItemStr.Substring(0, ItemStr.IndexOf(Chr(23)))
                        Continue For
                    End If
                End If
            Else
                ItemStr = LineArr(i)
                ItemArr = ItemStr.Split("|")
            End If


            If ItemArr.Count <= 0 Then
                Continue For
            End If
            ' Dim strTest As String = ItemArr(0).Substring(2, 1)
            ' If ItemArr(0).IndexOf("2P") <> -1 Then

            If ItemArr(0).IndexOf("P") <> -1 And analyzerModel <> "XS" Then
                Dim SpecimenIdStr As String
                'Get Lab No
                'Tui Add Imola
                If analyzerModel.Equals("Imola") Or analyzerModel.Equals("RXMODENA") Or analyzerModel.Equals("Modena") Or analyzerModel.Equals("XN350") Then 'Golf add DX300, RXMODENA
                    SpecimenIdStr = LineArr(i + 2)
                    'End Tui Add Emola
                ElseIf analyzerModel.Equals("DX300") Or analyzerModel.Equals("COBAS_C311") Then

                    SpecimenIdStr = LineArr(i + 1)

                Else

                    SpecimenIdStr = LineArr(i + 1)
                End If

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                'Golf create c ase1
                Select Case analyzerModel
                    Case "Imola", "RXMODENA", "Modena", "COBAS_C311"
                        SpecimenId = SpecimenIdArr(2).Trim
                    Case "DX300", "SUZUKA"
                        'Debug.WriteLine("Between 6 and 8, inclusive")
                        SpecimenId = SpecimenIdArr(2).Substring(0, SpecimenIdArr(2).IndexOf(Chr(94))).Trim
                    Case "XN350"
                        SpecimenId = SpecimenIdArr(3).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Trim
                    Case Else
                        SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("E", "").Trim.Replace("M", "").Trim
                        SpecimenId = SpecimenIdArr(2).Trim
                End Select

                If SpecimenId.Trim <> "" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                    If analyzerModel = "Imola" Then
                        If My.Settings.TrackError Then
                            log.Info(comPort & "GetResultCodeDilution")
                        End If
                        dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                        If My.Settings.TrackError Then
                            log.Info(comPort & "GetResultCodeDilution2")
                        End If
                    Else
                        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    End If

                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                End If

            End If

            If ItemArr(0).IndexOf("2P") <> -1 And analyzerModel = "XS" Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 2)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(3).Trim.Replace(" ", "").Replace("^", "").Replace("E", "").Trim.Replace("M", "").Trim

                If SpecimenId.Trim <> "" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("O") <> -1 Then

                If analyzerModel = "XS-Serie" Then
                    'Example 4O|1||^^    11070200022^M|.......
                    itemArrOrd = ItemArr(3).Split("^") 'Split ที่ ^ และใช้ตำแหน่งที่ 2
                    itemArrOrd(2) = itemArrOrd(2).Trim
                    SpecimenId = itemArrOrd(2).Trim
                Else
                    'DX300,Imola,Daytona,Pentra60,Liason... ใช้ 10 หลักแรก
                    If analyzerModel.Equals("Daytona") Or analyzerModel.Equals("DaytonaPlus") Then
                        ItemArr(2) = ItemArr(2).Trim
                        SpecimenId = ItemArr(2).Trim

                        'Golf 2017-05-04

                        If SpecimenId.Trim <> "" And dtTestResult.Rows.Count <= 0 Then
                            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                            dvTestResult = New DataView(dtTestResult)
                            dtTestResultToUpdate = dtTestResult.Clone()
                        End If

                        'Golf 2017-05-04
                    End If

                End If
                'Golf 2016-04-18
                If analyzerModel = "DX300" Then
                    If ItemArr(0).IndexOf("O") <> -1 Then
                        Dim SpecimenIdStr As String
                        SpecimenIdStr = LineArr(i)
                        Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                        SpecimenId = SpecimenIdArr(2).Substring(0, SpecimenIdArr(2).IndexOf(Chr(94))).Trim
                        If SpecimenId.Trim <> "" Then

                            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                            dvTestResult = New DataView(dtTestResult)
                            dtTestResultToUpdate = dtTestResult.Clone()

                        End If
                    End If
                ElseIf analyzerModel = "SUZUKA" Then
                    If ItemArr(0).IndexOf("O") <> -1 Then
                        Dim SpecimenIdStr As String
                        SpecimenIdStr = LineArr(i)
                        Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                        SpecimenId = SpecimenIdArr(2).Substring(0, SpecimenIdArr(2).IndexOf(Chr(94))).Trim
                        If SpecimenId.Trim <> "" Then

                            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy

                            dvTestResult = New DataView(dtTestResult)
                            dtTestResultToUpdate = dtTestResult.Clone()

                        End If
                    End If


                End If




                '<<<< ******* Result Line ******** >>>>
                ' Result
            ElseIf ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result

                If analyzerModel = "Pentra60" Then
                    'Example Pentra60  =  ^^^HGB^717-9
                    itemArrRes = ItemArr(2).Split("^")
                    ItemArr(2) = itemArrRes(3).Trim

                ElseIf analyzerModel = "XS-Serie" Then
                    'Example ^^^^NEUT%^1
                    itemArrRes = ItemArr(2).Split("^")
                    ItemArr(2) = itemArrRes(4).Trim

                    'Tui Add COBAS_e411, Feb 26, 2014
                ElseIf analyzerModel = "COBAS_e411" Then
                    'Example Result Item ^^^900^^0
                    itemArrRes = ItemArr(2).Split("^")
                    ItemArr(2) = itemArrRes(3).Trim

                    'Example Result Value -1^0.533
                    itemArrRes = ItemArr(3).Split("^")
                    If itemArrRes.Count > 1 Then
                        ItemArr(3) = itemArrRes(1).Trim
                    End If
                    'End COBAS_e411, Feb 26, 2014
                ElseIf analyzerModel = "COBAS_C311" Then
                    'Example Result Item ^^^900^^0
                    itemArrRes = ItemArr(2).Split("^")
                    ItemArr(2) = itemArrRes(3).Substring(0, itemArrRes(3).IndexOf("/")).Trim

                    Dilution = itemArrRes(3).Substring(itemArrRes(3).Length - 1)
                    If Dilution.IndexOf("/") = -1 Then
                        Dilution_flag = "Y"
                    End If
                    '|
                    'Example Result Value -1^0.533
                    itemArrRes = ItemArr(3).Split("^")
                    If itemArrRes.Count > 1 Then
                        ItemArr(3) = itemArrRes(1).Trim
                    End If
                    'End COBAS_e411, Feb 26, 2014
                ElseIf analyzerModel = "XS" Then
                    itemArrRes = ItemArr(2).Split("^")
                    ItemArr(2) = itemArrRes(4).Trim

                Else
                    'DX300,Imola,Daytona,Liason.....
                    ItemArr(2) = ItemArr(2).Replace(Chr(94), "") 'Chr(94) = ^
                    ItemArr(2) = ItemArr(2).Trim 'Mapping Testcode
                    If analyzerModel = "DX300" Then
                        ItemArr(3) = ItemArr(3).Replace("***", " ")
                    ElseIf analyzerModel = "SUZUKA" Then
                        ItemArr(3) = ItemArr(3).Replace("***", " ")
                    End If

                End If

                'Tui Add 2015-09-15 Get IQC Date
                If analyzerModel = "COBAS_e411" OrElse analyzerModel = "Imola" OrElse analyzerModel = "Daytona" OrElse analyzerModel = "DaytonaPlus" OrElse analyzerModel = "RXMODENA" OrElse analyzerModel = "Modena" OrElse analyzerModel = "XN350" OrElse analyzerModel = "COBAS_C311" OrElse analyzerModel = "Liaison" OrElse analyzerModel = "XS" OrElse analyzerModel = "Maglumi" Then

                    Try
                        If analyzerModel = "Liaison" OrElse analyzerModel = "XS" OrElse analyzerModel = "Maglumi" Then
                            If ItemArr.Count >= 13 Then
                                Dim dt As String = Mid(ItemArr(12), 1, 14)
                                dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                                resultDate = DateTime.Parse(dt)
                            End If
                        Else
                            If ItemArr.Count = 13 Then
                                Dim dt As String = Mid(ItemArr(12), 1, 14)
                                dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                                resultDate = DateTime.Parse(dt)
                            End If
                        End If
                        If ItemArr.Count = 13 Then
                            Dim dt As String = Mid(ItemArr(12), 1, 14)
                            dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                            resultDate = DateTime.Parse(dt)
                        End If
                    Catch ex As Exception
                        AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                        log.Error(ex.Message)
                    End Try

                End If
                'End Tui Add 2015-09-15 Get IQC Date
                Debug.WriteLine(SpecimenId)

                If analyzerModel = "COBAS_C311" And Dilution_flag = "Y" Then
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                End If

                'Golf 2017-08-17 start
                If analyzerModel = "Imola" And ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If analyzerModel = "Imola" And ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    If (ItemArr(6).ToString = "<" OrElse ItemArr(6).ToString = ">") And (analyzerModel = "Liaison" Or analyzerModel = "Maglumi" Or analyzerModel = "XS" Or analyzerModel = "Imola" Or analyzerModel = "RXMODENA" Or analyzerModel = "Modena") Then

                        row("result_value") = ItemArr(6).ToString() & ItemArr(3).ToString()

                    Else
                        row("result_value") = ItemArr(3).ToString()
                    End If

                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = resultDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    row("sample_type") = dvTestResult.Item(0)("sample_type")
                    dtTestResultToUpdate.Rows.Add(row)
                End If


            End If

            'Imola
            If analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Or analyzerModel = "COBAS_C311" Or analyzerModel = "XS" Then ' Or analyzerModel = "Imola" Then
                GoTo endLine
            End If
            If analyzerModel = "Imola" Or analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Or analyzerModel = "Liaison" Or analyzerModel = "RXMODENA" Or analyzerModel = "Modena" Or analyzerModel = "Maglumi" Then
                'ถ้า ผลของ DX300 และ SUZUKA เบิ้ล ต้องย้ายลงมาไว้หลัง Loop และเปิด if ด้านบนด้วย
                If ItemArr(0).IndexOf("R") <> -1 Then
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If

                    If analyzerModel = "Imola" Then
                        dtTestResultToUpdate.Clear()

                    End If

                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
                    'If My.Settings.TrackError Then
                    '    log.Info(comPort & "UpdateTestResultValue2")
                    'End If
                End If
            Else
                'Daytona
                If ItemArr(0).IndexOf("P") <> -1 Or ItemArr(0).IndexOf("L") <> -1 Then

                    UpdateTestResultValue(dtTestResultToUpdate, False)

                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
                    'ElseIf analyzerModel = "DX300" Then
                    '    UpdateTestResultValue(dtTestResultToUpdate, False)
                    '    dtTestResult.Dispose()
                    '    dtTestResultToUpdate.Dispose()
                End If
            End If
endLine:
        Next
        If analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Or analyzerModel = "COBAS_C311" Or analyzerModel = "XS" Then
            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()
        End If

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True







    End Function

    Public Function ExtractResult_Architecti1000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 And result_type = "F" Then
                    'Dim row As DataRow = dtTestResultToUpdate.NewRow
                    'row("order_skey") = dvTestResult.Item(0)("order_skey")
                    'row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    'row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    'row("alias_id") = dvTestResult.Item(0)("alias_id")
                    'row("result_value") = ItemArr(3).ToString()
                    'row("order_id") = dvTestResult.Item(0)("order_id")
                    'row("analyzer") = dvTestResult.Item(0)("analyzer")
                    'row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    'row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    'row("result_date") = resultDate
                    'row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    'row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    'row("rerun") = dvTestResult.Item(0)("rerun")
                    'row("sample_type") = dvTestResult.Item(0)("sample_type")
                    'dtTestResultToUpdate.Rows.Add(row)
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If



            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_C4000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 And result_type = "F" Then
                    'Dim row As DataRow = dtTestResultToUpdate.NewRow
                    'row("order_skey") = dvTestResult.Item(0)("order_skey")
                    'row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    'row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    'row("alias_id") = dvTestResult.Item(0)("alias_id")
                    'row("result_value") = ItemArr(3).ToString()
                    'row("order_id") = dvTestResult.Item(0)("order_id")
                    'row("analyzer") = dvTestResult.Item(0)("analyzer")
                    'row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    'row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    'row("result_date") = resultDate
                    'row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    'row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    'row("rerun") = dvTestResult.Item(0)("rerun")
                    'row("sample_type") = dvTestResult.Item(0)("sample_type")
                    'dtTestResultToUpdate.Rows.Add(row)
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Function ExtractResult_C8000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 And result_type = "F" Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Function ExtractResult_Echo(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        Dim assay_code As String = ""
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            'If ItemArr(0).IndexOf("P") <> -1 Then
            '    Dim SpecimenIdStr As String
            '    SpecimenIdStr = LineArr(i + 1)

            '    Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

            '    SpecimenId = SpecimenIdArr(2).Trim
            '    If SpecimenId.IndexOf("^") <> -1 Then
            '        Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
            '        SpecimenId = SpecimenIdArr2(0).Trim
            '    End If

            '    If SpecimenId.Trim <> "" Then
            '        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
            '        dvTestResult = New DataView(dtTestResult)
            '        dtTestResultToUpdate = dtTestResult.Clone()
            '    End If

            'End If
            If ItemArr(0).IndexOf("O") <> -1 Then
                SpecimenId = ItemArr(2)
                assay_code = ItemArr(4).Replace("^", "")
                'If SpecimenId.Trim <> "" Then
                '    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                '    dvTestResult = New DataView(dtTestResult)
                '    dtTestResultToUpdate = dtTestResult.Clone()
                'End If
            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                If SpecimenId.Trim <> "" Then
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

                ItemArr(3) = ItemArr(3).Trim 'Result

                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    ItemArr(2) = assay_codeArr(3).Trim
                End If
                Dim result_value As String = ""
                If ItemArr(3).IndexOf("^") <> -1 Then
                    Dim result_value_assay() As String = ItemArr(3).Split("^")
                    result_value = result_value_assay(1).Trim
                Else
                    result_value = ItemArr(3).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try

                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        If ii = 0 Then
                            Dim row As DataRow = dtTestResultToUpdate.NewRow
                            row("order_skey") = dvTestResult.Item(ii)("order_skey")
                            row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                            row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                            row("alias_id") = dvTestResult.Item(ii)("alias_id")
                            row("result_value") = result_value
                            row("order_id") = dvTestResult.Item(ii)("order_id")
                            row("analyzer") = dvTestResult.Item(ii)("analyzer")
                            row("analyzer_ref_cd") = assay_code & "|" & dvTestResult.Item(ii)("analyzer_ref_cd")
                            row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                            row("result_date") = resultDate
                            row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                            row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                            row("rerun") = dvTestResult.Item(ii)("rerun")
                            row("sample_type") = dvTestResult.Item(ii)("sample_type")
                            dtTestResultToUpdate.Rows.Add(row)
                        End If

                    Next
                End If
                If dtTestResultToUpdate.Rows.Count > 0 Then
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 And dtTestResultToUpdate.Rows.Count > 0 Then
                dtTestResultToUpdate.Clear()
                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        UpdateFormula(SpecimenId)
        Return True
    End Function

    Public Function ExtractResult_IQ200(ByVal DeviceStr As String) As Boolean 'Golf IQ200 2020-10-15
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        Dim result_type As String = "F"



        Dim dvTestResult As New DataView


        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""

        Dim docXDocument As XDocument = XDocument.Parse(DeviceStr)
        Dim root As XElement

        root = docXDocument.Root
        Dim resultValue As String
        If SpecimenId.Trim = "" Then
            SpecimenId = root.Attribute("ID").Value()
        End If
        If SpecimenId.Trim <> "" Then
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
            dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()
        End If
        If dtTestResult.Rows.Count() <= 0 Then
            GoTo endLoop
        End If
        For Each level1 In root.Elements()
            'Do Something with level1 attr'

            For Each level2 In level1.Elements()
                'Do Something with level2 attr'
                '<AR Key="GLU" SN="GLU" LN="Glucose" AF="1" NR="+- ">+4</AR>

                If level2.Name.LocalName = "AR" Then
                    resultValue = level2.Value
                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & level2.Attribute("Key").Value() & "'"
                    If dvTestResult.Count >= 1 And result_type = "F" Then
                        For ii = 0 To dvTestResult.Count - 1
                            Dim row As DataRow = dtTestResultToUpdate.NewRow
                            row("order_skey") = dvTestResult.Item(ii)("order_skey")
                            row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                            row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                            row("alias_id") = dvTestResult.Item(ii)("alias_id")
                            row("result_value") = resultValue
                            row("order_id") = dvTestResult.Item(ii)("order_id")
                            row("analyzer") = dvTestResult.Item(ii)("analyzer")
                            row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                            row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                            row("result_date") = resultDate
                            row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                            row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                            row("rerun") = dvTestResult.Item(ii)("rerun")
                            row("sample_type") = dvTestResult.Item(ii)("sample_type")
                            dtTestResultToUpdate.Rows.Add(row)
                        Next
                    End If
                End If

                For Each level3 In level2.Elements()
                    'Do Something with level3 attr'
                Next
            Next
        Next
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()

        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
endLoop:

        Return True
    End Function 'Golf IQ200 2020-10-15
    Public Function ExtractResult_PATHFAST(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = ItemArr(2)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("^")

                SpecimenId = SpecimenIdArr(0).Trim
                'If SpecimenId.IndexOf("^") <> -1 Then
                '    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                '    SpecimenId = SpecimenIdArr2(0).Trim
                'End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                Dim result_value As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End
                result_value = ""
                If ItemArr(3).IndexOf("^") <> -1 Then
                    Dim assay_result_value() As String = ItemArr(3).Split("^")
                    result_value = assay_result_value(0).Trim
                Else
                    result_value = ItemArr(3).ToString()
                End If
                If ItemArr(6).IndexOf("<") <> -1 Then
                    result_value = "<" + result_value
                End If
                If ItemArr(6).IndexOf(">") <> -1 Then
                    result_value = ">" + result_value
                End If

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then


                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = result_value
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_HbA1c(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemArrI() As String
        Dim ItemArrR() As String
        Dim ItemArrQ() As String
        Dim ItemArrX() As String
        Dim SpecimenId As String = ""
        Dim resultCode As String = ""
        Dim resultValue As String = ""
        Dim analyzerDateString As String
        Dim resultValueString As String

        Dim ds As New DataSet()

        If DeviceStr.IndexOf("<Q>") > 0 Then
            'IQC <<----------------------------------------------
            ' Dim aaa As String = DeviceStr.Substring(DeviceStr.IndexOf("<Q>")).Replace("</M></TRANSMIT>", "</TRANSMIT>")
            'LineArr = DeviceStr.Substring(DeviceStr.IndexOf("<Q>")).Replace("</M></TRANSMIT>", vbNewLine).Split(vbNewLine)
            Dim LineStr As String = DeviceStr.Substring(DeviceStr.IndexOf("<Q>")).Replace("</M></TRANSMIT>", "").Replace("</Q>", "</Q>" + Chr(10)).Replace("</X>", "</X>" + Chr(10))


            LineArr = LineStr.Split(Chr(10))


            For Each itemString As String In LineArr

                If itemString.IndexOf("<Q>") <> -1 Then
                    ItemArrQ = itemString.ToString.Split("|")
                    SpecimenId = ItemArrQ(2)
                ElseIf itemString.IndexOf("<X>") <> -1 Then
                    ItemArrX = itemString.ToString.Split("|")
                    analyzerDate = ItemArrX(1).Substring(0, ItemArrX(1).IndexOf("HbA1c"))
                    resultCode = "HbA1c"
                    resultValue = ItemArrX(2).Substring(0, ItemArrX(2).IndexOf("NO"))

                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    Dim dvTestResult As New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
                    If dvTestResult.Count >= 1 Then
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(0)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(0)("alias_id")
                        row("result_value") = resultValue
                        row("order_id") = dvTestResult.Item(0)("order_id")
                        row("analyzer") = dvTestResult.Item(0)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                        row("result_date") = analyzerDate
                        row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(0)("rerun")
                        dtTestResultToUpdate.Rows.Add(row)
                    End If

                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
                    UpdateFormula(SpecimenId)


                End If
            Next

        Else

            LineArr = DeviceStr.Replace("</TRANSMIT>", "</TRANSMIT>" + vbNewLine).Split(vbNewLine)

            For Each itemString As String In LineArr
                Using stringReader As New StringReader(itemString)
                    ds = New DataSet()
                    ds.ReadXml(stringReader)
                End Using
                Dim dt As DataTable = ds.Tables(0)
                ItemArrI = dt(0).Item("I").ToString.Split("|")
                analyzerDate = DateTime.Parse(ItemArrI(1))
                'analyzerDate = Now() ' Golf 2019-07-26
                SpecimenId = ItemArrI(4) 'Golf add 2017-02-28

                ItemArrR = dt(0).Item("R").ToString.Split("|")
                resultCode = "HbA1c" 'Golf Comments 
                resultValue = ItemArrR(1).ToString.Substring(ItemArrR(2).ToString.IndexOf("=") + 2)

                dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                Dim dvTestResult As New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
                If dvTestResult.Count = 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = resultValue
                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = analyzerDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)
                End If

                UpdateTestResultValue(dtTestResultToUpdate, False)
                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
                UpdateFormula(SpecimenId)

            Next
        End If



        Return True

    End Function

    Public Function ExtractResult_XN350(ByVal DeviceStr As String) As Boolean

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False

        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If

        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))
        j = LineArr.Count

        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")

            If ItemArr.Count <= 0 Then
                Continue For
            End If
            If ItemArr(0).IndexOf("P") <> -1 Then

                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 2)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")


                SpecimenId = SpecimenIdArr(3).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Trim


                If SpecimenId.Trim <> "" Then

                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                End If

            End If
            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result

                'DX300,Imola,Daytona,Liason.....
                ItemArr(2) = ItemArr(2).Replace(Chr(94), "") 'Chr(94) = ^
                ItemArr(2) = ItemArr(2).Trim 'Mapping Testcode


                Try
                    If ItemArr.Count = 13 Then
                        Dim dt As String = Mid(ItemArr(12), 1, 14)
                        dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                        resultDate = DateTime.Parse(dt)
                    End If
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try

                Debug.WriteLine(SpecimenId)
                Dim testOrder As String = ""
                testOrder = Mid(ItemArr(2), 1, ItemArr(2).Length - 1)
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & testOrder & "'"
                If dvTestResult.Count >= 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = ItemArr(3).ToString()
                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = resultDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)
                End If

            End If


            'ถ้า ผลของ DX300 และ SUZUKA เบิ้ล ต้องย้ายลงมาไว้หลัง Loop และเปิด if ด้านบนด้วย
            If ItemArr(0).IndexOf("R") <> -1 Then
                UpdateTestResultValue(dtTestResultToUpdate, False)
                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next


        'Tui add 2015-09-22  auto update formula
        UpdateFormula(SpecimenId)

        Return True

    End Function

    Public Function ExtractResult_ABXPentraXL80(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If





        Dim i, j As Integer
        'Tui Declare
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If
            ' Dim strTest As String = ItemArr(0).Substring(2, 1)
            ' If ItemArr(0).IndexOf("2P") <> -1 Then


            If ItemArr(0).IndexOf("3O") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim

                If SpecimenId.Trim <> "" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result

                itemArrRes = ItemArr(2).Split("^")
                ItemArr(2) = itemArrRes(4).Trim

                'Tui Add 2015-09-15 Get IQC Date
                If analyzerModel = "ABXPentraXL80" Then

                    Try
                        If analyzerModel = "ABXPentraXL80" Then
                            If ItemArr.Count >= 13 Then
                                Dim dt As String = Mid(ItemArr(12), 1, 14)
                                dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                                resultDate = DateTime.Parse(dt)
                            End If
                        Else
                            If ItemArr.Count = 13 Then
                                Dim dt As String = Mid(ItemArr(12), 1, 14)
                                dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                                resultDate = DateTime.Parse(dt)
                            End If
                        End If
                        If ItemArr.Count = 13 Then
                            Dim dt As String = Mid(ItemArr(12), 1, 14)
                            dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                            resultDate = DateTime.Parse(dt)
                        End If
                    Catch ex As Exception
                        AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                        log.Error(ex.Message)
                    End Try

                End If


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = ItemArr(3).ToString() 'ItemArr(6).ToString() & ItemArr(3).ToString()

                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = resultDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    row("sample_type") = dvTestResult.Item(0)("sample_type")
                    dtTestResultToUpdate.Rows.Add(row)
                End If


            End If

        Next
        If analyzerModel = "ABXPentraXL80" Then
            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()
        End If

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True







    End Function
    Public Function ExtractResult_Architecti2000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 And result_type = "F" Then
                    'Dim row As DataRow = dtTestResultToUpdate.NewRow
                    'row("order_skey") = dvTestResult.Item(0)("order_skey")
                    'row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    'row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    'row("alias_id") = dvTestResult.Item(0)("alias_id")
                    'row("result_value") = ItemArr(3).ToString()
                    'row("order_id") = dvTestResult.Item(0)("order_id")
                    'row("analyzer") = dvTestResult.Item(0)("analyzer")
                    'row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    'row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    'row("result_date") = resultDate
                    'row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    'row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    'row("rerun") = dvTestResult.Item(0)("rerun")
                    'row("sample_type") = dvTestResult.Item(0)("sample_type")
                    'dtTestResultToUpdate.Rows.Add(row)
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If



            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_DxC700AU(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim check2E() As String
        Dim checkEE As String = ""
        'Dim ItemArr() As String
        'Dim itemArrOrd() As String
        'Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""

        Dim Sex As String = ""
        'Dim firstPat As Boolean = True
        'Dim ResultStatus As Boolean = False
        Dim textTrim As String = ""
        Dim remain As String = ""

        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(3)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "DCOBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(2))
        End If

        j = LineArr.Count

        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)

            If ItemStr.IndexOf(Chr(6)) <> -1 Then
                Continue For
            End If

            If ItemStr.Length <= 1 Or ItemStr = Chr(32) Then
                Continue For
            End If

            If ItemStr.IndexOf("DB") <> -1 Then
                Continue For
            End If
            Dim CheckHU As String = ""
            If ItemStr.IndexOf("U") <> -1 Then
                Dim CheckU As String = ItemStr.Replace("U", Chr(32))
                CheckHU = "U"
                ItemStr = CheckU
            End If

            Dim CheckHB As String = ""
            If ItemStr.IndexOf("W") <> -1 Then

                Dim CheckW As String = ItemStr.Replace("W", Chr(32))
                CheckHB = "C"
                ItemStr = CheckW

            End If
            'Check Emergency
            check2E = ItemStr.Split("E")
            If check2E.Count > 2 Then
                SpecimenId = check2E(1).Trim()
                SpecimenId = SpecimenId.Substring(3)
                checkEE = "EE"
            Else
                SpecimenId = ItemStr
                Dim checkP As Integer = InStr(1, SpecimenId, "E", 1) - 1
                SpecimenId = SpecimenId.Substring(0, checkP)
                SpecimenId = SpecimenId.Replace(Chr(32), "^")
            End If
            If SpecimenId.IndexOf("^^^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^^^", "^")
            End If
            If SpecimenId.IndexOf("^^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^^", "^")
            End If
            If SpecimenId.IndexOf("^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^", "^")
            End If

            Dim subArr() As String = SpecimenId.Split("^")

            'Check 12 Digit
            If subArr.Count = 6 Then
                SpecimenId = subArr(4).Trim()
            End If

            'check 7 Digit
            If subArr.Count = 4 Then
                SpecimenId = subArr(2)
                SpecimenId = SpecimenId.Substring(4).Trim()
            End If

            'Check QC Test
            If subArr.Count = 3 Then
                SpecimenId = subArr(1)
                SpecimenId = SpecimenId.Substring(4).Trim()
            End If
            If SpecimenId = "" Then
                Continue For
            End If

            If SpecimenId.Length >= 3 Then


                dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()



                ItemStr = ItemStr.Replace("Hr", "n")
                ItemStr = ItemStr.Replace("r", "n")
                ItemStr = ItemStr.Replace("Tp", "n")
                ItemStr = ItemStr.Replace("T", "n")
                ItemStr = ItemStr.Replace("G", "n")
                ItemStr = ItemStr.Replace("-", "n")

                'For Each row As DataRow In dtTestResult.Rows

                '    Dim textC As String = row.Item("analyzer_ref_cd")
                '    Dim textC2 As String = "Hn" & textC
                '    If ItemStr.IndexOf(row.Item("analyzer_ref_cd")) <> -1 Then
                '        ItemStr = ItemStr.Replace(textC, textC2)
                '    End If
                'Next

                If checkEE = "EE" Then
                    textTrim = check2E(2).Trim()
                    ItemStr = textTrim
                End If

                ItemStr = ItemStr.Replace("n", "r")
                ItemStr = ItemStr.Replace(Chr(32), "n")



                If ItemStr.IndexOf("nnnnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnnnn", "Z")
                End If
                If ItemStr.IndexOf("nnnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnnn", "Z")
                End If
                If ItemStr.IndexOf("nnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnn", "Z")
                End If
                If ItemStr.IndexOf("nnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnn", "Z")
                End If
                If ItemStr.IndexOf("nnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnn", "Z")
                End If
                If ItemStr.IndexOf("nnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnn", "Z")
                End If
                If ItemStr.IndexOf("nn") <> -1 Then
                    ItemStr = ItemStr.Replace("nn", "|")
                End If
                If ItemStr.IndexOf("Ln") <> -1 Then
                    ItemStr = ItemStr.Replace("Ln", "|")
                End If
                If ItemStr.IndexOf("Hn") <> -1 Then
                    ItemStr = ItemStr.Replace("Hn", "|")
                End If

                'ItemStr = ItemStr.Replace("r", "n")

                Dim CheckE As String = ItemStr.Replace(Chr(32), "") 'Chr(32) = Space

                textTrim = CheckE.Substring(InStr(1, CheckE, "E", 1))

                textTrim = textTrim.Replace("Z", "").Trim

                Dim i2 As Long

                Dim result2() As String = textTrim.Split("|")

                Dim assay_code As String = ""
                Dim assay_rusult As String = ""

                For i2 = 0 To UBound(result2)

                    Dim trimArr As String = result2(i2)

                    If result2(i2).IndexOf("r") <> -1 Then
                        trimArr = result2(i2).Replace("r", "")
                    End If

                    If result2(i2).IndexOf("H") <> -1 Then
                        trimArr = result2(i2).Replace("H", "")
                    End If

                    If result2(i2).IndexOf("HH") <> -1 Then
                        trimArr = result2(i2).Replace("HH", "")
                    End If

                    If result2(i2).IndexOf("h") <> -1 Then
                        trimArr = result2(i2).Replace("h", "")
                    End If

                    If result2(i2).IndexOf("b") <> -1 Then
                        trimArr = result2(i2).Replace("b", "")
                    End If

                    If result2(i2).IndexOf("L") <> -1 Then
                        trimArr = result2(i2).Replace("L", "")
                    End If
                    If result2(i2).IndexOf(Chr(3)) <> -1 Then
                        trimArr = result2(i2).Replace(Chr(3), "")
                    End If

                    Dim resultT = trimArr.Trim()

                    If resultT.Length < 2 Then
                        Continue For

                    End If

                    assay_code = resultT.Substring(0, 3) 'Code

                    assay_rusult = resultT.Substring(3).Replace("C", "").Trim 'Result

                    If assay_rusult.IndexOf("?") <> -1 Then
                        Continue For
                    End If


                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & assay_code & "'"
                    If dvTestResult.Count >= 1 Then

                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(0)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(0)("alias_id")
                        row("result_value") = assay_rusult.ToString()
                        row("order_id") = dvTestResult.Item(0)("order_id")
                        row("analyzer") = dvTestResult.Item(0)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(0)("rerun")
                        row("sample_type") = dvTestResult.Item(0)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)

                    End If

                Next

            End If

        Next
endLine:

        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue1")
        End If
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)

        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If

        Return True

    End Function

    Public Function ReturnOrderDetail(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck6 ")
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        'Dim Running As Integer = 1
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If

            If analyzerModel = "DX300" Then
                SpecimenId = SpecimenId.Substring(0, SpecimenId.IndexOf(Chr(94))).Trim 'Golf 
            ElseIf analyzerModel = "Liaison" Or analyzerModel = "Maglumi" Then
                Return ReturnOrderDetail_LIASON(SpecimenId, Running) 'Golf 2016-10-06

                ''Tui Add COBAS_e411, Feb 24, 2014
            ElseIf analyzerModel = "COBAS_e411" Then
                Return ReturnOrderDetail_COBAS_e411(SpecimenId)
                ''End Tui Add COBAS_e411, Feb 24, 2014

            ElseIf analyzerModel = "SUZUKA" Then
                SpecimenId = SpecimenId.Substring(0, SpecimenId.IndexOf(Chr(94))).Trim 'Golf 
            ElseIf analyzerModel = "COBAS_C311" Then
                Return ReturnOrderDetail_COBAS_C311(SpecimenId)
            ElseIf analyzerModel = "Maglumi" Then
                Return ReturnOrderDetail_Maglumi(SpecimenId)
            ElseIf analyzerModel = "HA8180V" Then 'Ploy Add 2020.11.03
                Dim LineArr() As String
                LineArr = SpecimenId.Replace("-", "^").Split("^")
                SpecimenId = LineArr(0)
                'ElseIf analyzerModel = "Stago" Then 'Ploy Add 2020.12.01
                '    Dim LineArr() As String
                '    LineArr = SpecimenId.Replace("-", "^").Split("^")
                '    SpecimenId = LineArr(0)
            ElseIf analyzerModel = "Stago" Then 'Ploy Add 2020.12.02
                SpecimenId = SpecimenId.Replace("^", "")
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Select Case analyzerModel
                'Golf 2016-04-11
                Case "DX300"
                    i = 0
                    StrArr.Add(Chr(5))
                    ''Header
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1

                    'Patient
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "Test"
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)


                    For Each row As DataRow In dtTestResult.Rows

                        Running += 1


                        i += 1
                        Str &= SpecimenIdReturn & "||"

                        Str = ""

                        Str = Chr(2) & CStr(Running) & "O|" & i & "|"

                        Str &= SpecimenIdReturn & "||"
                        'Str &= "160037143^1^1^3^N" & "||"


                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "|"

                        Str &= Chr(13) & Chr(3)
                        Str &= GetCheckSumValue(Str.Remove(0, 1))

                        Str &= Chr(13) & Chr(10)
                        StrArr.Add(Str)



                        ''EndLine

                    Next
                    'EndLine
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))

                    'Golf 2016-04-11
                Case "SUZUKA"
                    i = 0
                    StrArr.Add(Chr(5))
                    ''Header
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1

                    'Patient
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "Test"
                    Str &= "||||||||||||^"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)


                    For Each row As DataRow In dtTestResult.Rows
                        Running += 1


                        i += 1
                        Str &= SpecimenIdReturn & "||"

                        Str = ""

                        Str = Chr(2) & CStr(Running) & "O|" & i & "|"

                        Str &= SpecimenIdReturn & "||"

                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "|"

                        ''GolfTest
                        Str &= "R|"
                        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                        Str &= "||||||||" & "|1||||||||||"
                        'End If 

                        Str &= Chr(13) & Chr(3)
                        Str &= GetCheckSumValue(Str.Remove(0, 1))

                        Str &= Chr(13) & Chr(10)
                        StrArr.Add(Str)

                    Next
                    'EndLine
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))

                    'Golf 2016-04-11 Start
                Case "Modena", "RXMODENA"
                    ' ''Solution 1 Start
                    'StrArr.Add(Chr(5)) 'ENQ enquiry 'Golf comments 2017-03-16
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    ' Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""

                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "\"
                    Next

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then
                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                Case "Daytona"
                    ' ''Solution 1 Start
                    'StrArr.Add(Chr(5)) 'ENQ enquiry 'Golf comments 2017-03-16
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    ' Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""

                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "\"
                    Next

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then
                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                Case "DaytonaPlus"
                    ' ''Solution 1 Start
                    'StrArr.Add(Chr(5)) 'ENQ enquiry 'Golf comments 2017-03-16
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    ' Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""

                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "\"
                    Next
                    Str &= "||||||||||||05" 'golf 2018-11-16

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then
                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                Case "Dimension"
                    Dim dvTestOrder As New DataView(dtTestResult)
                    dvTestOrder.RowFilter = "analyzer_ref_cd not in ('AGAP','LDL','GLOB','%ISAT','%MB','%FPSA','A/G','BN/CR','FTI','IBIL','MA/CR','MBRI','OSMO','UIBC','RISK')"

                    If dvTestOrder.Count > 0 Then
                        Str = Chr(2) & "D" & Chr(28) & "0" & Chr(28) & "0" & Chr(28) & "A" & Chr(28)
                        Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                        Str &= Chr(28)
                        Str &= SpecimenId
                        Str &= Chr(28)
                        Str &= "1" 'row.Item("sample_type")
                        Str &= Chr(28) & "" & Chr(28) & "0" & Chr(28) & "1" & Chr(28) & "**" & Chr(28)
                        Str &= "1" & Chr(28)

                        Dim LYTE_COUNT As Integer = 0
                        Dim Test_COUNT As Integer = 0
                        Dim LYTE_Flag As Boolean = False

                        For Each rowView2 As DataRowView In dvTestOrder
                            Dim row2 As DataRow = rowView2.Row
                            If row2.Item("analyzer_ref_cd") = "NA" Then
                                LYTE_COUNT = LYTE_COUNT + 1
                            End If
                            If row2.Item("analyzer_ref_cd") = "Cl" Then
                                LYTE_COUNT = LYTE_COUNT + 1
                            End If
                            If row2.Item("analyzer_ref_cd") = "K" Then
                                LYTE_COUNT = LYTE_COUNT + 1
                            End If
                        Next

                        If LYTE_COUNT = 3 Then
                            Test_COUNT = dvTestOrder.Count - 2
                        Else
                            Test_COUNT = dvTestOrder.Count
                        End If
                        Str &= CStr(Test_COUNT)
                        Str &= Chr(28)
                        For Each rowView As DataRowView In dvTestOrder
                            Dim row As DataRow = rowView.Row
                            If LYTE_COUNT = 3 Then
                                'LYTE
                                If row.Item("analyzer_ref_cd") = "NA" OrElse row.Item("analyzer_ref_cd") = "Cl" OrElse row.Item("analyzer_ref_cd") = "K" Then
                                    If Not (LYTE_Flag) Then
                                        Str &= "LYTE"
                                        Str &= Chr(28)
                                        LYTE_Flag = True
                                    End If
                                Else
                                    Str &= row.Item("analyzer_ref_cd")
                                    Str &= Chr(28)
                                End If
                            Else
                                Str &= row.Item("analyzer_ref_cd")
                                Str &= Chr(28)
                            End If
                        Next
                        Str &= GetCheckSumValue(Str.Remove(0, 1))
                        Str &= Chr(3)
                        StrArr.Add(Str)
                    Else
                        Str = Chr(2) & "N" & Chr(28)
                        Str &= GetCheckSumValue(Str.Remove(0, 1))
                        Str &= Chr(3)
                        StrArr.Add(Str)
                    End If
                Case "Imola"
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                            Str &= SepTest & row.Item("analyzer_ref_cd")
                            Str &= "\"
                            'Golf Start 2017-01-08 
                        Else
                            If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                                Str &= SepTest & "61"
                                Str &= "\"
                            End If
                            'Golf End 2017-01-08 
                        End If
                        'Str &= SepTest & row.Item("analyzer_ref_cd")
                        'Str &= "\"
                    Next

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    'Golf 2017-10-30
                    'If Str.IndexOf("^^^891") <> -1 Or Str.IndexOf("^^^99") <> -1 Or Str.IndexOf("^^^20") <> -1 Or Str.IndexOf("^^^21") <> -1 Then
                    '    Str &= "|R||||||A||||4|||||||"
                    'Else
                    '    Str &= "|R||||||A||||1|||||||"
                    'End If

                    'Str &= "|R||||||A|||||01||||||"
                    'Str &= "||||||||||||01"
                    'Str &= "|R||||||A||||||01|||||"
                    'Str &= "||||||||||||02"


                    'Str &= "|||||||||||||02"
                    'Golf 2017-10-30
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                Case "ABXPentraXL80"
                    StrArr.Add(Chr(5))
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    'For Each row As DataRow In dtTestResult.Rows
                    '    i += 1
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    'Next
                    Str &= SepTest & "CBC|R||||||A"

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))
                Case "HA8180V" 'Ploy Add 2020.11.02 
                    'H|$^&| | |HA-8180V^10900001^V01.00 | | | | | | | | | 200203111010<CR>
                    'P|1<CR>
                    'O|1|0123456789012345--^0001^01||^^^HbA1c|R| | | | | | | | | |H<CR>
                    'L|1|N<CR>
                    Str = Chr(2) & "1H|$^&|||HA-8180V^11710002^V01.46    |||||||||"
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") '& CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    'O|1|0123456789012345--^0001^01||^^^HbA1c|R| | | | | | | | | |H<CR>
                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenIdReturn & "||^^^HbA1c|R| | | | | | | | | |H"

                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                Case "Stago" 'Ploy Add 2020.12.01 
                    'ต้องส่ง
                    'H.P.O.L
                    'Ex
                    '[Stx]1H|\^&|||STA_Satlellite_Max^0007M10052^1|||||||P|1|20090504141209[CR][Etx]1E[CR]
                    '[Stx]2P|1[CR][Etx]3F[CR]
                    '[Stx]3O|1|300105^^0902-0901^3|205||R||||||||||||||||||||F[CR][Etx]46[CR]
                    '[Stx]6L|1|N[CR][Etx]09[CR]

                    'Ploy Add 2020.12.24

                    'StrArr.Add(Chr(5))
                    'Str &= Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|" 'Chr(2) = STX (start-of-text)
                    ''Str = Chr(5) & Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|" 'Chr(2) = STX (start-of-text)
                    'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    'Str &= Chr(13) & Chr(3)
                    'Str &= GetCheckSumValue(Str.Remove(0, 1))
                    'Str &= Chr(13) & Chr(10)
                    'StrArr.Add(Str)

                    StrArr.Add(Chr(5))
                    Str &= Chr(2) & "1H|\^&||||||||||P||" 'Chr(2) = STX (start-of-text)
                    ' Str &= Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|" 'Chr(2) = STX (start-of-text)
                    'Str = Chr(5) & Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    'Old Ploy Comment 2021.01.14
                    'Running += 1
                    'Str = ""
                    'Str = Chr(2) & CStr(Running) & "P|1|"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "|||"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Birthday").ToString)
                    'Str &= "|"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Sex").ToString)
                    'Str &= "||||||"
                    'Str &= Chr(13) & Chr(3)
                    'Str &= GetCheckSumValue(Str.Remove(0, 1))
                    'Str &= Chr(13) & Chr(10)
                    'StrArr.Add(Str)

                    'Ploy Add 2021.01.14
                    'Running += 1
                    'Str = ""
                    'Str = Chr(2) & CStr(Running) & "P|1|||"
                    ''Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    ''Str &= "||"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
                    'Str &= "^"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString)
                    'Str &= "|"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Birthday").ToString)
                    'Str &= "|"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Sex").ToString)
                    'Str &= "||||||"
                    'Str &= Chr(13) & Chr(3)
                    'Str &= GetCheckSumValue(Str.Remove(0, 1))
                    'Str &= Chr(13) & Chr(10)
                    'StrArr.Add(Str)

                    'Ploy Add 2021.01.18
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1||"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||"
                    'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "||"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
                    Str &= "^"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString)
                    Str &= "|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Birthday").ToString)
                    Str &= "|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Sex").ToString)
                    Str &= "|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
                    Str &= "^"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString)
                    Str &= "||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    ''Test 1111 เหมือน 10 แต่ตัด F กับ 64 ออก
                    'Running += 1
                    'Str = ""
                    'Str = Chr(2) & CStr(Running) & "O|1|"
                    'Str &= SpecimenId & "^^^||"
                    'For Each row As DataRow In dtTestResult.Rows
                    '    If row.Item("analyzer_ref_cd") = "1F" Then
                    '        Str &= "^^^1\"
                    '    ElseIf row.Item("analyzer_ref_cd") = "2F" Then
                    '        Str &= "^^^2\"
                    '    ElseIf row.Item("analyzer_ref_cd") = "3F" Then
                    '        Str &= "^^^3\"
                    '    ElseIf row.Item("analyzer_ref_cd") = "4F" Then
                    '        Str &= "^^^4\"
                    '    End If
                    'Next
                    ''ตัด \ ตัวสุดท้ายออก
                    'Str = Str.Substring(0, Str.Length - vbCrLf.Length + 1)

                    'Str &= "^^^^^^|R||N|O"
                    'Str &= Chr(13) & Chr(3)
                    'Str &= GetCheckSumValue(Str.Remove(0, 1))
                    'Str &= Chr(13) & Chr(10)
                    'StrArr.Add(Str)
                    ''End 1111

                    'Test 1121 เปลี่ยน LIS Code เป็น Internal Code
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "^^^||"
                    For Each row As DataRow In dtTestResult.Rows
                        If row.Item("analyzer_ref_cd") = "1F" Then
                            Str &= "^^^CPTO1\"
                        ElseIf row.Item("analyzer_ref_cd") = "2F" Then
                            Str &= "^^^CPTO3\"
                        ElseIf row.Item("analyzer_ref_cd") = "3F" Then
                            Str &= "^^^Co15\"
                        ElseIf row.Item("analyzer_ref_cd") = "4F" Then
                            Str &= "^^^Co49\"
                        End If
                    Next
                    'ตัด \ ตัวสุดท้ายออก
                    Str = Str.Substring(0, Str.Length - vbCrLf.Length + 1)

                    Str &= "^F^64^^^^|R||N|O"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    'End 1121

                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))
                Case Else


                    StrArr.Add(Chr(5))
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "\"
                    Next

                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))


            End Select
            ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.13 " & Str)

            '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.14 ")
            Debug.Write("Header Str----------")
            Debug.Write(String.Join("", StrArr.ToArray()))
            Debug.Write("Footer Str----------")
            Select Case analyzerModel
                Case "Imola"
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to Imola.")
                    AppSetting.WriteErrorLog(comPort, "Imola", String.Join("", StrArr.ToArray()))
                Case "DX300" 'Golf Create
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to DX300.")
                    AppSetting.WriteErrorLog(comPort, "DX300", String.Join("", StrArr.ToArray()))
                Case "SUZUKA" 'Golf Create
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to SUZUKA.")
                    AppSetting.WriteErrorLog(comPort, "SUZUKA", String.Join("", StrArr.ToArray()))
                Case "Liaison" 'Golf Create
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to Liaison.")
                    AppSetting.WriteErrorLog(comPort, "Liaison", String.Join("", StrArr.ToArray()))
                Case "RXMODENA", "Modena"
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to RXMODENA.")
                    AppSetting.WriteErrorLog(comPort, "RXMODENA", String.Join("", StrArr.ToArray()))

                Case "Daytona"
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to Daytona.")
                    AppSetting.WriteErrorLog(comPort, "Daytona", String.Join("", StrArr.ToArray()))
                Case "DaytonaPlus"
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to DaytonaPlus.")
                    AppSetting.WriteErrorLog(comPort, "DaytonaPlus", String.Join("", StrArr.ToArray()))
                Case "Dimension"
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to Dimension.")
                    AppSetting.WriteErrorLog(comPort, "Dimension", String.Join("", StrArr.ToArray()))
                Case "HA8180V" 'Ploy Add 2020.11.03
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to HA8180V.")
                    AppSetting.WriteErrorLog(comPort, "HA8180V", String.Join("", StrArr.ToArray()))
                Case "Stago" 'Ploy Add 2020.12.16
                    AppSetting.WriteErrorLog(comPort, "Send", "Send to Stago.")
                    AppSetting.WriteErrorLog(comPort, "Stago", String.Join("", StrArr.ToArray()))
            End Select

            ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.15 ")

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.16 ")
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.17 " & (String.Join("", StrArr.ToArray())))
        Return StrArr

    End Function

    Public Function ReturnOrderDetailArchitecti1000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                    'If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    '    'Golf Start 2017-01-08 
                    'Else
                    '    If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                    '        Str &= SepTest & "61"
                    '        Str &= "\"
                    '    End If
                    '    'Golf End 2017-01-08 
                    'End If
                Next

                If Str.Substring(Str.Length - 1, 1) = "\" Then
                    If dtTestResult.Rows.Count > 1 Then
                        Str = Str.Substring(0, Str.Length - 1)
                        Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                        Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                        Dim StrList As String() = TestItemList.Split("\")
                        Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                        Dim TestItemListCheckDup2 As String = ""
                        For Each xx As String In TestItemListCheckDup
                            TestItemListCheckDup2 &= xx
                            TestItemListCheckDup2 &= "\"
                        Next

                        Str = HeadList & TestItemListCheckDup2
                    End If
                    Str = Str.Substring(0, Str.Length - 1)
                End If

                'New
                Str &= "|||||||"
                'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0") 'Collection Date and Time
                'Str &= "||||"
                Str &= "A" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "|"
                Str &= "Hep"
                Str &= "|"
                Str &= "lipemic"
                Str &= "||"
                Str &= "serum" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Q" 'Report Types O: Order, Q: Order in response to a Query Request.
                'New

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Architect i1000.")
            AppSetting.WriteErrorLog(comPort, "Architecti1000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function

    Public Function ReturnOrderDetailC4000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                    'If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    '    'Golf Start 2017-01-08 
                    'Else
                    '    If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                    '        Str &= SepTest & "61"
                    '        Str &= "\"
                    '    End If
                    '    'Golf End 2017-01-08 
                    'End If
                Next

                If Str.Substring(Str.Length - 1, 1) = "\" Then
                    If dtTestResult.Rows.Count > 1 Then
                        Str = Str.Substring(0, Str.Length - 1)
                        Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                        Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                        Dim StrList As String() = TestItemList.Split("\")
                        Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                        Dim TestItemListCheckDup2 As String = ""
                        For Each xx As String In TestItemListCheckDup
                            TestItemListCheckDup2 &= xx
                            TestItemListCheckDup2 &= "\"
                        Next

                        Str = HeadList & TestItemListCheckDup2
                    End If
                    Str = Str.Substring(0, Str.Length - 1)
                End If

                'New
                Str &= "|||||||"
                'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0") 'Collection Date and Time
                'Str &= "||||"
                Str &= "A" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "|"
                Str &= "Hep"
                Str &= "|"
                Str &= "lipemic"
                Str &= "||"
                Str &= "serum" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Q" 'Report Types O: Order, Q: Order in response to a Query Request.
                'New

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Architect C4000.")
            AppSetting.WriteErrorLog(comPort, "C4000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ReturnOrderDetailC8000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                    'If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    '    'Golf Start 2017-01-08 
                    'Else
                    '    If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                    '        Str &= SepTest & "61"
                    '        Str &= "\"
                    '    End If
                    '    'Golf End 2017-01-08 
                    'End If
                Next

                If Str.Substring(Str.Length - 1, 1) = "\" Then
                    If dtTestResult.Rows.Count > 1 Then
                        Str = Str.Substring(0, Str.Length - 1)
                        Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                        Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                        Dim StrList As String() = TestItemList.Split("\")
                        Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                        Dim TestItemListCheckDup2 As String = ""
                        For Each xx As String In TestItemListCheckDup
                            TestItemListCheckDup2 &= xx
                            TestItemListCheckDup2 &= "\"
                        Next

                        Str = HeadList & TestItemListCheckDup2
                    End If
                    Str = Str.Substring(0, Str.Length - 1)
                End If

                'New
                Str &= "|||||||"
                'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0") 'Collection Date and Time
                'Str &= "||||"
                Str &= "A" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "|"
                Str &= "Hep"
                Str &= "|"
                Str &= "lipemic"
                Str &= "||"
                Str &= "serum" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Q" 'Report Types O: Order, Q: Order in response to a Query Request.
                'New

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to C8000.")
            AppSetting.WriteErrorLog(comPort, "C8000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function

    Public Function ReturnOrderDetailEcho(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SampleType As String = ""
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "\" & "^&|||LIS|||||Echo|||LIS2-A2|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                If dtTestResult.Rows.Count > 0 Then
                    Select Case dtTestResult.Rows(0).Item("sample_type")
                        Case 10
                            SampleType = "ABOD Full Screen"
                        Case 11
                            SampleType = "Crossmatch"
                        Case 12
                            SampleType = "Donor Short"
                        Case 13
                            SampleType = "ABOD Group"
                        Case 14
                            SampleType = "Screen"
                        Case 15
                            SampleType = "ABORH"
                        Case 16
                            SampleType = "ABO"
                        Case 17
                            SampleType = "Ab screening"
                        Case 18
                            SampleType = "Weak D"
                    End Select

                End If
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                'Str &= "|"
                'Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""

                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                Str &= SepTest & SampleType
                Str &= "|R||||||N||||Blood^Patient"

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Echo.")
            AppSetting.WriteErrorLog(comPort, "Echo", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ReturnOrderDetailPATHFAST(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()
            'Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|" & "@^" & "\|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count > 0 Then
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str = Chr(2) & CStr(Running) & "O|" & i & "|"
                    Str &= SpecimenId & "||"
                    'For Each row As DataRow In dtTestResult.Rows
                    'i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                    'Next

                    If Str.Substring(Str.Length - 1, 1) = "\" Then
                        If dtTestResult.Rows.Count > 1 Then
                            Str = Str.Substring(0, Str.Length - 1)
                            Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                            Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                            Dim StrList As String() = TestItemList.Split("\")
                            Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                            Dim TestItemListCheckDup2 As String = ""
                            For Each xx As String In TestItemListCheckDup
                                TestItemListCheckDup2 &= xx
                                TestItemListCheckDup2 &= "\"
                            Next

                            Str = HeadList & TestItemListCheckDup2
                        End If
                        Str = Str.Substring(0, Str.Length - 1)
                    End If

                    'New
                    Str &= "|||||||"
                    Str &= "" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                    Str &= "|"
                    Str &= ""
                    Str &= "|"
                    Str &= ""
                    Str &= "||"
                    Str &= "" 'Specimen Type
                    Str &= "||||||||||"
                    Str &= "O" 'Report Types O: Order, Q: Order in response to a Query Request.
                    Str &= "|||||" 'golf 2020-07-14
                    'New

                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                Next
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to PATHFAST")
            AppSetting.WriteErrorLog(comPort, "Imola", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ReturnOrderDetail_COBAS_e411(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        'ReturnOrderDetailDimension
        Dim StrArr As New ArrayList
        'Dim Running As Integer = 1
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim hn As String = "N/A"
        Dim dtPatient As New DataTable
        Dim SpecimenIdArr() As String
        Dim labSeq As String = ""

        Try

            If SpecimenId.Length > 0 Then
                SpecimenIdArr = SpecimenId.Split("^")
                SpecimenId = SpecimenIdArr(1)
                labSeq = SpecimenIdArr(2)
            End If

            Debug.WriteLine("COBAS_e411 SpecimenId = " & SpecimenId)
            Debug.WriteLine("COBAS_e411 Seq = " & labSeq)

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            StrArr.Add(Chr(5))

            ''Header
            ''Example
            Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3)

            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1

            'Patient
            Str = ""
            Str = Chr(2) & CStr(Running) & "P|1|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
            Str &= "||||||||||||"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1

            'Order & TestCode
            Str = ""
            i = 0
            Str = Chr(2) & CStr(Running) & "O|1|"
            Str &= SpecimenId & "|" & labSeq & "|"    'Tui Comment 2013-11-05
            'Str &= "20130101-000" & "||" 'Tui fix 2013-11-04-05

            'For Each dr_testlist In dt_testlist
            '    i += 1
            '    dv_Analyzer.RowFilter = "TestCode = '" & dr_testlist.TestCode & "'"

            '    If dv_Analyzer.Count = 1 Then
            '        Str &= SepTest & dv_Analyzer.Item(0)("MappingName")
            '        'Str &= "¥"
            '        Str &= "\"
            '    End If
            'Next

            Debug.WriteLine("row count = " & dtTestResult.Rows.Count)

            For Each row As DataRow In dtTestResult.Rows
                i += 1
                'dv_Analyzer.RowFilter = "TestCode = '" & dr_testlist.TestCode & "'"

                'If dv_Analyzer.Count = 1 Then
                'Str &= SepTest & dv_Analyzer.Item(0)("MappingName")
                Str &= SepTest & row.Item("analyzer_ref_cd") & "^0"
                'Str &= "¥"
                Str &= "\"
                'End If
            Next

            If Str.Substring(Str.Length - 1, 1) = "\" Then
                Str = Str.Substring(0, Str.Length - 1)
            End If

            Str &= "|R||||||N||||||||||||||"

            If dtTestResult.Rows.Count > 0 Then
                Str &= "Q"
            Else
                Str &= "Z"
            End If

            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1

            'EndLine
            Str = ""
            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            StrArr.Add(Chr(4))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_COBAS_e411: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_COBAS_e411: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return StrArr
    End Function

    Public Function ReturnOrderDetailArchitecti2000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                    'If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    '    'Golf Start 2017-01-08 
                    'Else
                    '    If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                    '        Str &= SepTest & "61"
                    '        Str &= "\"
                    '    End If
                    '    'Golf End 2017-01-08 
                    'End If
                Next

                If Str.Substring(Str.Length - 1, 1) = "\" Then
                    If dtTestResult.Rows.Count > 1 Then
                        Str = Str.Substring(0, Str.Length - 1)
                        Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                        Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                        Dim StrList As String() = TestItemList.Split("\")
                        Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                        Dim TestItemListCheckDup2 As String = ""
                        For Each xx As String In TestItemListCheckDup
                            TestItemListCheckDup2 &= xx
                            TestItemListCheckDup2 &= "\"
                        Next

                        Str = HeadList & TestItemListCheckDup2
                    End If
                    Str = Str.Substring(0, Str.Length - 1)
                End If

                'New
                Str &= "|||||||"
                'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0") 'Collection Date and Time
                'Str &= "||||"
                Str &= "A" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "|"
                Str &= "Hep"
                Str &= "|"
                Str &= "lipemic"
                Str &= "||"
                Str &= "serum" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Q" 'Report Types O: Order, Q: Order in response to a Query Request.
                'New

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Architect i2000.")
            AppSetting.WriteErrorLog(comPort, "Architecti2000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function

    Public Function ReturnHeader(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Dim Str As String
        Dim StrArr As New ArrayList

        StrArr.Add(Chr(5))

        'Header
        Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||Liaison||||"
        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
        Str &= Chr(13) & Chr(3)

        Str &= GetCheckSumValue(Str.Remove(0, 1))

        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)

        Running = set_i(Running)

        Return StrArr

    End Function
    Public Function ReturnEnd(ByVal SpecimenId As String, Optional ByRef running As Integer = 1) As ArrayList

        Dim Str As String
        Dim StrArr As New ArrayList

        'EndLine
        Str = ""
        Str = Chr(2) & CStr(running) & "L|1|N" & Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)

        StrArr.Add(Chr(4))

        Return StrArr

    End Function
    Public Function ReturnOrderDetail_LIASON(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList

        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim SepTest2 As String = "^1:22"
        Dim EndStr As String = Chr(13)
        Dim i As Integer

        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
        Dim dvTestResult = New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()




        ''''3 Start Chapter 5 Page 5-30 - 5-32 ################################################
        'Header
        StrArr.Add(Chr(5))
        'Str = Chr(2) & "1H|" & "\" & "^&|"
        Str = Chr(2) & "1H|" & "\" & "^&|||LaborEDV|||||Liaison|||1|"
        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
        Str &= Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)

        'Patient
        Running += 1
        If Running >= 8 Then
            Running = 0
        End If
        'Patient
        Str = ""
        Str = Chr(2) & CStr(Running) & "P|1"
        Str &= Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)

        For Each row As DataRow In dtTestResult.Rows
            Running += 1
            If Running >= 8 Then
                Running = 0
            End If
            i += 1
            Str = ""
            '   Str = Chr(2) & CStr(Running) & "O|" & i & "|"
            Str = Chr(2) & CStr(Running) & "O|" & "1" & "|"
            Str &= SpecimenId & "||"
            Str &= SepTest & row.Item("analyzer_ref_cd")
            Str &= "|N|"
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0")
            Str &= "|||||||||S||||||||||O|"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
        Next

        'EndLine
        Running += 1
        If Running = 8 Then
            Running = 0
        End If
        Str = ""
        Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)
        StrArr.Add(Chr(4))
        ''''3 Chapter 5 Page 5-30 - 5-32 ################################################



        AppSetting.WriteErrorLog(comPort, "Send", "Send to LIAISON1")
        AppSetting.WriteErrorLog(comPort, "LIAISON", String.Join("", StrArr.ToArray()))
        Return StrArr

    End Function


    Public Function ReturnOrderDetail_LIASON9(ByVal SpecimenIdArray As ArrayList, Optional ByRef Running As Integer = 1) As ArrayList

        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim SepTest2 As String = "^1:22"
        Dim EndStr As String = Chr(13)
        Dim i As Integer

        'Header
        StrArr.Add(Chr(5))
        'Str = Chr(2) & "1H|" & "\" & "^&|"
        Str = Chr(2) & "1H|" & "\" & "^&|||LaborEDV|||||Liaison|||1|"
        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
        Str &= Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)


        For Each SpecimenId As String In SpecimenIdArray

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Running += 1
            If Running = 8 Then
                Running = 0
            End If
            'Patient
            Str = ""
            Str = Chr(2) & CStr(Running) & "P|1"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)


            For Each row As DataRow In dtTestResult.Rows
                Running += 1
                If Running = 8 Then
                    Running = 0
                End If
                i += 1
                Str = ""
                '   Str = Chr(2) & CStr(Running) & "O|" & i & "|"
                Str = Chr(2) & CStr(Running) & "O|" & "1" & "|"
                Str &= SpecimenId & "||"
                Str &= SepTest & row.Item("analyzer_ref_cd")
                Str &= "|N|"
                Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0")
                Str &= "|||||||||S||||||||||O|"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
            Next
        Next

        'EndLine
        Running += 1
        If Running = 8 Then
            Running = 0
        End If
        Str = ""
        Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(13) & Chr(10)
        StrArr.Add(Str)
        StrArr.Add(Chr(4))


        AppSetting.WriteErrorLog(comPort, "Send", "Send to LIAISON9")
        AppSetting.WriteErrorLog(comPort, "LIAISON", String.Join("", StrArr.ToArray()))
        Return StrArr

    End Function

    Public Function ReturnOrderDetailIQ200(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList 'Golf IQ200 2020-10-15
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & "<UNKID>" & SpecimenId & "</UNKID>"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
            Else
                Str = Chr(2) & "<?xml version= ""1.0""?><SI BF=""URN"" WO=""run"">" & SpecimenId
                Str &= "<DEMOG LName=""" & dtPatient.Rows(0).Item("lastname").ToString & """" & " FNAME=""" & dtPatient.Rows(0).Item("lastname").ToString & """"
                Str &= " MNAME=""" & dtPatient.Rows(0).Item("middle_name").ToString & """"
                Str &= " DOB=""" & dtPatient.Rows(0).Item("birthday").ToString.Substring(0, 4) & "-" & dtPatient.Rows(0).Item("birthday").ToString.Substring(4, 2) & "-" & dtPatient.Rows(0).Item("birthday").ToString.Substring(6, 2) & """"
                Str &= " Loc="""""
                Str &= " RecNum="""""
                Str &= " Gender=""" & dtPatient.Rows(0).Item("sex").ToString & """"
                Str &= "/><SI>"
                '<DEMOG LName="DOE" FNAME="JOHN" MName="" DOB="1933-04-05" Loc="TESTING" RecNum="257314" Gender="M"/></SI>
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
            End If


            AppSetting.WriteErrorLog(comPort, "Send", "Send to IQ200")
            AppSetting.WriteErrorLog(comPort, "IQ200", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function 'Golf IQ200 2020-10-15
    Private Function set_i(ByRef i As Integer) As Integer

        i = i + 1

        If i > 7 Then
            i = 0
        End If

        Return i

    End Function

    Public Class T
        Public Const ENQ As Byte = 5
        Public Const ACK As Byte = 6
        Public Const NAK As Byte = 21
        Public Const EOT As Byte = 4
        Public Const ETX As Byte = 3
        Public Const ETB As Byte = 23
        Public Const STX As Byte = 2
        Public Const NEWLINE As Byte = 10
        Public Shared ACK_BUFF As Byte() = {ACK}
        Public Shared ENQ_BUFF As Byte() = {ENQ}
        Public Shared NAK_BUFF As Byte() = {NAK}
        Public Shared EOT_BUFF As Byte() = {EOT}
    End Class

    ''' <summary>
    ''' Reads checksum of an ASTM frame, calculates characters after STX,
    ''' up to and including the ETX or ETB. Method assumes the frame contains an ETX or ETB.
    ''' </summary>
    ''' <param name="frame">frame of ASTM data to evaluate</param>
    ''' <returns>string containing checksum</returns>
    ''' 



    'Public Shared Function GetCheckSumValue(ByVal frame As String) As String
    '    Dim checksum As String = "00"

    '    Dim sumOfChars As Integer = 0
    '    Dim complete As Boolean = False

    '    Try
    '        'take each byte in the string and add the values
    '        For idx As Integer = 0 To frame.Length - 1
    '            Dim byteVal As Integer = 0
    '            Dim b() As Byte = Encoding.UTF8.GetBytes(frame(idx))

    '            For i As Integer = 0 To b.Length - 1
    '                byteVal = byteVal + Convert.ToInt32(b(i))
    '            Next

    '            Select Case byteVal
    '                Case 2
    '                    sumOfChars = 0
    '                    Exit Select
    '                Case 3, 23
    '                    sumOfChars += byteVal
    '                    complete = True
    '                    Exit Select
    '                Case Else
    '                    sumOfChars += byteVal
    '                    Exit Select
    '            End Select

    '            If complete Then
    '                Exit For
    '            End If
    '        Next

    '        If sumOfChars > 0 Then
    '            checksum = Convert.ToString(sumOfChars Mod 256, 16).ToUpper()
    '        End If

    '        Return DirectCast(If(checksum.Length = 1, "0" & checksum, checksum), String)
    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function
    Public Function GetCheckSumValue(ByVal frame As String) As String
        Dim checksum As String = "00"

        Dim byteVal As Integer = 0
        Dim sumOfChars As Integer = 0
        Dim complete As Boolean = False

        Try

            'take each byte in the string and add the values
            For idx As Integer = 0 To frame.Length - 1
                'byteVal = Asc(frame(idx))
                byteVal = Convert.ToInt32(frame(idx))
                If byteVal > 255 Then
                    byteVal = 63
                End If
                Select Case byteVal
                    Case T.STX
                        sumOfChars = 0
                        Exit Select
                    Case T.ETX, T.ETB
                        sumOfChars += byteVal
                        complete = True
                        Exit Select
                    Case Else
                        sumOfChars += byteVal
                        Exit Select
                End Select

                If complete Then
                    Exit For
                End If
            Next

            If sumOfChars > 0 Then
                'hex value mod 256 is checksum, return as hex value in upper case
                'checksum = Convert.ToString(sumOfChars Mod 256, 16).ToUpper()
                checksum = Convert.ToString(sumOfChars Mod 256, 16).ToUpper()
            End If

        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetCheckSumValue: " & ex.Message)
            log.Error(ex.Message)
        End Try

        'if checksum is only 1 char then prepend a 0
        Return DirectCast(If(checksum.Length = 1, "0" & checksum, checksum), String)
    End Function

    Function GetResultCode(ByVal SpecimenId As String, ByVal analyzer As String, ByVal analyzerSkey As Int32, ByVal analyzerDate As DateTime, ByVal flagExist As String, ByVal flagScan As String) As DataTable

        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            'OLD
            'parm = SqlProvider.GetParameterArray(5) 'golf 2018-11-19
            'parm(0) = SqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            'parm(1) = SqlProvider.GetParameter("analyzer_model", DbType.String, analyzer, ParameterDirection.Input)
            'parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, analyzerSkey, ParameterDirection.Input)
            'parm(3) = SqlProvider.GetParameter("analyzer_date", DbType.DateTime, analyzerDate, ParameterDirection.Input)
            'parm(4) = SqlProvider.GetParameter("flag_exist", DbType.String, flagExist, ParameterDirection.Input)

            'New
            parm = SqlProvider.GetParameterArray(6) 'golf 2018-11-19
            parm(0) = SqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            parm(1) = SqlProvider.GetParameter("analyzer_model", DbType.String, analyzer, ParameterDirection.Input)
            parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, analyzerSkey, ParameterDirection.Input)
            parm(3) = SqlProvider.GetParameter("analyzer_date", DbType.DateTime, analyzerDate, ParameterDirection.Input)
            parm(4) = SqlProvider.GetParameter("flag_exist", DbType.String, flagExist, ParameterDirection.Input)
            parm(5) = SqlProvider.GetParameter("flagScan", DbType.String, flagScan, ParameterDirection.Input) 'golf 2018-11-19


            SqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_test_result_model", parm)
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.1 ")
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.2 " & SpecimenId)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.21 ")
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            'log.Error(ex.Message)
        Catch ex As Exception
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.22 ")
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            'log.Error(ex.Message)
        Finally
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.23 ")
            SqlProvider.Dispose()
        End Try
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9.3 ")
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9.4 " & dt.Rows.Count.ToString)
        If dt.Rows.Count <= 0 Then
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: No test order for Specimen ID. " & SpecimenId)
        End If

        Return dt

    End Function

    Function GetResultCodeDilution(ByVal SpecimenId As String, ByVal analyzer As String, ByVal analyzerSkey As Int32, ByVal analyzerDate As DateTime, ByVal flagExist As String, ByVal Dilution_flag As String, ByVal ref_cd As String, ByVal flagScan As String) As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            'OLD
            'parm = SqlProvider.GetParameterArray(7) 'golf 2018-11-19
            'parm(0) = SqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            'parm(1) = SqlProvider.GetParameter("analyzer_model", DbType.String, analyzer, ParameterDirection.Input)
            'parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, analyzerSkey, ParameterDirection.Input)
            'parm(3) = SqlProvider.GetParameter("analyzer_date", DbType.DateTime, analyzerDate, ParameterDirection.Input)
            'parm(4) = SqlProvider.GetParameter("flag_exist", DbType.String, flagExist, ParameterDirection.Input)
            'parm(5) = SqlProvider.GetParameter("Dilution_flag", DbType.String, Dilution_flag, ParameterDirection.Input)
            'parm(6) = SqlProvider.GetParameter("ref_cd", DbType.String, ref_cd, ParameterDirection.Input)


            'New
            parm = SqlProvider.GetParameterArray(8) 'golf 2018-11-19
            parm(0) = SqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            parm(1) = SqlProvider.GetParameter("analyzer_model", DbType.String, analyzer, ParameterDirection.Input)
            parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, analyzerSkey, ParameterDirection.Input)
            parm(3) = SqlProvider.GetParameter("analyzer_date", DbType.DateTime, analyzerDate, ParameterDirection.Input)
            parm(4) = SqlProvider.GetParameter("flag_exist", DbType.String, flagExist, ParameterDirection.Input)
            parm(5) = SqlProvider.GetParameter("flagScan", DbType.String, flagScan, ParameterDirection.Input) 'golf 2018-11-19
            parm(6) = SqlProvider.GetParameter("Dilution_flag", DbType.String, Dilution_flag, ParameterDirection.Input)
            parm(7) = SqlProvider.GetParameter("ref_cd", DbType.String, ref_cd, ParameterDirection.Input)



            SqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_test_result_model", parm)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            'log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            'log.Error(ex.Message)
        Finally
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.23 ")
            SqlProvider.Dispose()
        End Try
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9.3 ")
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9.4 " & dt.Rows.Count.ToString)
        If dt.Rows.Count <= 0 Then
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: No test order for Specimen ID. " & SpecimenId)
        End If

        Return dt

    End Function

    Function UpdateTestResultValue(ByVal dtResult As DataTable, ByVal isResend As Boolean) As Int16
        Dim retryCounter As Integer = 1

RETRY:

        Try

            If dtResult.Rows.Count <= 0 Then
                Exit Function
            End If

            Dim fm As New frmMain
            fm = DirectCast(Application.OpenForms("frmMain"), frmMain)

            Dim dtTemp As New DataTable
            dtTemp = dtResult.Clone

            'Golf 2017-11-09
            Dim SqlProviderLocal As SqlDataProvider = Nothing

            ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("DataSource_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.ServerName)
            ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("InitialCatalog_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.DatabaseName)
            ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("SQLUserID_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.SQLUserID)
            ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("SQLPassword_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.SQLPassword)

            Dim bConnect As Boolean = True
            If ValueSoft.Common.Utility.IsIPAddress(My.Settings.Server) Then
                If Not ValueSoft.Common.Utility.CanConnectTelnet1433(My.Settings.Server, 5) Then
                    Throw New Exception(String.Format("Can not connect to Server {0}", My.Settings.Server))
                End If
            End If

            ' SqlProviderLocal = SqlDataProvider.GetInstance(ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName)
            SqlProviderLocal = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
            If Not SqlProviderLocal.DataSourceAvailable Then
                SqlDataProvider.ClearDataProvider(ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName)
                Throw New Exception(SqlProviderLocal.ErrorMsg)
            End If
            'Golf 2017-11-09

            ' SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)


            If My.Settings.TrackError Then
                log.Info(comPort & " in UpdateTestResultValue1")
            End If

            For Each row As DataRow In dtResult.Rows
                If My.Settings.TrackError Then
                    log.Info(comPort & " in UpdateTestResultValue2 Loop" & row("order_skey"))
                End If
                parm = SqlProviderLocal.GetParameterArray(8)
                parm(0) = SqlProviderLocal.GetParameter("order_skey", DbType.Int32, row("order_skey"), ParameterDirection.Input)
                parm(1) = SqlProviderLocal.GetParameter("result_item_skey", DbType.Int32, row("result_item_skey"), ParameterDirection.Input)
                parm(2) = SqlProviderLocal.GetParameter("analyzer_skey", DbType.Int32, row("analyzer_skey"), ParameterDirection.Input)
                parm(3) = SqlProviderLocal.GetParameter("result_value", DbType.String, row("result_value"), ParameterDirection.Input)
                parm(4) = SqlProviderLocal.GetParameter("lot_skey", DbType.Int32, row("lot_skey"), ParameterDirection.Input)
                parm(5) = SqlProviderLocal.GetParameter("result_date", DbType.DateTime, row("result_date"), ParameterDirection.Input)
                parm(6) = SqlProviderLocal.GetParameter("rerun", DbType.String, row("rerun"), ParameterDirection.Input)
                parm(7) = SqlProviderLocal.GetParameter("analyzer_ref_cd", DbType.String, row("analyzer_ref_cd"), ParameterDirection.Input)
                SqlProviderLocal.ExecuteNonQuerySP("sp_lis_interface_lis_lab_result_value_model_ins", parm)

                row("sending_status") = "Sent"
                dtTemp.Rows.Add(row("order_skey").ToString, row("result_item_skey").ToString, row("analyzer_skey").ToString, row("analyzer_ref_cd"), row("alias_id"), row("result_value"), row("specimen_type_id").ToString, row("analyzer").ToString, row("sending_status"), 0, row("result_date"), row("order_id"), row("analyzer_cd").ToString)

                If dtTemp.Rows.Count > 0 And fm IsNot Nothing Then
                    fm.AppendResult(dtTemp)
                End If

                dtTemp.Clear()

            Next

            If My.Settings.TrackError Then
                log.Info(comPort & " in UpdateTestResultValue3")
            End If
            'SqlProvider.CommitTransaction()  'Golf comment
            SqlProviderLocal.Dispose()

        Catch ex As SqlException
            If ex.Number = 1205 Then

                'SqlProvider.RollbackTransaction() 'Golf comment
                AppSetting.WriteErrorLog(comPort, "Error", "Deadlock: " & ex.Message)
                retryCounter += 1

                If retryCounter >= 3 Then

                    SqlProvider.Dispose()
                    Exit Function
                End If

                System.Threading.Thread.Sleep(100)
                GoTo RETRY

            Else

                'SqlProvider.RollbackTransaction() 'Golf comment
                AppSetting.WriteErrorLog(comPort, "Error", "UpdateTestResultValue: " & ex.Message)

            End If
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "UpdateTestResultValue: " & ex.Message)
            ' SqlProvider.RollbackTransaction() 'Golf comment
            log.Error(ex.Message)
        Finally
            SqlProvider.Dispose()
        End Try

    End Function

    Function UpdateFormula(ByVal specimenTypeId As String) As Int16


        'Golf 2017-11-09
        Dim SqlProviderLocal As SqlDataProvider = Nothing

        ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("DataSource_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.ServerName)
        ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("InitialCatalog_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.DatabaseName)
        ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("SQLUserID_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.SQLUserID)
        ValueSoft.CoreControlsLib.CoreCommon.WriteApplicationSetting(String.Format("SQLPassword_{0}_{1}", ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName), ValueSoft.CoreControlsLib.VariableClass.SQLPassword)

        Dim bConnect As Boolean = True
        If ValueSoft.Common.Utility.IsIPAddress(My.Settings.Server) Then
            If Not ValueSoft.Common.Utility.CanConnectTelnet1433(My.Settings.Server, 5) Then
                Throw New Exception(String.Format("Can not connect to Server {0}", My.Settings.Server))
            End If
        End If

        'SqlProviderLocal = New SqlDataProvider(String.Format("server={0};database={1};uid={2};pwd={3}", _ServerName, _DatabaseName, _UserID, Password), 60)
        SqlProviderLocal = New SqlDataProvider(New SqlConnection(conString).ConnectionString)

        'SqlProviderLocal = SqlDataProvider.GetInstance(ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName)

        If Not SqlProviderLocal.DataSourceAvailable Then
            SqlDataProvider.ClearDataProvider(ValueSoft.CoreControlsLib.VariableClass.ServerName, ValueSoft.CoreControlsLib.VariableClass.DatabaseName)
            Throw New Exception(SqlProviderLocal.ErrorMsg)
        End If
        'Golf 2017-11-09




        '  SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)

        Try
            'Golf Comment cancel call WCF 2017-06-26
            'SqlProvider.BeginTransaction()
            'parm = SqlProvider.GetParameterArray(1)
            'parm(0) = SqlProvider.GetParameter("specimen_type_id", DbType.String, specimenTypeId, ParameterDirection.Input)
            'SqlProvider.ExecuteNonQuerySP("sp_lis_interface_update_formula", parm)
            'SqlProvider.CommitTransaction()
            'Golf Comment cancel call WCF 2017-06-26

            'Golf Add call WCF 2017-06-26
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula In")
            End If


            ' WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("Calculate Formula {1} Start", Now(), specimenTypeId))
            '  SqlProvider.BeginTransaction()
            For Each drLabResultItemFormula As System.Data.DataRow In LISFormulaHelper.GetLabResultItemFormulaBySpecimenTypeID(SqlProviderLocal, specimenTypeId).Rows
                Try
                    Application.DoEvents() 'Golf  Application.DoEvents()
                    ' WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("In Loop Calculate Formula {1} {2} Start", Now(), specimenTypeId, drLabResultItemFormula("result_item_id")))

                    Dim FormulaValue As String = LISFormulaHelper.CalculateResultItemFormulaBySpecimenTypeID(SqlProviderLocal, specimenTypeId, drLabResultItemFormula("result_item_id"))
                    '  WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("In Loop after CalculateResultItemFormulaBySpecimenTypeID  Calculate Formula {1} {2} Start", Now(), specimenTypeId, drLabResultItemFormula("result_item_id")))
                    builder = New StringBuilder
                    builder.AppendLine("update lis_lab_result_item")
                    builder.AppendLine("set result_value = @result_value")
                    builder.AppendLine("where order_skey = @order_skey and")
                    builder.AppendLine("result_item_skey = @result_item_skey")

                    parm = SqlProviderLocal.GetParameterArray(3)
                    'parm(0) = SqlProvider.GetParameter("order_skey", DbType.Int32, OrderSkey, ParameterDirection.Input)
                    parm(0) = SqlProviderLocal.GetParameter("order_skey", DbType.Int32, drLabResultItemFormula("order_skey"), ParameterDirection.Input)
                    parm(1) = SqlProviderLocal.GetParameter("result_item_skey", DbType.Int32, drLabResultItemFormula("result_item_skey"), ParameterDirection.Input)
                    parm(2) = SqlProviderLocal.GetParameter("result_value", DbType.String, FormulaValue, ParameterDirection.Input)

                    SqlProviderLocal.ExecuteNonQuery(CommandType.Text, builder.ToString(), parm)


                    parm = SqlProviderLocal.GetParameterArray(2)
                    parm(0) = SqlProviderLocal.GetParameter("order_skey", DbType.Int32, drLabResultItemFormula("order_skey"), ParameterDirection.Input)
                    parm(1) = SqlProviderLocal.GetParameter("result_item_skey", DbType.Int32, drLabResultItemFormula("result_item_skey"), ParameterDirection.Input)

                    SqlProviderLocal.ExecuteNonQuerySP("sp_lab_order_category_test_item_update_status", parm)


                    ' WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("SID. {2} : Calculate Formula {1} Successful", Now(), drLabResultItemFormula("result_item_desc"), specimenTypeId))
                Catch ex As Exception
                    '   WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("(error in loop) Calculate Formula {1} {2} ", Now(), ex.Message, drLabResultItemFormula("result_item_id")))
                    AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.UpdateFormula_loop: " & ex.Message)
                    If ex.Message.IndexOf("OutOfMemoryException") <> -1 Then
                        AppSetting.WriteErrorCal(comPort, "Error", specimenTypeId)
                    End If
                    log.Error(ex.Message)
                End Try
            Next
            ' SqlProvider.CommitTransaction()
            'Golf Add call WCF 2017-06-26

            '   WriteLog("Calculate Formula", "Receive", Now().ToString("dd-MMMM-yyyy HH:mm:ss") & "   " & String.Format("Calculate Formula {1} End", Now(), specimenTypeId))
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula In End")
            End If
            SqlProviderLocal.Dispose()

        Catch ex As Exception

            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.UpdateFormula_main: " & ex.Message)
            'SqlProvider.RollbackTransaction()
            'log.Error(ex.Message)
        Finally
            '  SqlProvider.Dispose()
        End Try

    End Function

    Public Function GetResultMicro(value As String) As String
        Dim output As StringBuilder = New StringBuilder

        Try
            If value = "30-300" Or value = "30 - 300" Then

                Return value
            End If
            For i = 0 To value.Length - 1

                If value(i).Equals(Chr(43)) Then
                    output.Append("pos")
                    Exit For
                End If

                If value(i).Equals(Chr(45)) Then
                    output.Append("neg")
                    Exit For
                ElseIf IsNumeric(value(i)) Or value(i).Equals(Chr(46)) Then
                    output.Append(value(i))
                End If
            Next

        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultMicro: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return output.ToString()
    End Function

    Public Function GetPatientInformation(ByVal SpecimenId As String) As DataTable

        Dim patientSqlProvider As SqlDataProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable

        Try
            parm = patientSqlProvider.GetParameterArray(1)
            parm(0) = patientSqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            patientSqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_patient_name", parm)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            If patientSqlProvider.Transaction IsNot Nothing AndAlso patientSqlProvider.GetTransactionCount > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetPatientInformation: " & ex.Message)
            End If
            log.Error(ex.Message)
        Catch ex As Exception
            If patientSqlProvider.Transaction IsNot Nothing AndAlso patientSqlProvider.GetTransactionCount > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetPatientInformation: " & ex.Message)
            End If
            log.Error(ex.Message)
        Finally
            patientSqlProvider.Dispose()
        End Try

        Return dt
    End Function

    Public Function sendData(ByVal SpecimenId As String, ByVal code As String, code_test As String, result As String, code_result_skey As String, code_RM As String) As DataTable

        Dim patientSqlProvider As SqlDataProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable

        Try
            parm = patientSqlProvider.GetParameterArray(6)
            parm(0) = patientSqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            parm(1) = patientSqlProvider.GetParameter("code", DbType.String, code, ParameterDirection.Input)
            parm(2) = patientSqlProvider.GetParameter("code_test", DbType.String, code_test, ParameterDirection.Input)
            parm(3) = patientSqlProvider.GetParameter("value", DbType.String, result, ParameterDirection.Input)
            parm(4) = patientSqlProvider.GetParameter("code_result_skey", DbType.String, code_result_skey, ParameterDirection.Input)
            parm(5) = patientSqlProvider.GetParameter("code_RM", DbType.String, code_RM, ParameterDirection.Input)
            patientSqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_insert_result_value_bactria", parm)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            If patientSqlProvider.Transaction IsNot Nothing AndAlso patientSqlProvider.GetTransactionCount > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetPatientInformation: " & ex.Message)
            End If
            log.Error(ex.Message)
        Catch ex As Exception
            If patientSqlProvider.Transaction IsNot Nothing AndAlso patientSqlProvider.GetTransactionCount > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetPatientInformation: " & ex.Message)
            End If
            log.Error(ex.Message)
        Finally
            patientSqlProvider.Dispose()
        End Try
        Return dt
    End Function

    Function GetUA_LIS_FORMULA_IN_INTERFACE_ANALYZER() As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            Dim builder As StringBuilder
            builder = New StringBuilder
            builder.AppendLine("select parm_value from global_parms where global_key = 'LIS_FORMULA_IN_INTERFACE_ANALYZER'")
            dt = SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString()).Tables(0)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            log.Error(ex.Message)
        Finally
            SqlProvider.Dispose()
        End Try


        Return dt

    End Function




    Public Function ExtractResult_humaclot_pro(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemArrI() As String
        Dim ItemArrR() As String
        Dim SpecimenId As String = ""
        Dim resultCode As String = ""
        Dim resultValue As String = ""

        LineArr = DeviceStr.Replace("</TRANSMIT>", "</TRANSMIT>" + vbNewLine).Split(vbNewLine)

        Dim ds As New DataSet()

        For Each itemString As String In LineArr
            Using stringReader As New StringReader(itemString)
                ds = New DataSet()
                ds.ReadXml(stringReader)
            End Using
            Dim dt As DataTable = ds.Tables(0)

            ItemArrI = dt(0).Item("I").ToString.Split("|")
            analyzerDate = DateTime.Parse(ItemArrI(1))
            'SpecimenId = ItemArrI(3) 'Golf Comments
            SpecimenId = ItemArrI(4) 'Golf add 2017-02-28

            ItemArrR = dt(0).Item("R").ToString.Split("|")
            resultCode = "HbA1c" 'Golf Comments 
            resultValue = ItemArrR(1).ToString.Substring(ItemArrR(2).ToString.IndexOf("=") + 2)
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
            Dim dvTestResult As New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = resultValue
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = analyzerDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)
            End If

            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()

            'Tui add 2015-09-22  auto update formula
            UpdateFormula(SpecimenId)

        Next

        Return True

    End Function

    Public Function ExtractResult_Dimension(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim resultCode As String = ""
        Dim ItemStr As String
        Dim resultValue As String = ""
        Dim ItemValue As String = ""
        Dim i, j As Integer
        LineArr = DeviceStr.Split(Chr(28))
        Dim ds As New DataSet()

        j = LineArr.Count

        SpecimenId = LineArr(3)



        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "N").Copy
        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        If dtTestResult.Rows.Count <= 0 Then
            Return False
        End If

        For i = 11 To j - 1

            ItemStr = LineArr(i)
            ItemValue = LineArr(i + 1)

            dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemStr & "'"
            If dvTestResult.Count = 1 Then
                Dim row As DataRow = dtTestResultToUpdate.NewRow
                row("order_skey") = dvTestResult.Item(0)("order_skey")
                row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                row("alias_id") = dvTestResult.Item(0)("alias_id")
                row("result_value") = ItemValue
                row("order_id") = dvTestResult.Item(0)("order_id")
                row("analyzer") = dvTestResult.Item(0)("analyzer")
                row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                row("result_date") = resultDate
                row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                row("rerun") = dvTestResult.Item(0)("rerun")
                dtTestResultToUpdate.Rows.Add(row)

            End If




            i = i + 3
            If (j - 1) - i <= 1 Then
                Exit For
            End If
        Next
        UpdateTestResultValue(dtTestResultToUpdate, False)
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()
        UpdateFormula(SpecimenId)


        Return True

    End Function

    Function ResultAcceptance(ByVal STATUS As String) As ArrayList

        Dim Str As String
        Dim StrArr As New ArrayList
        Dim Reason As String = ""
        '<STX>M<FS>A<FS><FS>E2<ETX
        'Str = Chr(2) & "M" & Chr(28) & STATUS & Chr(28) & Chr(28) & "" & Chr(28)
        'Str &= GetCheckSumValue(Str.Remove(0, 1))
        'Str &= Chr(3)
        'StrArr.Add(Str)

        If STATUS = "R" Then
            Reason = "1"
        End If

        Str = Chr(2) & "M" & Chr(28)
        Str &= STATUS & Chr(28) 'STATUS ACCEPT(A), Reject(R)
        Str &= Reason & Chr(28) 'Reason ACCEPT(empty field), Reject(1)  computer out of memory
        Str &= GetCheckSumValue(Str.Remove(0, 1))
        Str &= Chr(3)
        StrArr.Add(Str)
        Return StrArr
    End Function

    Public Function ReturnOrderDetail_COBAS_C311(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim hn As String = "N/A"
        Dim dtPatient As New DataTable
        Dim SpecimenIdArr() As String
        Dim labSeq As String = ""
        Dim RackIDNo As String = ""
        Dim PositionNo As String = ""
        Dim SampleType As String = ""
        Dim ContainerType As String = ""
        Dim SpecimenIdString As String = SpecimenId
        Dim SpecimenIdSend As String = ""
        Dim Priority As String = ""
        If SpecimenId.Length > 0 Then

            SpecimenIdArr = SpecimenIdString.Split("^")
            SpecimenIdSend = SpecimenIdArr(2)
            SpecimenId = SpecimenIdArr(2).Trim
            labSeq = SpecimenIdArr(3)
            RackIDNo = SpecimenIdArr(4)
            PositionNo = SpecimenIdArr(5)
            SampleType = SpecimenIdArr(7)
            ContainerType = SpecimenIdArr(8)
            If SpecimenIdArr.Length > 9 Then
                Priority = SpecimenIdArr(9)
            End If
            ' Priority = SpecimenIdArr(9) ' Golf 2017-06-15
        End If

        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
        Dim dvTestResult = New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()


        ContainerType = "SC"
        Str = Chr(2) & "1H|" & "\" & "^&|||host^1|||||cobasc311|TSDWN^REPLY|P|1" 'Chr(2) = STX (start-of-text)
        Str &= Chr(13) & "P|1|"
        Str &= Chr(13) & "O|1|"
        Str &= SpecimenIdSend & "|"
        'If Str.IndexOf("^^^698") <> -1 Or Str.IndexOf("^^^714") <> -1 Or Str.IndexOf("^^^965") <> -1 Or Str.IndexOf("^^^452") <> -1 Or Str.IndexOf("^^^989") <> -1 Or Str.IndexOf("^^^990") <> -1 Or Str.IndexOf("^^^991") <> -1 Or Str.IndexOf("^^^704") <> -1 Or Str.IndexOf("^^^668") <> -1 Then
        '    SampleType = "S2"
        'End If
        If dtTestResult.Rows.Count > 0 Then
            Select Case dtTestResult.Rows(0).Item("sample_type")
                Case 2
                    SampleType = "S2"
                Case 1
                    SampleType = "S1"
            End Select

        End If

        Str &= labSeq & "^" & RackIDNo & "^" & PositionNo & "^^" & SampleType & "^" & ContainerType & "|"
        For Each row As DataRow In dtTestResult.Rows
            i += 1
            Str &= SepTest & row.Item("analyzer_ref_cd") & "^"
            Str &= "\"
        Next

        If Str.Substring(Str.Length - 1, 1) = "\" Then
            Str = Str.Substring(0, Str.Length - 1)
        End If

        'Golf 2017-06-15
        If Str.IndexOf("^^^891") <> -1 Then
            Str &= "|" & Priority & "||||||A||||4|||||||"


        Else
            If dtTestResult.Rows.Count > 0 Then
                Select Case dtTestResult.Rows(0).Item("sample_type")
                    Case 2
                        Str &= "|" & Priority & "||||||A||||2|||||||"
                    Case Else
                        Str &= "|" & Priority & "||||||A||||1|||||||"
                End Select
            Else
                Str &= "|" & Priority & "||||||A||||1|||||||"
            End If
        End If

        'If Str.IndexOf("^^^891") <> -1 Then
        '    Str &= "|R||||||A||||4|||||||"
        'Else
        '    Str &= "|R||||||A||||1|||||||"
        'End If
        'Golf 2017-06-15

        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
        Str &= "|||O"
        Str &= Chr(13) & "L|1|N"
        Dim Str1 As String
        Dim Str2 As String
        If Str.Length > 242 Then
            Str1 = Str.Substring(0, 242)
            Str2 = Str.Substring(242, Str.Length - Str1.Length)
            Str1 &= Chr(23) '23	17	00010111	ETB	end of trans. block
            Str1 &= GetCheckSumValue(Str1.Remove(0, 1))
            Str1 &= Chr(13) & Chr(10)
            StrArr.Add(Str1)

            If Str2.Length <= 242 Then
                Str2 = Chr(2) & "2" & Str2
                'Str2 = Chr(2) & Str2
                Str2 &= Chr(13) & Chr(3)
                'Str2 &= GetCheckSumValue(Str2.Remove(0, 1))
                Str2 &= GetCheckSumValue(Str2)
                Str2 &= Chr(13) & Chr(10)
                StrArr.Add(Str2)
            End If
        Else
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
        End If

        AppSetting.WriteErrorLog(comPort, "Send", "Send to COBAS_C311.")
        AppSetting.WriteErrorLog(comPort, "COBAS_C311", String.Join("", StrArr.ToArray()))


        Return StrArr
    End Function

    Public Function ReturnOrderDetail_Maglumi(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim hn As String = "N/A"
        Dim dtPatient As New DataTable
        Dim SpecimenIdArr() As String
        Dim labSeq As String = ""
        Dim RackIDNo As String = ""
        Dim PositionNo As String = ""
        Dim SampleType As String = ""
        Dim ContainerType As String = ""
        Dim SpecimenIdString As String = SpecimenId
        Dim SpecimenIdSend As String = ""
        Dim Priority As String = ""
        If SpecimenId.Length > 0 Then

            SpecimenIdArr = SpecimenIdString.Split("^")
            SpecimenIdSend = SpecimenIdArr(2)
            SpecimenId = SpecimenIdArr(2).Trim
            labSeq = SpecimenIdArr(3)
            RackIDNo = SpecimenIdArr(4)
            PositionNo = SpecimenIdArr(5)
            SampleType = SpecimenIdArr(7)
            ContainerType = SpecimenIdArr(8)
            If SpecimenIdArr.Length > 9 Then
                Priority = SpecimenIdArr(9)
            End If
            ' Priority = SpecimenIdArr(9) ' Golf 2017-06-15
        End If

        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
        Dim dvTestResult = New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()


        ContainerType = "SC"
        Str = Chr(2) & "1H|" & "\" & "^&|||host^1|||||cobasc311|TSDWN^REPLY|P|1" 'Chr(2) = STX (start-of-text)
        Str &= Chr(13) & "P|1|"
        Str &= Chr(13) & "O|1|"
        Str &= SpecimenIdSend & "|"

        If dtTestResult.Rows.Count > 0 Then
            Select Case dtTestResult.Rows(0).Item("sample_type")
                Case 2
                    SampleType = "S2"
                Case 1
                    SampleType = "S1"
            End Select

        End If

        Str &= labSeq & "^" & RackIDNo & "^" & PositionNo & "^^" & SampleType & "^" & ContainerType & "|"
        For Each row As DataRow In dtTestResult.Rows
            i += 1
            Str &= SepTest & row.Item("analyzer_ref_cd") & "^"
            Str &= "\"
        Next

        If Str.Substring(Str.Length - 1, 1) = "\" Then
            Str = Str.Substring(0, Str.Length - 1)
        End If

        'Golf 2017-06-15
        If Str.IndexOf("^^^891") <> -1 Then
            Str &= "|" & Priority & "||||||A||||4|||||||"


        Else
            If dtTestResult.Rows.Count > 0 Then
                Select Case dtTestResult.Rows(0).Item("sample_type")
                    Case 2
                        Str &= "|" & Priority & "||||||A||||2|||||||"
                    Case Else
                        Str &= "|" & Priority & "||||||A||||1|||||||"
                End Select
            Else
                Str &= "|" & Priority & "||||||A||||1|||||||"
            End If
        End If

        'If Str.IndexOf("^^^891") <> -1 Then
        '    Str &= "|R||||||A||||4|||||||"
        'Else
        '    Str &= "|R||||||A||||1|||||||"
        'End If
        'Golf 2017-06-15

        Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
        Str &= "|||O"
        Str &= Chr(13) & "L|1|N"
        Dim Str1 As String
        Dim Str2 As String
        If Str.Length > 242 Then
            Str1 = Str.Substring(0, 242)
            Str2 = Str.Substring(242, Str.Length - Str1.Length)
            Str1 &= Chr(23) '23	17	00010111	ETB	end of trans. block
            Str1 &= GetCheckSumValue(Str1.Remove(0, 1))
            Str1 &= Chr(13) & Chr(10)
            StrArr.Add(Str1)

            If Str2.Length <= 242 Then
                Str2 = Chr(2) & "2" & Str2
                'Str2 = Chr(2) & Str2
                Str2 &= Chr(13) & Chr(3)
                'Str2 &= GetCheckSumValue(Str2.Remove(0, 1))
                Str2 &= GetCheckSumValue(Str2)
                Str2 &= Chr(13) & Chr(10)
                StrArr.Add(Str2)
            End If
        Else
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
        End If

        AppSetting.WriteErrorLog(comPort, "Send", "Send to COBAS_C311.")
        AppSetting.WriteErrorLog(comPort, "COBAS_C311", String.Join("", StrArr.ToArray()))


        Return StrArr
    End Function



    Public Shared Function GetLabResultItemFormula(ByVal SqlProvider As SqlDataProvider, ByVal OrderID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder


            builder.AppendLine("select     lis_lab_specimen_type.specimen_type_id ")
            builder.AppendLine("  from  lis_lab_order inner join lis_lab_specimen_type on ")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_specimen_type.order_skey inner join lis_lab_specimen_type_test_item on ")
            builder.AppendLine("lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey and ")
            builder.AppendLine("lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id inner join lis_lab_test_result_item on ")
            builder.AppendLine("lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey and ")
            builder.AppendLine("lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey inner join lis_lab_result_item on ")
            builder.AppendLine("lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey and ")
            builder.AppendLine("lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey inner join lis_result_item on ")
            builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey  ")
            builder.AppendLine(" where logical_location_cd = 'O-602' and ")
            builder.AppendLine(" lis_result_item.result_item_skey = 70 and lis_result_item.result_type = 'F' and lis_lab_result_item.result_value is null and lis_lab_order.status <> 'CA' and lis_lab_order.status = 'RL' and isnull(logical_location_cd,'')   like 'O%' ")
            builder.AppendLine(" and lis_lab_order.order_skey in (select  distinct a.order_skey from lis_lab_result_value a inner join lis_lab_order b on a.order_skey = b.order_skey where b.status <> 'CA' ) ")
            builder.AppendLine(" group by   lis_lab_specimen_type.specimen_type_id ")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("order_id", DbType.String, OrderID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function


    Function GetSpecimenCal() As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            Dim builder As StringBuilder
            builder = New StringBuilder
            builder.AppendLine("select      lis_lab_specimen_type.specimen_type_id ")
            builder.AppendLine("  from  lis_lab_order inner join lis_lab_specimen_type on ")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_specimen_type.order_skey inner join lis_lab_specimen_type_test_item on ")
            builder.AppendLine("lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey and ")
            builder.AppendLine("lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id inner join lis_lab_test_result_item on ")
            builder.AppendLine("lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey and ")
            builder.AppendLine("lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey inner join lis_lab_result_item on ")
            builder.AppendLine("lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey and ")
            builder.AppendLine("lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey inner join lis_result_item on ")
            builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey  ")
            builder.AppendLine("  where    lis_lab_specimen_type.specimen_type_id like '17%' and convert(nvarchar(6),lis_lab_order.received_date,112) in ('201707','201708') and ")
            builder.AppendLine(" lis_result_item.result_item_skey = 70 and lis_result_item.result_type = 'F' and lis_lab_result_item.result_value is null and lis_lab_order.status <> 'CA' and lis_lab_order.status = 'RL' and isnull(logical_location_cd,'')   like 'O%' ")
            builder.AppendLine(" and lis_lab_order.order_skey in (select  distinct a.order_skey from lis_lab_result_value a inner join lis_lab_order b on a.order_skey = b.order_skey where b.status <> 'CA' ) ")
            builder.AppendLine(" group by   lis_lab_specimen_type.specimen_type_id ")

            dt = SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString()).Tables(0)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
            log.Error(ex.Message)
        Finally
            SqlProvider.Dispose()
        End Try


        Return dt

    End Function


    Function GetANALYZER_skey(ByVal model_cd As String) As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            'Dim builder As StringBuilder
            'builder = New StringBuilder

            'parm = SqlProvider.GetParameterArray(1)
            'parm(0) = SqlProvider.GetParameter("model_cd", DbType.String, model_cd, ParameterDirection.Input) 
            'SqlProvider.ExecuteDataTableSP(dt, "select analyzer_skey from lis_analyzer inner join lis_analyzer_model on lis_analyzer.model_skey=lis_analyzer_model.model_skey where lis_analyzer_model.model_cd = @model_cd", parm)

            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder
            builder.AppendLine("select analyzer_skey from lis_analyzer inner join lis_analyzer_model on lis_analyzer.model_skey=lis_analyzer_model.model_skey where lis_analyzer_model.model_cd = @model_cd ")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("model_cd", DbType.String, model_cd, ParameterDirection.Input)

            dt = SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetANALYZER_skey: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetANALYZER_skey: " & ex.Message)
            log.Error(ex.Message)
        Finally
            SqlProvider.Dispose()
        End Try


        Return dt

    End Function
    Public Function Data_Hex_Asc(ByRef Data As String) As String
        Dim Data1 As String = ""
        Dim sData As String = ""
        While Data.Length > 0
            'first take two hex value using substring.
            'then convert Hex value into ascii.
            'then convert ascii value into character.
            If Data <> " " Then
                If Data <> "" Then
                    Data1 = System.Convert.ToChar(System.Convert.ToUInt32(Data.Substring(0, 2), 16)).ToString()
                    sData = sData & Data1
                    Data = Data.Substring(2, Data.Length - 2)
                End If

            Else
                Data = ""
            End If
        End While
        Return sData
    End Function
    Function FromHex(ByVal Text As String) As String

        If Text Is Nothing OrElse Text.Length = 0 Then
            Return String.Empty
        End If

        Dim Bytes As New List(Of Byte)
        For Index As Integer = 0 To Text.Length - 1 Step 2
            Bytes.Add(Convert.ToByte(Text.Substring(Index, 2), 16))
        Next

        Dim E As System.Text.Encoding = System.Text.Encoding.Unicode
        Return E.GetString(Bytes.ToArray)

    End Function

    Public Function TESTCheckSum()
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SampleType As String = ""
        Dim Running As Int32 = 1
        Try



            Str = Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|20201216201652" 'Chr(2) = STX (start-of-text)
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""

            Str = Chr(2) & CStr(Running) & "P|1|123456|||19920430|M||||||"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""

            Str = Chr(2) & CStr(Running) & "O|1|200000446|201||R||||||||||||||||||||F"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""

            Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Echo.")
            AppSetting.WriteErrorLog(comPort, "Echo", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ExtractResult_HA8180V(ByVal DeviceStr As String) As Boolean 'PLOY ADD 2020.10.05
        Dim LineArr() As String
        Dim ItemArrQ() As String
        Dim SpecimenId As String = ""
        Dim resultCode As String = ""
        Dim resultValue As String = "" 'Ploy Add 
        Dim resultUnit As String = ""   'Ploy Add 
        Dim AbnormalFlag As String = ""   'Ploy Add 
        Dim TestResultStatus As String = ""   'Ploy Add 
        Dim ItemArr() As String
        Dim tempDate As String = String.Empty

        Dim ds As New DataSet()

        LineArr = DeviceStr.Split(vbNewLine)

        For Each itemString As String In LineArr
            If itemString.IndexOf("O|") <> -1 Then

                ItemArrQ = itemString.ToString.Split("|")
                SpecimenId = ItemArrQ(2)
                Dim TmpSpecimenId = SpecimenId.Split("-")
                SpecimenId = TmpSpecimenId(0)
            End If
        Next

        'ดึงค่า Result ทั้งหมดที่ Order ไว้
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy

        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        For Each itemString As String In LineArr
            If itemString.IndexOf("O|") <> -1 Then

            ElseIf itemString.IndexOf("R|") <> -1 Then
                ItemArr = itemString.Split(New String() {"^^^", "|", "Values"}, StringSplitOptions.None)

                If ItemArr(3) = "ValueHbA1c" Then
                    resultCode = ItemArr(3)
                    resultValue = ItemArr(4)
                Else
                    Continue For
                End If


                'กรอง ResultCode จาก DvTestResult ที่ดึงจาก Database
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"

                If dvTestResult.Count = 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = resultValue
                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = analyzerDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)
                End If

                UpdateTestResultValue(dtTestResultToUpdate, False)
                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()

                'Tui add 2015-09-22  auto update formula
                UpdateFormula(SpecimenId)
            End If

        Next
        Return True

    End Function
    Public Function ReturnOrderDetailStago(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList  'Ploy Add 21.01.26
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()


            Str &= Chr(2) & "1H|\^&|||STA_Satlellite_Max^CL79100693^1|||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            'Old Ploy Comment 2021.01.14
            Running += 1
            Str = ""
            Str = Chr(2) & CStr(Running) & "P|1|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
            Str &= "|||"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Birthday").ToString)
            Str &= "|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("Sex").ToString)
            Str &= "||||||"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            Running += 1
            Str = ""
            Str = Chr(2) & CStr(Running) & "O|1|"
            Str &= SpecimenId & "^^^||"
            For Each row As DataRow In dtTestResult.Rows
                If row.Item("analyzer_ref_cd") = "1F" Then
                    Str &= "^^^1\"
                ElseIf row.Item("analyzer_ref_cd") = "2F" Then
                    Str &= "^^^2\"
                ElseIf row.Item("analyzer_ref_cd") = "3F" Then
                    Str &= "^^^3\"
                ElseIf row.Item("analyzer_ref_cd") = "4F" Then
                    Str &= "^^^4\"
                End If
            Next
            'ตัด \ ตัวสุดท้ายออก
            Str = Str.Substring(0, Str.Length - vbCrLf.Length + 1)

            Str &= "^F^64^^^^|R||N|O"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            Running += 1
            Str = ""
            Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            StrArr.Add(Chr(4))

            AppSetting.WriteErrorLog(comPort, "Send", "Send to Stago.")
            AppSetting.WriteErrorLog(comPort, "Stago", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ExtractResult_Stago(ByVal DeviceStr As String) As Boolean 'PLOY ADD 2020.10.05
        Dim LineArr() As String
        Dim ItemArrQ() As String
        Dim SpecimenId As String = ""
        Dim resultCode As String = ""
        Dim resultValue As String = "" 'Ploy Add 
        Dim AbnormalFlag As String = ""
        Dim TestResultStatus As String = ""
        Dim ItemArr() As String
        Dim tempDate As String = String.Empty
        Dim ResultType As String = "" 'Ploy Add 2020.12.02

        Dim ds As New DataSet()

        LineArr = DeviceStr.Split(vbNewLine)
        For Each itemString As String In LineArr
            If itemString.IndexOf("O|") <> -1 Then

                ItemArrQ = itemString.ToString.Split("|")
                SpecimenId = ItemArrQ(2)
                ' Dim TmpSpecimenId = SpecimenId.Split("^")
                Dim TmpSpecimenId = SpecimenId.Split(New String() {"^", "-"}, StringSplitOptions.None)
                SpecimenId = TmpSpecimenId(0)
            End If
        Next

        'ดึงค่า Result ทั้งหมดที่ Order ไว้
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy

        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        For Each itemString As String In LineArr
            resultCode = ""

            If itemString.IndexOf("O|") <> -1 Then

            ElseIf itemString.IndexOf("R|") <> -1 Then
                ItemArr = itemString.Split(New String() {"^", "|"}, StringSplitOptions.None)

                If ItemArr(5) = "1" Or ItemArr(5) = "2" Or ItemArr(5) = "3" Or ItemArr(5) = "4" Then
                    resultCode = ItemArr(5) + ItemArr(6)
                    resultValue = ItemArr(8)
                Else
                    Continue For
                End If

                If Not ItemArr(5) = "1" And Not ItemArr(5) = "2" And Not ItemArr(5) = "3" And Not ItemArr(5) = "4" Then
                    Continue For
                End If

                'กรอง ResultCode จาก DvTestResult ที่ดึงจาก Database
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & resultCode & "'"

                If dvTestResult.Count = 1 Then
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(0)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(0)("alias_id")
                    row("result_value") = resultValue
                    row("order_id") = dvTestResult.Item(0)("order_id")
                    row("analyzer") = dvTestResult.Item(0)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                    row("result_date") = analyzerDate
                    row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(0)("rerun")
                    dtTestResultToUpdate.Rows.Add(row)

                Else
                    Continue For
                End If


            End If

        Next
        UpdateTestResultValue(dtTestResultToUpdate, False)
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        UpdateFormula(SpecimenId)

        Return True

    End Function
    Public Function ReturnOrderDetailDxC700AU(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList  'Ploy Add 21.02.22
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim Rack_No As String
        Dim CupPos As String
        Dim S_No As String
        Dim Dummy As String
        Dim check_urine As String = ""
        'Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try
            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If

            Dim SpecimenIdArr() As String

            'SpecimenIdArr = SpecimenId.Split("R ", Chr(2), "N", Chr(3))
            'SpecimenId = SpecimenIdArr(2).Substring(5, SpecimenIdArr(2).Length - 5).Replace(" ", "")
            SpecimenIdArr = SpecimenId.Split("R ", Chr(2), "N", Chr(3), "E", Chr(3))
            SpecimenId = SpecimenIdArr(2).Substring(4).Trim()
            Rack_No = SpecimenIdArr(1).Substring(0, 4)
            CupPos = SpecimenIdArr(1).Substring(4)
            S_No = SpecimenIdArr(2).Substring(0, 4).Trim()
            Dummy = "    "

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'If dtTestResult.Rows.Count > 0 Then
            'Str = Chr(2) & "S " & SpecimenIdArr(1).Substring(1, SpecimenIdArr(1).Length - 1) & dtTestResult.Rows(0).Item("sample_type") & SpecimenIdArr(2) 'Ploy comment 2021.03.06

            'Str &= "    E"
            'For Each row As DataRow In dtTestResult.Rows
            'Go
            'Str &= row.Item("analyzer_ref_cd")
            'Next
            'Str &= Chr(3)
            'StrArr.Add(Str)
            'End If


            If dtTestResult.Rows.Count > 0 Then

                For Each row As DataRow In dtTestResult.Rows
                    'Go
                    If row.Item("sample_type") = "U" Then
                        check_urine = "U"
                    End If
                Next

                If S_No.IndexOf("P") <> -1 Then
                    If check_urine = "U" Then
                        Str = Chr(2) & "S" & Rack_No & CupPos & "U" & S_No & SpecimenId
                    Else
                        Str = Chr(2) & "S" & Rack_No & CupPos & " " & S_No & SpecimenId
                    End If
                Else
                    If check_urine = "U" Then
                        Str = Chr(2) & "S" & Rack_No & CupPos & "U" & S_No & SpecimenId
                    Else
                        Str = Chr(2) & "S" & Rack_No & CupPos & " " & S_No & SpecimenId
                    End If
                End If

                Str &= Dummy & "E"

                For Each row As DataRow In dtTestResult.Rows
                    'Go
                    Str &= row.Item("analyzer_ref_cd")
                Next
                Str &= Chr(3)
                StrArr.Add(Str)
            End If

            AppSetting.WriteErrorLog(comPort, "Send", "Send to DxC700AU.")
            AppSetting.WriteErrorLog(comPort, "DxC700AU", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function

    Public Function ReturnOrderDetailCoagACL(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList  'Ploy Add 21.02.22
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|" & "@^\" & "<0_0><1025080549_50>||ACL-TOP-21|||||LIS-HOST-03||P|1394-97|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("full_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                i = 0
                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= SpecimenId & "||"
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "|"
                Next

                If Str.Substring(Str.Length - 1, 1) = "|" Then
                    If dtTestResult.Rows.Count > 1 Then
                        Str = Str.Substring(0, Str.Length - 1)
                        Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                        Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                        Dim StrList As String() = TestItemList.Split("|")
                        Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                        Dim TestItemListCheckDup2 As String = ""
                        For Each xx As String In TestItemListCheckDup
                            TestItemListCheckDup2 &= xx
                            TestItemListCheckDup2 &= "|"
                        Next

                        Str = HeadList & TestItemListCheckDup2
                    End If
                    Str = Str.Substring(0, Str.Length - 1)
                End If

                'New
                Str &= "|R"
                Str &= "||||||||||"
                Str &= "|"
                Str &= "P" 'plasma
                Str &= "||||||||||"
                Str &= "O" 'Order record
                'New

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "L | 1 | N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to CoagACL.")
            AppSetting.WriteErrorLog(comPort, "CoagACL", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function

    Public Function ExtractResult_CoagACL(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False

        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(3)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "DCOBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count

        For i = 0 To j - 1

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True

    End Function

    Public Function ExtractResult_GEM4000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 2)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(0).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    result_type = ItemArr(8).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 And result_type = "F" Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailGEM4000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(3)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "MSH|" & "^~\" & "&|||||" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= "||OML^O21|1249|||||AL|NE"
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Str = Chr(2) & "ORC|NW|99999|"
            Str &= AppSetting.chkDBNull(SpecimenId.ToString)
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "^^^ALL"
                Str &= "||||||||X"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "OBR|"
                Str &= "1|"
                Str &= "99999|"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "|"
                Running += 1
                'New
                i = 0
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= row.Item("analyzer_ref_cd")
                    Str &= "~"
                    'If row.Item("analyzer_ref_cd") <> "62" And row.Item("analyzer_ref_cd") <> "63" Then
                    '    Str &= SepTest & row.Item("analyzer_ref_cd")
                    '    Str &= "\"
                    '    'Golf Start 2017-01-08 
                    'Else
                    '    If Str.IndexOf("^^^61") = -1 And (row.Item("analyzer_ref_cd") <> "62" Or row.Item("analyzer_ref_cd") <> "63") Then
                    '        Str &= SepTest & "61"
                    '        Str &= "\"
                    '    End If
                    '    'Golf End 2017-01-08 
                    'End If
                Next


                'New
                Str &= "||||||O||||BLDA^^N^^^^P|Freeman^Dyson^J"

                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))

                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            End If

            Str = Chr(2) & CStr(Running) & "SAC|10" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to GEM4000.")
            AppSetting.WriteErrorLog(comPort, "GEM4000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function

    Public Function ExtractResult_V3600(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                Dim SpecimenIdArr2() As String = SpecimenIdArr(2).Split("^") 'Golf Edited 2021-08-31
                If SpecimenIdArr2(0).IndexOf("D") <> -1 Then
                    SpecimenId = SpecimenIdArr2(0)
                Else
                    SpecimenId = SpecimenIdArr2(0).Trim.Replace(" ", "").Replace("^", "").Trim 'Golf Edited 2021-08-31
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                If ItemArr(3).IndexOf("No Result") <> -1 Then
                    Continue For
                End If
                Dim assay_code As String = ""
                    Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    Dim assay_codeArr2() As String = assay_code.Split("+") 'Golf Edited 2021-08-31
                    ItemArr(2) = assay_codeArr2(1).Trim 'Golf Edited 2021-08-31
                    assay_code = assay_codeArr2(1) 'Golf Edited 2021-08-31
                    'result_type = assay_codeArr(10).Trim 'Golf Edited 2021-08-31
                End If

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                    If dvTestResult.Count >= 1 Then
                        For ii = 0 To dvTestResult.Count - 1
                            Dim row As DataRow = dtTestResultToUpdate.NewRow
                            row("order_skey") = dvTestResult.Item(ii)("order_skey")
                            row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                            row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                            row("alias_id") = dvTestResult.Item(ii)("alias_id")
                            row("result_value") = ItemArr(3).ToString()
                            row("order_id") = dvTestResult.Item(ii)("order_id")
                            row("analyzer") = dvTestResult.Item(ii)("analyzer")
                            row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                            row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                            row("result_date") = resultDate
                            row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                            row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                            row("rerun") = dvTestResult.Item(ii)("rerun")
                            row("sample_type") = dvTestResult.Item(ii)("sample_type")
                            dtTestResultToUpdate.Rows.Add(row)
                        Next
                    End If


                End If

                If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailV3600(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdListArray As String() = SpecimenId.Split("^")
                SpecimenId = SpecimenIdListArray(1)
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Str = Chr(2) & "1H|\^&|||||||||||LIS2-A|" 'Chr(2) = STX (start-of-text)
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "Q|1|"
                Str &= "^"
                Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                Str &= "||"
                Str &= "ALL"
                Str &= "||||||||A"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                'New
                Str = Chr(2) & CStr(Running) & "P|1|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                Str &= "|||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                'Start 
                Str = Chr(2) & CStr(Running) & "O|" & CStr(Running) & "|"
                Str &= SpecimenId & "||"
                i = 0
                Dim x As Integer = dtTestResult.Rows.Count
                For Each row As DataRow In dtTestResult.Rows
                    If i = 0 Then
                        Str &= SepTest & "1.0000+"
                    Else
                        Str &= ""
                    End If

                    Str &= row.Item("analyzer_ref_cd")
                    Str &= "+1.0"
                    i += 1
                    If x = i Then
                        Str &= ""
                    Else
                        Str &= "\"
                    End If
                Next
                'End Start

                Str &= "|R" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "||||||"
                Str &= "N"
                Str &= "||||"
                Str &= "5" 'Specimen Type
                Str &= "||||||||||"
                Str &= "F" 'Report Types O: Order, Q: Order in response to a Query Request.

                'New

                Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""

            End If

            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "L|1|T" & Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
            Else
                Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
            End If

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)


            AppSetting.WriteErrorLog(comPort, "Send", "Send to V3600.")
            AppSetting.WriteErrorLog(comPort, "V3600", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function

    Public Function ReturnOrderDetailXN550(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim
            End If



            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|" & "\" & "^&||||||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|" & "\" & "^&|||||||||||E1394-97" 'Chr(2) = STX 
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            Str = Chr(2) & CStr(Running) & "P|1|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
            Str &= "|||"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
            Str &= "||"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
            Str &= "|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
            Str &= "|"
            Str &= "||||||||||||||||"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            i = 0
            Str = Chr(2) & CStr(Running) & "O|1|"
            Str &= PacketSpecimenTypeID & "||"
            For Each row As DataRow In dtTestResult.Rows
                i += 1
                Str &= SepTest & row.Item("analyzer_ref_cd")
                Str &= "\"
            Next

            If Str.Substring(Str.Length - 1, 1) = "\" Then
                If dtTestResult.Rows.Count > 1 Then
                    Str = Str.Substring(0, Str.Length - 1)
                    Dim HeadList As String = Str.Substring(0, Str.IndexOf("^"))
                    Dim TestItemList As String = Str.Substring(Str.IndexOf("^"))

                    Dim StrList As String() = TestItemList.Split("\")
                    Dim TestItemListCheckDup As String() = StrList.Distinct().ToArray()
                    Dim TestItemListCheckDup2 As String = ""
                    For Each xx As String In TestItemListCheckDup
                        TestItemListCheckDup2 &= xx
                        TestItemListCheckDup2 &= "\"
                    Next

                    Str = HeadList & TestItemListCheckDup2
                End If
                Str = Str.Substring(0, Str.Length - 1)
            End If

            'New
            If dtTestResult.Rows.Count <= 0 Then
                Str &= "|||||||"
                Str &= ""
                Str &= "|"
                Str &= "" 'Relevant Clinical Information()
                Str &= "|"
                Str &= ""
                Str &= "||"
                Str &= "" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Y"
            Else
                Str &= "|||||||"
                'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0") 'Collection Date and Time
                'Str &= "||||"
                Str &= "N" ' Action Code  N: New order for a patient sample, A: Unconditional Add order for a patient sample, C: Cancel or Delete the existing order, Q: Control Sample
                Str &= "|"
                Str &= "" 'Relevant Clinical Information()
                Str &= "|"
                Str &= ""
                Str &= "||"
                Str &= "" 'Specimen Type
                Str &= "||||||||||"
                Str &= "Q" 'Report Types O: Order, Q: Order in response to a Query Request.
            End If

            'New

            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""

            Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))

            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to C8000.")
            AppSetting.WriteErrorLog(comPort, "C8000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_XN550(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        If analyzerModel = "COBAS_C311" Then
            LineArr = DeviceStr.Split(Chr(13))
        Else
            LineArr = DeviceStr.Split(Chr(10))
        End If

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 And SpecimenId.Length <= 1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                Dim SpecimenIdArr2() As String = SpecimenIdArr(3).Split("^")
                SpecimenId = SpecimenIdArr2(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(4).Trim
                    ItemArr(2) = assay_codeArr(4).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailXN3000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim
            End If



            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|" & "\" & "^&||||||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|\^&|||||||||||E1394-97" 'Chr(2) = STX 
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            If dtTestResult.Rows.Count <= 0 Then
                Str = Chr(2) & CStr(Running) & "P|1"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New

                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= PacketSpecimenTypeID & "||||"
                Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                Str &= "|||||||||||||||||||Y"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
            Else
                Str = Chr(2) & CStr(Running) & "P|1|||100|^"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
                Str &= "||"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
                Str &= "|"
                Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
                Str &= "|"
                Str &= "||||"
                Str &= "^Dr.1||||||||||||^^^WEST"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New

                Str = Chr(2) & CStr(Running) & "O|1|"
                Str &= PacketSpecimenTypeID & "||"
                i = 0
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest & row.Item("analyzer_ref_cd")
                    Str &= "\"
                Next

                Str &= "^^"
                Str &= Chr(13) & Chr(23)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New

                Str = Chr(2) & CStr(Running) & "^^PCT"
                Str &= "||"
                Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                Str &= "|||||"
                Str &= "N"
                Str &= "||||||||||||||"
                Str &= "Q"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New

            End If
            Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to XN3000.")
            AppSetting.WriteErrorLog(comPort, "XN3000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_XN3000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 And SpecimenId.Length <= 1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                Dim SpecimenIdArr2() As String = SpecimenIdArr(3).Split("^")
                SpecimenId = SpecimenIdArr2(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(4).Trim
                    ItemArr(2) = assay_codeArr(4).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_HCLAB(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("OBR") <> -1 And SpecimenId.Length <= 1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                SpecimenId = SpecimenIdArr(3).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("OBX") <> -1 Then

                ItemArr(5) = ItemArr(5).Replace("^", "").Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(3).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(3).Split("^")
                    assay_code = assay_codeArr(0).Trim
                    ItemArr(3) = assay_codeArr(0).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(3) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(5).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("OBX") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailXN1000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim
            End If



            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|" & "\" & "^&||||||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|\^&|||||||||||E1394-97" 'Chr(2) = STX 
            'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            Str = Chr(2) & CStr(Running) & "P|1|||100|^"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
            Str &= "||"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
            Str &= "|"
            Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
            Str &= "|"
            Str &= "||||"
            Str &= "^Dr.1||||||||||||^^^WEST"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            i = 0
            Str = Chr(2) & CStr(Running) & "O|1|2^1^"
            Str &= PacketSpecimenTypeID & "||"

            For Each row As DataRow In dtTestResult.Rows
                i += 1
                Str &= SepTest & row.Item("analyzer_ref_cd")
                Str &= "\"
            Next

            Str &= "^^"
            Str &= Chr(23)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str = Chr(2) & CStr(Running) & "^^PCT"
            Str &= "||"
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= "|||||"
            Str &= "N"
            Str &= "||||||||||||||"
            Str &= "Q"
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to XN1000.")
            AppSetting.WriteErrorLog(comPort, "XN1000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_XN1000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 And SpecimenId.Length <= 1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)
                SpecimenIdStr = SpecimenIdStr.Replace("||", "|")
                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                Dim SpecimenIdArr2() As String = SpecimenIdArr(2).Split("^")
                SpecimenId = SpecimenIdArr2(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(4).Trim
                    ItemArr(2) = assay_codeArr(4).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailAU5800(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList  'Pong Add 22.01.02

        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim Rack_No As String
        Dim CupPos As String
        Dim S_No As String
        Dim Dummy As String
        'Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try
            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If

            Dim SpecimenIdArr() As String
            'SID 7 digits 
            SpecimenIdArr = SpecimenId.Split("R ", Chr(2), "N", Chr(3), "E", Chr(3))
            SpecimenId = SpecimenIdArr(2).Substring(4).Trim()
            Rack_No = SpecimenIdArr(1).Substring(0, 5).Trim()
            CupPos = SpecimenIdArr(1).Substring(5).Trim()
            S_No = SpecimenIdArr(2).Substring(0, 4).Trim()
            Dummy = "    "
            'dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            If dtTestResult.Rows.Count > 0 Then

                Str = Chr(2) & "S " & Rack_No & CupPos & " " & S_No & SpecimenId
                Str &= Dummy & "E" '=(22_02_10_1_OK)
                'Str = Chr(2) & "S " & Rack_No & CupPos & " " & S_No & "               " & SpecimenId & Dummy & "E" '=(22_02_10_2)
                For Each row As DataRow In dtTestResult.Rows
                    'Go
                    Str &= row.Item("analyzer_ref_cd")
                Next
                Str &= Chr(3)
                StrArr.Add(Str)
            End If

            AppSetting.WriteErrorLog(comPort, "Send", "Send to AU5800.")
            AppSetting.WriteErrorLog(comPort, "AU5800", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_AU5800(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        'Dim ItemArr() As String
        'Dim itemArrOrd() As 

        'Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        'Dim firstPat As Boolean = True
        'Dim ResultStatus As Boolean = False
        Dim textTrim As String = ""
        Dim remain As String = ""

        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If

        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(2))

        j = LineArr.Count

        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)

            If ItemStr.IndexOf(Chr(6)) <> -1 Then
                Continue For
            End If

            If ItemStr.Length <= 1 Then
                Continue For
            End If

            If ItemStr.IndexOf("DB") <> -1 Then
                Continue For
            End If
            Dim CheckHU As String = ""
            If ItemStr.IndexOf("U") <> -1 Then
                Dim CheckU As String = ItemStr.Replace("U", Chr(32))
                CheckHU = "U"
                ItemStr = CheckU
            End If

            Dim CheckHB As String = ""
            If ItemStr.IndexOf("W") <> -1 Then

                Dim CheckW As String = ItemStr.Replace("W", Chr(32))
                CheckHB = "C"
                ItemStr = CheckW

            End If

            SpecimenId = ItemStr
            Dim checkP As Integer = InStr(1, SpecimenId, "E", 1) - 1
            SpecimenId = SpecimenId.Substring(0, checkP)
            SpecimenId = SpecimenId.Replace(Chr(32), "^")

            If SpecimenId.IndexOf("^^^^^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^^^^^", "^")
            End If
            If SpecimenId.IndexOf("^^^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^^^", "^")
            End If
            If SpecimenId.IndexOf("^^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^^", "^")
            End If
            If SpecimenId.IndexOf("^^") <> -1 Then
                SpecimenId = SpecimenId.Replace("^^", "^")
            End If

            Dim subArr() As String = SpecimenId.Split("^")

            'Check 12 Digit
            If subArr.Count = 6 Then
                SpecimenId = subArr(4).Trim()
            End If

            'check 7 Digit
            If subArr.Count = 4 Then
                SpecimenId = subArr(2)
                SpecimenId = SpecimenId.Substring(4).Trim()
            End If

            'Check QC Test
            If subArr.Count = 3 Then
                SpecimenId = subArr(2).Trim()
            End If

            If subArr.Count = 2 Then
                SpecimenId = subArr(1).Trim()
            End If

            If SpecimenId = "" Then
                Continue For
            End If

            If SpecimenId.Length >= 2 Then


                dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()



                ItemStr = ItemStr.Replace("Hr", "n")
                ItemStr = ItemStr.Replace("r", "n")
                ItemStr = ItemStr.Replace("Tp", "n")
                ItemStr = ItemStr.Replace("T", "n")
                ItemStr = ItemStr.Replace("G", "n")
                ItemStr = ItemStr.Replace("-", "n")

                For Each row As DataRow In dtTestResult.Rows

                    Dim textC As String = row.Item("analyzer_ref_cd")
                    Dim textC2 As String = "Hn" & textC
                    ItemStr = ItemStr.Replace(textC, textC2)

                Next

                ItemStr = ItemStr.Replace("n", "r")
                ItemStr = ItemStr.Replace(Chr(32), "n")

                If ItemStr.IndexOf("nnnnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnnnn", "")
                End If
                If ItemStr.IndexOf("nnnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnnn", "")
                End If
                If ItemStr.IndexOf("nnnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnnn", "")
                End If
                If ItemStr.IndexOf("nnnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnnn", "")
                End If
                If ItemStr.IndexOf("nnnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnnn", "")
                End If
                If ItemStr.IndexOf("nnn") <> -1 Then
                    ItemStr = ItemStr.Replace("nnn", "")
                End If
                If ItemStr.IndexOf("nn") <> -1 Then
                    ItemStr = ItemStr.Replace("nn", "")
                End If

                ItemStr = ItemStr.Replace("r", "n")
                Dim CheckE As String = ItemStr.Replace(Chr(32), "")

                textTrim = CheckE.Substring(InStr(1, CheckE, "E", 1))

                Dim i2 As Long

                Dim result2() As String = textTrim.Split("n")

                Dim assay_code As String = ""
                Dim assay_rusult As String = ""

                For i2 = 0 To UBound(result2)

                    Dim trimArr As String = result2(i2)

                    If result2(i2).Length <= 3 Then
                        Continue For
                    End If

                    If result2(i2).IndexOf("r") <> -1 Then
                        trimArr = result2(i2).Replace("r", "")
                    End If

                    If result2(i2).IndexOf("H") <> -1 Then
                        trimArr = result2(i2).Replace("H", "")
                    End If

                    If result2(i2).IndexOf("h") <> -1 Then
                        trimArr = result2(i2).Replace("h", "")
                    End If

                    If result2(i2).IndexOf("b") <> -1 Then
                        trimArr = result2(i2).Replace("b", "")
                    End If

                    If result2(i2).IndexOf("L") <> -1 Then
                        trimArr = result2(i2).Replace("L", "")
                    End If
                    If result2(i2).IndexOf(Chr(3)) <> -1 Then
                        trimArr = result2(i2).Replace(Chr(3), "")
                    End If

                    Dim resultT = trimArr.Trim()

                    If resultT.Length < 2 Then
                        Continue For

                    End If

                    assay_code = resultT.Substring(0, 3) 'Code

                    assay_rusult = resultT.Substring(3) 'Result

                    If assay_rusult.IndexOf("?") <> -1 Then
                        Continue For
                    End If

                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & assay_code & "'"
                    If dvTestResult.Count >= 1 Then

                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(0)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(0)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(0)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(0)("alias_id")
                        row("result_value") = assay_rusult.ToString()
                        row("order_id") = dvTestResult.Item(0)("order_id")
                        row("analyzer") = dvTestResult.Item(0)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(0)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(0)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(0)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(0)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(0)("rerun")
                        row("sample_type") = dvTestResult.Item(0)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)

                    End If

                Next

            End If

        Next
endLine:

        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue1")
        End If
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)

        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If

        Return True

    End Function

    Public Function ReturnOrderDetailAU480(ByVal DistinctionName As String, ByVal ResultStr As String) As String
        Debug.Write("ReturnOrderDetailAU480")
        Dim Str As String = ""
        Dim Result As String = ""
        Dim dtPatient As New DataTable
        Try
            Select Case True
                Case DistinctionName = "RequestStart"

                Case DistinctionName = "NormalRequest"
                    Str = Chr(2)
                    Str &= "S" & Chr(32)
                    Str &= ResultStr.Substring(3, 4)
                    Str &= ResultStr.Substring(7, 2)
                    Str &= If(ResultStr.Substring(9, 1) = "N", Chr(32), ResultStr.Substring(9, 1))
                    Str &= ResultStr.Substring(10, 4)
                    Dim LastData As String = ResultStr.Substring(14)
                    Dim SpecimenId As String = LastData.Substring(0, LastData.IndexOf(Chr(3)))
                    Dim dtTestResult As DataTable = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
                    Dim dvTestResult = New DataView(dtTestResult)

                    Str &= SpecimenId
                    Str &= Chr(32) & Chr(32) & Chr(32) & Chr(32)
                    Str &= "E"
                    AppSetting.WriteErrorLog(comPort, "Error", "ReturnOrderDetailAU480 Get SpecimenId: =>" & SpecimenId & "<-")

                    If dvTestResult.Count >= 1 Then
                        For ii = 0 To dvTestResult.Count - 1
                            Dim ItemRefCD = dvTestResult.Item(ii)("analyzer_ref_cd")
                            AppSetting.WriteErrorLog(comPort, "Error", "ReturnOrderDetailAU480 ItemRefCD => " & ItemRefCD & " :ItemRefCD")
                            If ItemRefCD <> "" Then
                                Str &= ItemRefCD
                            End If
                            AppSetting.WriteErrorLog(comPort, "Error", "ReturnOrderDetailAU480 ItemRefCD => " & Str & " :ItemRefCD")
                        Next

                        Str &= Chr(3)
                        AppSetting.WriteErrorLog(comPort, "Error", "ReturnOrderDetailAU480 Finish => " & Str)
                    Else
                        AppSetting.WriteErrorLog(comPort, "Error", "ReturnOrderDetailAU480 Finish => Empty SID: " & SpecimenId)
                        Str = ""
                    End If
                    dtTestResult.Dispose()
                Case DistinctionName = "RepeatRunRequest"
                    Str = Chr(2)
                    Str &= "SH"
                    Str &= ResultStr.Substring(3, 4)
                    Str &= ResultStr.Substring(7, 2)
                    Str &= ResultStr.Substring(9, 1)
                    Str &= ResultStr.Substring(10, 4)
                    Str &= ResultStr.Substring(10, 4)
                    Str &= "E"
                    Str &= "0"
                    Str &= Chr(32) & Chr(32) & Chr(32)
                    Str &= Chr(32) & Chr(32)
                    Str &= "1"
                    Str &= "01"
                    Str &= "0"
                    Str &= Chr(3)
                Case DistinctionName = "AutomaticRepeatRunRequest"
                    Str = Chr(2)
                    Str &= "Sh"
                    Str &= ResultStr.Substring(3, 4)
                    Str &= ResultStr.Substring(7, 2)
                    Str &= ResultStr.Substring(9, 1)
                    Str &= ResultStr.Substring(10, 4)
                    Dim LastData As String = ResultStr.Substring(14)
                    Str &= LastData.Substring(0, LastData.IndexOf(Chr(3)))
                    Str &= Chr(32) & Chr(32) & Chr(32) & Chr(32)
                    Str &= "E"
                    Str &= "01"
                    Str &= "0"
                    Str &= Chr(3)
                Case DistinctionName = "RequestEnd"
                    Str = Chr(2)
                    Str &= "SE"
                    Str &= Chr(3)
                Case Else
            End Select
            Result = Str

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return Result
    End Function
    Private Shared Function GetNumbers(ByVal input As String) As String
        Return New String(input.Where(Function(c) Char.IsDigit(c) Or c = ".").ToArray())
    End Function
    Public Function ExtractResult_AU480(ByVal DeviceStr As String) As Boolean
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_AU480")
        End If
        If DeviceStr.IndexOf(Chr(3)) = -1 Then  '****** ETX *******
            Return False
        End If

        Dim dvTestResult As New DataView
        Dim analyzer_ref_cd As String = ""
        Dim Process As String = ""
        Dim DistinctionCode As String = ""
        Dim RackNo As String = ""
        Dim CupPosition As String = ""
        Dim SampleType As String = ""
        Dim SampleNo As String = ""
        Dim SampleID As String = ""
        Dim OriginalSampleNo As String = ""
        Dim Dummy As String = ""
        Dim DataClassificationNo As String = ""
        Dim Sex As String = ""
        Dim YearAge As String = ""
        Dim MonthAge As String = ""
        Dim PatientInformations As ArrayList = New ArrayList()
        Dim OnlineTestNo As String = ""
        Dim DiluentType As String = ""
        Dim ReagentLotNo As String = ""
        Dim ReagentBottleNo As String = ""
        Dim AnalysisData As String = ""
        Dim DataFlag As String = ""
        Dim ReagentBlankDataClassification As String = ""
        Dim CalibratorNo As String = ""
        Dim ControlNo As String = ""
        Dim UnivTestID As String = ""
        Dim Level As String = ""

        Process = DeviceStr.Substring(1, 1)
        DistinctionCode = DeviceStr.Substring(2, 1)
        AppSetting.WriteErrorLog(comPort, "Error", "Get Process->" & Process & "<-")
        AppSetting.WriteErrorLog(comPort, "Error", "Get DistinctionCode->" & DistinctionCode & "<-")

        Dim _DistinctionName = String.Empty
        DeviceStr = DeviceStr.Replace(Chr(10), String.Empty).Trim
        Dim datas() As String = DeviceStr.Remove(DeviceStr.Length - 1, 1).Split(Chr(32))
        Select Case True
            Case Process = "D" And DistinctionCode = "B"
                Return True
            Case Process = "D" And DistinctionCode = " "
                Dim StartSampleID As String = DeviceStr.Substring(14)
                Dim NumberSID As Int32 = 0
                Dim DummySpace As Int32 = 0
                For Each item As String In StartSampleID
                    If (item.Trim() <> "") Then
                        NumberSID += 1
                    Else
                        DummySpace += 1
                    End If
                    If DummySpace = 4 Then
                        Exit For
                    End If
                Next
                SampleID = DeviceStr.Substring(14, NumberSID)
                Dim LastData As String = DeviceStr.Substring(14 + NumberSID)

                '' Force Logic
                LastData = LastData.Replace(Chr(10), String.Empty).Trim
                Dim SplitDatas() As String = LastData.Split(New [Char]() {"r"c, "L"c, "H"c})
                If SplitDatas IsNot Nothing AndAlso SplitDatas.Count = 1 Then
                    SplitDatas = LastData.Split("r")
                End If
                If SplitDatas IsNot Nothing AndAlso SplitDatas.Count > 0 Then
                    For Each valueDatas As String In SplitDatas
                        Dim Items() As String = valueDatas.Split(" "c)
                        Dim Values As String = ""
                        For Each Item As String In Items.Where(Function(c) c <> "")
                            Dim Value As String = If(IsNumeric(Item), Item, GetNumbers(Item))
                            Values &= If(Values = "", Value, "_" & Value)
                        Next
                        If Values <> "" Then
                            PatientInformations.Add(Values)
                        End If
                    Next
                End If
            Case Process = "D" And DistinctionCode = "H"
                SampleID = DeviceStr.Substring(10, 4)
            Case Process = "D" And DistinctionCode = "R"
                SampleID = DeviceStr.Substring(10, 4)
            Case Process = "D" And DistinctionCode = "A"
                SampleID = DeviceStr.Substring(10, 4)
            Case Process = "d" And DistinctionCode = " "
                SampleID = DeviceStr.Substring(10, 4)
            Case Process = "d" And DistinctionCode = "H"
                SampleID = DeviceStr.Substring(10, 4)
            Case Process = "D" And DistinctionCode = "Q"
                Dim LastData As String = DeviceStr.Substring(14)
                LastData = If(LastData <> "", LastData.Trim(), "")
                SampleID = LastData.Substring(0, 3)
                LastData = LastData.Substring(4).Replace(Chr(10), String.Empty).Trim
                Dim SplitDatas() As String = LastData.Split(New [Char]() {"r"c, "L"c})
                If SplitDatas IsNot Nothing AndAlso SplitDatas.Count = 1 Then
                    SplitDatas = LastData.Split("r")
                End If
                If SplitDatas IsNot Nothing AndAlso SplitDatas.Count > 0 Then
                    For Each valueDatas As String In SplitDatas
                        Dim Items() As String = valueDatas.Split(" "c)
                        Dim Values As String = ""
                        For Each Item As String In Items.Where(Function(c) c <> "")
                            Dim Value As String = If(IsNumeric(Item), Item, GetNumbers(Item))
                            Values &= If(Values = "", Value, "_" & Value)
                        Next
                        If Values <> "" Then
                            PatientInformations.Add(Values)
                        End If
                    Next
                End If
            Case Process = "D" And DistinctionCode = "E"
            Case Else
        End Select
        AppSetting.WriteErrorLog(comPort, "Error", "Get SampleID1->" & SampleID & "<-")
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        Try
            If SampleID <> "" And PatientInformations IsNot Nothing AndAlso PatientInformations.Count > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "Create data in loop")
                For Each UpdateValues As String In PatientInformations
                    If UpdateValues = "" Then
                        Continue For
                    End If
                    Dim UpdateSplit() As String = UpdateValues.Split("_")
                    UnivTestID = UpdateSplit(0)
                    Dim UpdateValue As String = UpdateSplit(1)
                    Dim Dilution As String = ""
                    Dim Dilution_flag As String = "Y"

                    AppSetting.WriteErrorLog(comPort, "Error", "Get SampleID->" & SampleID & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get UnivTestID->" & UnivTestID & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get UpdateValue->" & UpdateValue & "<-")
                    dtTestResult = GetResultCodeDilution(SampleID, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, UnivTestID, "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & UnivTestID & "'"
                    If dvTestResult.Count >= 1 Then
                        For ii = 0 To dvTestResult.Count - 1
                            Dim row As DataRow = dtTestResultToUpdate.NewRow
                            row("order_skey") = dvTestResult.Item(ii)("order_skey")
                            row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                            row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                            row("alias_id") = dvTestResult.Item(ii)("alias_id")
                            row("result_value") = UpdateValue
                            row("order_id") = dvTestResult.Item(ii)("order_id")
                            row("analyzer") = dvTestResult.Item(ii)("analyzer")
                            row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                            row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                            row("result_date") = Now
                            row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                            row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                            row("rerun") = dvTestResult.Item(ii)("rerun")
                            row("sample_type") = dvTestResult.Item(ii)("sample_type")
                            dtTestResultToUpdate.Rows.Add(row)
                        Next
                    End If

                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If


                    dtTestResultToUpdate.Clear()

                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
                Next
            End If
endLine:
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula Before")
            End If

            UpdateFormula(SampleID)
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula end")
            End If
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ExtractResult: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return True
    End Function

    Public Function ExtractResult_UriscanPRO(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(3)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count


        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(2))


            If ItemArr.Count <= 0 Then
                Continue For
            End If
            'Check Date for M/C
            If ItemArr(0).IndexOf("Date") <> -1 Then
                Dim dt As String = ItemArr(0).Substring(6)
                resultDate = DateTime.Parse(dt)
            End If
            If ItemArr(0).IndexOf("Date") <> -1 Or ItemArr(0).IndexOf("Ward") <> -1 Or ItemArr(0).IndexOf("Name") <> -1 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("ID_NO") <> -1 And SpecimenId.Length <= 1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("-")
                SpecimenId = SpecimenIdArr(1).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If i > 3 Then

                Dim assay_code As String = ""
                Dim result As String = ""
                ItemArr(0) = ItemArr(0).Replace("-", "").Replace("+", "").Replace("DK. ", "DK.X").Replace(" ", "^")
                If ItemArr(0).IndexOf("^^^^^") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("^^^^^", "^")
                End If
                If ItemArr(0).IndexOf("^^^^") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("^^^^", "^")
                End If
                If ItemArr(0).IndexOf("^^^") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("^^^", "^")
                End If
                If ItemArr(0).IndexOf("^^") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("^^", "^")
                End If
                If ItemArr(0).IndexOf("DK.X") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK.X", "DK. ")
                End If
                If ItemArr(0).IndexOf("LT. Yellow") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("LT. Yellow", "PALEL YELLOW")
                End If
                If ItemArr(0).IndexOf("Yellow") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Yellow", "YELLOW")
                End If
                If ItemArr(0).IndexOf("DK. Yellow") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK. Yellow", "DEEP YELLOW")
                End If
                If ItemArr(0).IndexOf("Orange") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Orange", "ORANGE")
                End If
                If ItemArr(0).IndexOf("DK. Orange") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK. Orange", "ORANGE")
                End If
                If ItemArr(0).IndexOf("Red") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Red", "RED")
                End If
                If ItemArr(0).IndexOf("DK. Red") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK. Red", "RED")
                End If
                If ItemArr(0).IndexOf("Brown") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Brown", "BROWN")
                End If
                If ItemArr(0).IndexOf("DK. Brown") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK. Brown", "BROWN")
                End If
                If ItemArr(0).IndexOf("Green") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Green", "GREEN")
                End If
                If ItemArr(0).IndexOf("DK. Green") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("DK. Green", "GREEN")
                End If
                If ItemArr(0).IndexOf("Clear") <> -1 Then
                    ItemArr(0) = ItemArr(0).Replace("Clear", "CLEAR")
                End If

                Dim assay_codeArr() As String = ItemArr(0).Split("^")
                assay_code = assay_codeArr(0).Trim
                result = assay_codeArr(1).Trim

                Select Case assay_code
                    Case "URO"
                        Select Case result
                            Case "norm"
                                result = "NEGATIVE"
                            Case "1"
                                result = "1+"
                            Case "4"
                                result = "2+"
                            Case "8"
                                result = "3+"
                            Case "12"
                                result = "4+"
                        End Select
                    Case "BIL"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "0.5"
                                result = "1+"
                            Case "1.0"
                                result = "2+"
                            Case "3.0"
                                result = "3+"
                        End Select
                    Case "KET"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "5"
                                result = "TRACE"
                            Case "10"
                                result = "1+"
                            Case "50"
                                result = "2+"
                            Case "100"
                                result = "3+"
                        End Select
                    Case "BLD"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "5"
                                result = "TRACE"
                            Case "10"
                                result = "1+"
                            Case "50"
                                result = "2+"
                            Case "250"
                                result = "3+"
                        End Select
                    Case "PRO"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "10"
                                result = "TRACE"
                            Case "30"
                                result = "1+"
                            Case "100"
                                result = "2+"
                            Case "300"
                                result = "3+"
                            Case "1000"
                                result = "4+"
                        End Select
                    Case "NIT"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "pos"
                                result = "POSITIVE"
                        End Select
                    Case "LEU"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "10"
                                result = "TRACE"
                            Case "25"
                                result = "1+"
                            Case "75"
                                result = "2+"
                            Case "500"
                                result = "3+"
                        End Select
                    Case "GLU"
                        Select Case result
                            Case "neg"
                                result = "NEGATIVE"
                            Case "100"
                                result = "TRACE"
                            Case "250"
                                result = "1+"
                            Case "500"
                                result = "2+"
                            Case "1000"
                                result = "3+"
                            Case "2000"
                                result = "4+"

                        End Select
                End Select


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & assay_code & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = result.ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If i > 3 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_BACTEC9000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        Dim code As String = ""
        Dim codeArr() As String
        Dim code_test As String = ""
        Dim code_testArr() As String
        Dim result As String = ""
        Dim resultArr() As String
        Dim code_result_skey As String = ""
        Dim code_RM As String = ""

        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If

        Dim i, j As Integer

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count

        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""

            If LineArr(i).IndexOf("R") = 0 Then
                If LineArr(i).IndexOf("F") = -1 Or LineArr(i).IndexOf("CR") = -1 Or LineArr(i).IndexOf(Chr(13)) = -1 Then
                    If LineArr(i).IndexOf(Chr(23)) <> -1 Then
                        Dim check23 As Integer = InStr(1, LineArr(i), Chr(23), 1) - 1
                        LineArr(i) = LineArr(i).Substring(0, check23)
                    End If
                    LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(10), "").Replace("C13", "")
                    Dim testLine = LineArr(i + 1).Replace(Chr(2), "").Replace(Chr(10), "").Replace("C13", "")
                    Dim putLine As String = LineArr(i) + testLine
                    LineArr(i) = putLine
                    LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(2), "").Replace(Chr(10), "").Trim
                End If
            End If

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")

            If ItemArr.Count <= 0 Or ItemArr(0).IndexOf(Chr(3)) <> -1 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)
                SpecimenIdStr = SpecimenIdStr.Replace("||", "|")
                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                Dim SpecimenIdArr2() As String = SpecimenIdArr(2).Split("^")
                If SpecimenIdArr2(0).IndexOf("-") <> -1 Then
                    Dim check_ As Integer = InStr(1, SpecimenIdArr2(0), "-", 1) - 1
                    SpecimenIdArr2(0) = SpecimenIdArr2(0).Substring(0, check_)
                End If
                SpecimenId = SpecimenIdArr2(0).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                If ItemArr(2).IndexOf("ID") <> -1 Then
                    ItemArr(2) = ItemArr(2).Replace("^^^", "").Replace("^^", "^")
                    code_testArr = ItemArr(3).Split("^")
                    codeArr = ItemArr(2).Split("^")
                    code = codeArr(0)
                    code_test = code_testArr(1).Trim()
                    If code_test = "^" Then
                        code_test = code_testArr(2)
                    End If
                    result = ""
                    code_result_skey = ""
                    code_RM = ""


                    If code_testArr(3).Length > 0 Then
                        code_RM = code_testArr(3)
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If
                    If code_testArr(4).Length > 0 Then
                        code_RM = code_testArr(4)
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If
                    If code_testArr(5).Length > 0 Then
                        code_RM = code_testArr(5)
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If
                    If code_testArr(6).Length > 0 Then
                        code_RM = code_testArr(6)
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If
                    If code_testArr(7).Length > 0 Then
                        code_RM = code_testArr(7)
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If
                    If code_RM = "" Then
                        sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                    End If

                End If
                If ItemArr(2).IndexOf("AST") <> -1 Or ItemArr(2).IndexOf("AST_MIC") <> -1 Then
                    code_testArr = ItemArr(2).Split("^^^", "", "^^", "")
                    code = code_testArr(3).Trim()
                    code_test = code_testArr(5).Replace("2", "").Replace("AC", "").Trim()
                    resultArr = ItemArr(3).Split("^")
                    result = resultArr(1).Replace("C13", "").Replace("B83", "").Trim
                    code_result_skey = resultArr(2).Trim()
                    code_RM = ""
                    sendData(SpecimenId, code, code_test, result, code_result_skey, code_RM)
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If

                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If


            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_Rapidlab348EX100(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Bloodgas")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** not found EOT then return *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView
        Dim ReadStatus As Boolean = False

        If DeviceStr.ToLower.Contains("qc") Then
            ExtractResultQC_Rapidlab348EX100(DeviceStr)
            Return True
        End If

        LineArr = DeviceStr.Split(Chr(10)) '//<FS>

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""

        'find specimenid
        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space
            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 0 Then
                If ItemArr(0).ToString.ToLower = "operator" Then
                    'SpecimenId = ItemArr(2).Trim
                    'If ItemArr.Length > 3 Then
                    '    SpecimenId = ItemArr(3).Trim
                    'End If

                    For kk As Integer = 2 To ItemArr.Length - 1
                        If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                            SpecimenId = ItemArr(kk).Trim
                            Exit For
                        End If
                    Next

                    If SpecimenId.Trim <> "" Then
                        dtTestResult = New DataTable
                        dtTestResult = GetResultCode(SpecimenId.Trim, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                        dvTestResult = New DataView(dtTestResult)
                        dtTestResultToUpdate = dtTestResult.Clone()
                    End If
                End If
            End If
        Next
        'end find specimenid

        For i = 0 To j - 1

            ''//case reference range
            'If i > 0 Then
            '    Dim refstr As String = LineArr(i - 1)
            '    Dim refarr() As String = refstr.Split(Chr(32))
            '    If refarr.Length > 0 Then
            '        If refarr(0).Trim.ToLower = "reference" Then Continue For
            '    End If
            'End If
            ''//end case reference range

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space

            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 1 Then

                If ItemArr(0).ToString.ToLower = "measured" Or ItemArr(0).ToString.ToLower = "calculated" Then
                    ReadStatus = True
                End If
                If ItemArr(0).ToString.ToLower = "reference" Then
                    ReadStatus = False
                End If

                If ReadStatus = True Then
                    Dim ref_cd As String = ""
                    Dim ref_value As String = ""
                    Dim k As Integer = 0
                    For Each Arr In ItemArr 'hardcode found first value => code, found value two => result value
                        If Not String.IsNullOrEmpty(Arr) Then
                            k += 1
                            If k = 1 Then
                                ref_cd = Arr
                            End If
                            If k = 2 Then
                                If Arr.Length > 0 Then
                                    ref_value = Arr
                                    ref_value = ref_value.Replace(Chr(118), String.Empty)   'remove value "v"
                                    ref_value = ref_value.Replace(Chr(94), String.Empty)    'remove value "^"
                                End If
                            End If
                        End If
                    Next
                    If Not String.IsNullOrEmpty(ref_value) Then
                        dvTestResult.RowFilter = "analyzer_ref_cd = '" & ref_cd & "'" 'analyzer_ref_cd
                        If dvTestResult.Count >= 1 Then
                            For ii = 0 To dvTestResult.Count - 1
                                Dim row As DataRow = dtTestResultToUpdate.NewRow
                                row("order_skey") = dvTestResult.Item(ii)("order_skey")
                                row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                                row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                                row("alias_id") = dvTestResult.Item(ii)("alias_id")
                                row("result_value") = ref_value 'value
                                row("order_id") = dvTestResult.Item(ii)("order_id")
                                row("analyzer") = dvTestResult.Item(ii)("analyzer")
                                row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                                row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                                row("result_date") = resultDate
                                row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                                row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                                row("rerun") = dvTestResult.Item(ii)("rerun")
                                row("sample_type") = dvTestResult.Item(ii)("sample_type")
                                dtTestResultToUpdate.Rows.Add(row)
                            Next
                        End If
                    End If

                End If

            End If
endLine:
        Next

        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue1")
        End If
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Function ExtractResult_Rapidlab348EX132(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Rapidlab348EX 1.32")
        End If
        If DeviceStr.Length <= 0 Then
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView
        Dim ReadStatus As Boolean = False

        If DeviceStr.ToLower.Contains("qc") Then
            ExtractResultQC_Rapidlab348EX132(DeviceStr)
            Return True
        End If

        LineArr = DeviceStr.Split(Chr(10)) '//<FS>

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""

        'find specimenid
        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space
            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 0 Then
                If ItemArr(0).ToString.ToLower = "operator" Then
                    For kk As Integer = 2 To ItemArr.Length - 1
                        If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                            SpecimenId = ItemArr(kk).Trim
                            Exit For
                        End If
                    Next
                    'SpecimenId = ItemArr(2).Trim
                    'If ItemArr.Length > 3 Then
                    '    SpecimenId = ItemArr(3).Trim
                    'End If
                    If SpecimenId.Trim <> "" Then
                        dtTestResult = New DataTable
                        dtTestResult = GetResultCode(SpecimenId.Trim, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                        dvTestResult = New DataView(dtTestResult)
                        dtTestResultToUpdate = dtTestResult.Clone()
                    End If
                End If
            End If
        Next
        'end find specimenid

        For i = 0 To j - 1

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space

            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 1 Then

                If ItemArr(0).ToString.ToLower = "measured" Or ItemArr(0).ToString.ToLower = "calculated" Then
                    ReadStatus = True
                End If
                If ItemArr(0).ToString.ToLower = "reference" Then
                    ReadStatus = False
                End If

                If ReadStatus = True Then
                    Dim ref_cd As String = ""
                    Dim ref_value As String = ""
                    Dim k As Integer = 0
                    For Each Arr In ItemArr 'hardcode found first value => code, found value two => result value
                        If Not String.IsNullOrEmpty(Arr) Then
                            k += 1
                            If k = 1 Then
                                ref_cd = Arr
                            End If
                            If k = 2 Then
                                If Arr.Length > 0 Then
                                    ref_value = Arr
                                    ref_value = ref_value.Replace(Chr(118), String.Empty)   'remove value "v"
                                    ref_value = ref_value.Replace(Chr(94), String.Empty)    'remove value "^"
                                End If
                            End If
                        End If
                    Next
                    If Not String.IsNullOrEmpty(ref_value) Then
                        dvTestResult.RowFilter = "analyzer_ref_cd = '" & ref_cd & "'" 'analyzer_ref_cd
                        If dvTestResult.Count >= 1 Then
                            For ii = 0 To dvTestResult.Count - 1
                                Dim row As DataRow = dtTestResultToUpdate.NewRow
                                row("order_skey") = dvTestResult.Item(ii)("order_skey")
                                row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                                row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                                row("alias_id") = dvTestResult.Item(ii)("alias_id")
                                row("result_value") = ref_value 'value
                                row("order_id") = dvTestResult.Item(ii)("order_id")
                                row("analyzer") = dvTestResult.Item(ii)("analyzer")
                                row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                                row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                                row("result_date") = resultDate
                                row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                                row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                                row("rerun") = dvTestResult.Item(ii)("rerun")
                                row("sample_type") = dvTestResult.Item(ii)("sample_type")
                                dtTestResultToUpdate.Rows.Add(row)
                            Next
                        End If
                    End If

                End If

            End If
endLine:
        Next

        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue1")
        End If
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Function ExtractResult_bloodgas(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Bloodgas")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** not found EOT then return *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView
        Dim ReadStatus As Boolean = False

        LineArr = DeviceStr.Split(Chr(10)) '//<FS>

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""

        'find specimenid
        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space
            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 0 Then
                If ItemArr(0).ToString.ToLower = "operator" Then
                    SpecimenId = ItemArr(2).Trim
                    If ItemArr.Length > 3 Then
                        SpecimenId = ItemArr(3).Trim
                    End If
                    If SpecimenId.Trim <> "" Then
                        dtTestResult = New DataTable
                        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                        dvTestResult = New DataView(dtTestResult)
                        dtTestResultToUpdate = dtTestResult.Clone()
                    End If
                End If
            End If
        Next
        'end find specimenid

        For i = 0 To j - 1

            ''//case reference range
            'If i > 0 Then
            '    Dim refstr As String = LineArr(i - 1)
            '    Dim refarr() As String = refstr.Split(Chr(32))
            '    If refarr.Length > 0 Then
            '        If refarr(0).Trim.ToLower = "reference" Then Continue For
            '    End If
            'End If
            ''//end case reference range

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split(Chr(32)) '//space

            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr.Length > 1 Then

                If ItemArr(0).ToString.ToLower = "measured" Or ItemArr(0).ToString.ToLower = "calculated" Then
                    ReadStatus = True
                End If
                If ItemArr(0).ToString.ToLower = "reference" Then
                    ReadStatus = False
                End If

                If ReadStatus = True Then
                    Dim ref_cd As String = ""
                    Dim ref_value As String = ""
                    Dim k As Integer = 0
                    For Each Arr In ItemArr 'hardcode found first value => code, found value two => result value
                        If Not String.IsNullOrEmpty(Arr) Then
                            k += 1
                            If k = 1 Then
                                ref_cd = Arr
                            End If
                            If k = 2 Then
                                If Not Arr.IndexOf("-") = 0 Then
                                    ref_value = Arr.Replace(Chr(94), Space(1)).Trim
                                End If
                            End If
                        End If
                    Next
                    If Not String.IsNullOrEmpty(ref_value) Then
                        dvTestResult.RowFilter = "analyzer_ref_cd = '" & ref_cd & "'" 'analyzer_ref_cd
                        If dvTestResult.Count >= 1 Then
                            For ii = 0 To dvTestResult.Count - 1
                                Dim row As DataRow = dtTestResultToUpdate.NewRow
                                row("order_skey") = dvTestResult.Item(ii)("order_skey")
                                row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                                row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                                row("alias_id") = dvTestResult.Item(ii)("alias_id")
                                row("result_value") = ref_value 'value
                                row("order_id") = dvTestResult.Item(ii)("order_id")
                                row("analyzer") = dvTestResult.Item(ii)("analyzer")
                                row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                                row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                                row("result_date") = resultDate
                                row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                                row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                                row("rerun") = dvTestResult.Item(ii)("rerun")
                                row("sample_type") = dvTestResult.Item(ii)("sample_type")
                                dtTestResultToUpdate.Rows.Add(row)
                            Next
                        End If
                    End If

                End If

            End If
endLine:
        Next

        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue1")
        End If
        UpdateTestResultValue(dtTestResultToUpdate, False)
        If My.Settings.TrackError Then
            log.Info(comPort & " UpdateTestResultValue2")
        End If

        dtTestResultToUpdate.Clear()
        dtTestResult.Dispose()
        dtTestResultToUpdate.Dispose()

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Sub ExtractResultQC_Rapidlab348EX100(DeviceStr As String)
        Try
            'find QC
            Dim dr As DataRow
            Dim LineArr() As String
            Dim ItemStr As String
            Dim ItemArr() As String
            Dim j As Integer
            Dim level As String = ""
            Dim ref_lot_no As String = ""
            Dim SpecimenId As String
            Dim qcDate As DateTime
            Dim strTime As String = ""
            Dim strDate As String = ""
            Dim dvTestResult As New DataView
            Dim ReadStatus As Boolean = False
            Dim strQC As String()
            Dim dtLot As New DataTable
            dtLot.Columns.Add("seq", GetType(Integer))
            dtLot.Columns.Add("specimen_id", GetType(String))
            dtLot.Columns.Add("ref_cd", GetType(String))
            dtLot.Columns.Add("ref_value", GetType(String))
            dtLot.Columns.Add("qc_date", GetType(DateTime))

            Dim dtSeq As New DataTable
            dtSeq.Columns.Add("seq", GetType(Integer))
            dtSeq.Columns.Add("specimen_id", GetType(String))

            Dim seq As Integer

            strQC = DeviceStr.Split(Chr(4)) 'split EOT
            For Each str As String In strQC

                LineArr = str.Split(Chr(10)) 'split enter

                If str.ToLower.Contains("qc") Then

                    For i = 0 To LineArr.Count - 1

                        ItemStr = LineArr(i)
                        ItemArr = ItemStr.Split(Chr(32)) '//space
                        If ItemArr.Count <= 0 Then
                            Continue For
                        End If

                        If ItemArr.Length > 0 Then

                            If ItemArr.Count >= 2 Then
                                If ItemArr(0).ToString.Trim.ToLower = "qc" And ItemArr(1).ToString.Trim.ToLower = "ranges" Then
                                    ReadStatus = False
                                End If
                            End If


                            If ReadStatus = True Then
                                If Not String.IsNullOrEmpty(SpecimenId) Then
                                    Dim ref_cd As String = ""
                                    Dim ref_value As String = ""
                                    Dim k As Integer = 0
                                    For Each Arr In ItemArr 'hardcode found first value => code, found value two => result value
                                        If Not String.IsNullOrEmpty(Arr) Then
                                            k += 1
                                            If k = 1 Then
                                                ref_cd = Arr
                                            End If
                                            If k = 2 Then
                                                If Arr.Length > 0 Then
                                                    ref_value = Arr
                                                    ref_value = ref_value.Replace(Chr(118), String.Empty)   'remove value "v"
                                                    ref_value = ref_value.Replace(Chr(94), String.Empty)    'remove value "^"
                                                End If
                                            End If
                                        End If
                                    Next

                                    seq = (From row In dtSeq.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()

                                    If Not String.IsNullOrEmpty(ref_value) Then
                                        dr = dtLot.Rows.Add
                                        dr("seq") = seq
                                        dr("specimen_id") = SpecimenId
                                        dr("ref_cd") = ref_cd
                                        dr("ref_value") = ref_value
                                        dr("qc_date") = qcDate
                                    End If
                                End If
                            End If

                            If ItemArr(0).ToString.ToLower.Contains("348") Then 'get qc date
                                Dim l As Integer = 0
                                For kk As Integer = 1 To ItemArr.Length - 1
                                    If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                        l += 1
                                        If l = 1 Then
                                            strTime = ItemArr(kk)
                                        End If
                                        If l >= 2 Then
                                            strDate = strDate & ItemArr(kk)
                                        End If
                                    End If
                                Next
                            End If
                            If ItemArr(0).ToString.ToLower = "level" Then 'get level
                                For kk As Integer = 1 To ItemArr.Length - 1
                                    If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                        level = ItemArr(kk).Trim
                                        Exit For
                                    End If
                                Next
                            End If
                            If ItemArr(0).ToString.ToLower = "lot" Then '//found lot no
                                For kk As Integer = 1 To ItemArr.Length - 1
                                    If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                        ref_lot_no = ItemArr(kk).Trim
                                        Exit For
                                    End If
                                Next
                                SpecimenId = ref_lot_no & "-" & level
                                qcDate = Date.Parse(strDate & " " & strTime)
                                ReadStatus = True
                                If dtSeq.Select("specimen_id = '" & SpecimenId & "'").Count = 0 Then
                                    dtSeq.Rows.Add({1, SpecimenId})
                                Else
                                    seq = (From row In dtSeq.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()
                                    If Not seq = Nothing Then
                                        dtSeq.Rows.Add({seq + 1, SpecimenId})
                                    End If
                                End If
                            End If
                        End If 'If ItemArr.Length > 0 Then
                    Next 'For i = 0 To LineArr.Count - 1
                End If 'If str.ToLower.Contains("qc") Then
            Next


            For Each dr In dtLot.DefaultView.ToTable(True, "specimen_id").Rows
                SpecimenId = IIf(dr("specimen_id") Is DBNull.Value, "", dr("specimen_id"))
                If Not String.IsNullOrEmpty(SpecimenId) Then
                    dtTestResult = New DataTable
                    dtTestResult = GetResultCode(SpecimenId.Trim, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = New DataTable
                    dtTestResultToUpdate = dtTestResult.Clone()

                    seq = (From row In dtLot.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()

                    For Each r As DataRow In dtLot.Select("specimen_id = '" & SpecimenId & "' and seq = " & seq.ToString)
                        If IsNumeric(r("ref_value")) Then
                            dvTestResult.RowFilter = "analyzer_ref_cd = '" & r("ref_cd") & "'" 'analyzer_ref_cd
                            If dvTestResult.Count >= 1 Then
                                For ii = 0 To dvTestResult.Count - 1
                                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                                    row("order_skey") = dvTestResult.Item(ii)("order_skey")
                                    row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                                    row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                                    row("alias_id") = dvTestResult.Item(ii)("alias_id")
                                    row("result_value") = r("ref_value") 'value
                                    row("order_id") = dvTestResult.Item(ii)("order_id")
                                    row("analyzer") = dvTestResult.Item(ii)("analyzer")
                                    row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                                    row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                                    row("result_date") = r("qc_date")
                                    row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                                    row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                                    row("rerun") = dvTestResult.Item(ii)("rerun")
                                    row("sample_type") = dvTestResult.Item(ii)("sample_type")
                                    dtTestResultToUpdate.Rows.Add(row)
                                Next
                            End If
                        End If
                    Next

                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, True)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If

                    'Tui add 2015-09-22  auto update formula
                    If My.Settings.TrackError Then
                        log.Info(analyzerModel & "UpdateFormula Before")
                    End If
                    UpdateFormula(SpecimenId)
                    If My.Settings.TrackError Then
                        log.Info(analyzerModel & "UpdateFormula end")
                    End If

                End If 'If Not String.IsNullOrEmpty(SpecimenId) Then
            Next 'For Each dr In dtLot.DefaultView.ToTable(True, "specimen_id").Rows

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub ExtractResultQC_Rapidlab348EX132(DeviceStr As String)
        Try
            'find QC
            Dim dr As DataRow
            Dim LineArr() As String
            Dim ItemStr As String
            Dim ItemArr() As String
            Dim j As Integer
            Dim level As String = ""
            Dim ref_lot_no As String = ""
            Dim SpecimenId As String
            Dim qcDate As DateTime
            Dim strTime As String = ""
            Dim strDate As String = ""
            Dim dvTestResult As New DataView
            Dim ReadStatus As Boolean = False
            'Dim strQC As String()
            Dim dtLot As New DataTable
            dtLot.Columns.Add("seq", GetType(Integer))
            dtLot.Columns.Add("specimen_id", GetType(String))
            dtLot.Columns.Add("ref_cd", GetType(String))
            dtLot.Columns.Add("ref_value", GetType(String))
            dtLot.Columns.Add("qc_date", GetType(DateTime))

            Dim dtSeq As New DataTable
            dtSeq.Columns.Add("seq", GetType(Integer))
            dtSeq.Columns.Add("specimen_id", GetType(String))

            Dim seq As Integer

            'strQC = DeviceStr.Split(Chr(4)) 'split EOT
            'For Each str As String In strQC

            LineArr = DeviceStr.Split(Chr(10)) 'split enter

            If DeviceStr.ToLower.Contains("qc") Then

                For i = 0 To LineArr.Count - 1

                    ItemStr = LineArr(i)
                    ItemArr = ItemStr.Split(Chr(32)) '//space
                    If ItemArr.Count <= 0 Then
                        Continue For
                    End If

                    If ItemArr.Length > 0 Then

                        If ItemArr.Count >= 2 Then
                            If ItemArr(0).ToString.Trim.ToLower = "qc" And ItemArr(1).ToString.Trim.ToLower = "ranges" Then
                                ReadStatus = False
                            End If
                        End If


                        If ReadStatus = True Then
                            If Not String.IsNullOrEmpty(SpecimenId) Then
                                Dim ref_cd As String = ""
                                Dim ref_value As String = ""
                                Dim k As Integer = 0
                                For Each Arr In ItemArr 'hardcode found first value => code, found value two => result value
                                    If Not String.IsNullOrEmpty(Arr) Then
                                        k += 1
                                        If k = 1 Then
                                            ref_cd = Arr
                                        End If
                                        If k = 2 Then
                                            If Arr.Length > 0 Then
                                                ref_value = Arr
                                                ref_value = ref_value.Replace(Chr(118), String.Empty)   'remove value "v"
                                                ref_value = ref_value.Replace(Chr(94), String.Empty)    'remove value "^"
                                            End If
                                        End If
                                    End If
                                Next

                                seq = (From row In dtSeq.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()

                                If Not String.IsNullOrEmpty(ref_value) Then
                                    dr = dtLot.Rows.Add
                                    dr("seq") = seq
                                    dr("specimen_id") = SpecimenId
                                    dr("ref_cd") = ref_cd
                                    dr("ref_value") = ref_value
                                    dr("qc_date") = qcDate
                                End If
                            End If
                        End If

                        If ItemArr(0).ToString.ToLower.Contains("348") Then 'get qc date
                            Dim l As Integer = 0
                            For kk As Integer = 1 To ItemArr.Length - 1
                                If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                    l += 1
                                    If l = 1 Then
                                        strTime = ItemArr(kk)
                                    End If
                                    If l >= 2 Then
                                        strDate = strDate & ItemArr(kk)
                                    End If
                                End If
                            Next
                        End If
                        If ItemArr(0).ToString.ToLower = "level" Then 'get level
                            For kk As Integer = 1 To ItemArr.Length - 1
                                If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                    level = ItemArr(kk).Trim
                                    Exit For
                                End If
                            Next
                        End If
                        If ItemArr(0).ToString.ToLower = "lot" Then '//found lot no
                            For kk As Integer = 1 To ItemArr.Length - 1
                                If Not String.IsNullOrEmpty(ItemArr(kk)) Then
                                    ref_lot_no = ItemArr(kk).Trim
                                    Exit For
                                End If
                            Next
                            SpecimenId = ref_lot_no & "-" & level
                            qcDate = Date.Parse(strDate & " " & strTime)
                            ReadStatus = True
                            If dtSeq.Select("specimen_id = '" & SpecimenId & "'").Count = 0 Then
                                dtSeq.Rows.Add({1, SpecimenId})
                            Else
                                seq = (From row In dtSeq.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()
                                If Not seq = Nothing Then
                                    dtSeq.Rows.Add({seq + 1, SpecimenId})
                                End If
                            End If
                        End If
                    End If 'If ItemArr.Length > 0 Then
                Next 'For i = 0 To LineArr.Count - 1
            End If 'If str.ToLower.Contains("qc") Then
            'Next


            For Each dr In dtLot.DefaultView.ToTable(True, "specimen_id").Rows
                SpecimenId = IIf(dr("specimen_id") Is DBNull.Value, "", dr("specimen_id"))
                If Not String.IsNullOrEmpty(SpecimenId) Then
                    dtTestResult = New DataTable
                    dtTestResult = GetResultCode(SpecimenId.Trim, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = New DataTable
                    dtTestResultToUpdate = dtTestResult.Clone()

                    seq = (From row In dtLot.Select("specimen_id = '" & SpecimenId & "'").AsEnumerable() Select _seq = row.Field(Of Integer)("seq") Order By _seq Descending).FirstOrDefault()

                    For Each r As DataRow In dtLot.Select("specimen_id = '" & SpecimenId & "' and seq = " & seq.ToString)
                        If IsNumeric(r("ref_value")) Then
                            dvTestResult.RowFilter = "analyzer_ref_cd = '" & r("ref_cd") & "'" 'analyzer_ref_cd
                            If dvTestResult.Count >= 1 Then
                                For ii = 0 To dvTestResult.Count - 1
                                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                                    row("order_skey") = dvTestResult.Item(ii)("order_skey")
                                    row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                                    row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                                    row("alias_id") = dvTestResult.Item(ii)("alias_id")
                                    row("result_value") = r("ref_value") 'value
                                    row("order_id") = dvTestResult.Item(ii)("order_id")
                                    row("analyzer") = dvTestResult.Item(ii)("analyzer")
                                    row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                                    row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                                    row("result_date") = r("qc_date")
                                    row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                                    row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                                    row("rerun") = dvTestResult.Item(ii)("rerun")
                                    row("sample_type") = dvTestResult.Item(ii)("sample_type")
                                    dtTestResultToUpdate.Rows.Add(row)
                                Next
                            End If
                        End If
                    Next

                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, True)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If

                    'Tui add 2015-09-22  auto update formula
                    If My.Settings.TrackError Then
                        log.Info(analyzerModel & "UpdateFormula Before")
                    End If
                    UpdateFormula(SpecimenId)
                    If My.Settings.TrackError Then
                        log.Info(analyzerModel & "UpdateFormula end")
                    End If

                End If 'If Not String.IsNullOrEmpty(SpecimenId) Then
            Next 'For Each dr In dtLot.DefaultView.ToTable(True, "specimen_id").Rows

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Function ExtractResult_D10(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        'Dim itemArrOrd() As String
        'Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_D10")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView
        '''''''''''''''''''
        Dim InstSpecID As String
        Dim TestOrderUnivTestID As String
        Dim analyzer_ref_cd As String = ""
        Dim RptType As String

        LineArr = DeviceStr.Split(Chr(10))
        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        Try


            For i = 0 To j - 1
                Dim Dilution As String = ""
                Dim Dilution_flag As String = "Y"
                ItemStr = LineArr(i)
                ItemArr = ItemStr.Split("|")

                If ItemArr.Count <= 0 Or ItemArr(0) = "" Or ItemArr(0).Length < 3 Then
                    Continue For
                End If

                If ItemArr(0).Substring(2, 1) = "O" Then
                    If Not ItemArr(2) Is Nothing And Not ItemArr(3) Is Nothing And Not ItemArr(4) Is Nothing And Not ItemArr(ItemArr.Length - 1) Is Nothing Then
                        SpecimenId = ItemArr(2)
                        InstSpecID = ItemArr(3)
                        TestOrderUnivTestID = ItemArr(4)
                        analyzer_ref_cd = TestOrderUnivTestID.Chars(3)
                        RptType = ItemArr(ItemArr.Length - 1)
                    End If
                End If
                'LabNo Line
                If ItemArr(0).Substring(2, 1) = "R" And SpecimenId <> "" Then 'If ItemArr(0).IndexOf("R") <> -1 Then
                    ItemArr(3) = ItemArr(3).Trim 'Result
                    Dim UnivTestID As String = ""
                    Dim UnivTestType As String = ""
                    Dim ResultUnivTest As String = ""
                    If ItemArr(2).IndexOf("^") <> -1 Then
                        Dim assay_codeArr() As String = ItemArr(2).Split("^")
                        UnivTestID = assay_codeArr(3).Trim
                        UnivTestType = assay_codeArr(4).Trim
                    End If

                    If Not ItemArr(3) Is Nothing Then
                        ResultUnivTest = ItemArr(3)
                    End If

                    If UnivTestType <> "AREA" Or UnivTestID.ToLower() <> "a1c" Then
                        Continue For
                    End If

                    Dim dt As String = ""
                    Dim resultDate = Today
                    Try
                        dt = Mid(ItemArr(12), 1, 14)
                        dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                        resultDate = DateTime.Parse(dt)

                        'resultDate = DateTime.ParseExact(dateText, formatString, Globalization.CultureInfo.InvariantCulture)
                    Catch ex As Exception
                        AppSetting.WriteErrorLog(comPort, "Error", "Error Get IQC Date" & ex.Message & "; Get source data: " & dt)
                        log.Error(ex.Message)
                    End Try

                    AppSetting.WriteErrorLog(comPort, "Error", "Get ItemArr(3)->" & ItemArr(3) & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get ItemArr(12)->" & ItemArr(12) & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get dt->" & dt & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get SpecimenId->" & SpecimenId & "<-")
                    AppSetting.WriteErrorLog(comPort, "Error", "Get UnivTestID->" & UnivTestID & "<-")
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, UnivTestID, "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                    dvTestResult.RowFilter = "analyzer_ref_cd = '" & UnivTestID & "'"
                    If dvTestResult.Count >= 1 Then
                        For ii = 0 To dvTestResult.Count - 1
                            Dim row As DataRow = dtTestResultToUpdate.NewRow
                            row("order_skey") = dvTestResult.Item(ii)("order_skey")
                            row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                            row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                            row("alias_id") = dvTestResult.Item(ii)("alias_id")
                            row("result_value") = ItemArr(3).ToString()
                            row("order_id") = dvTestResult.Item(ii)("order_id")
                            row("analyzer") = dvTestResult.Item(ii)("analyzer")
                            row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                            row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                            row("result_date") = resultDate
                            row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                            row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                            row("rerun") = dvTestResult.Item(ii)("rerun")
                            row("sample_type") = dvTestResult.Item(ii)("sample_type")
                            dtTestResultToUpdate.Rows.Add(row)
                        Next
                    End If

                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If

                    dtTestResultToUpdate.Clear()

                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()

                End If
endLine:
            Next

            'Tui add 2015-09-22  auto update formula
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula Before")
            End If

            UpdateFormula(SpecimenId)
            If My.Settings.TrackError Then
                log.Info(analyzerModel & "UpdateFormula end")
            End If
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ExtractResult: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return True
    End Function
    Public Function ReturnOrderDetailD10(ByVal SpecimenId As String, ByVal UniversalTestID As String, ByVal Process As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim dtPatient As New DataTable
        Try
            Select Case Process
                Case "H"
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host^10^1.0|||||||||" 'Chr(2) = STX (start-of-text)
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                Case "Q"
                    Str = Chr(2) & CStr(Running) & "Q|1|"
                    Str &= "^"
                    Str &= AppSetting.chkDBNull(SpecimenId.ToString)
                    Str &= "||"
                    Str &= UniversalTestID ' 4 (A1c), 1 (A2F)
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                Case "L"
                    Str = Chr(2) & CStr(Running) & "L|1|N"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                Case Else
            End Select

            AppSetting.WriteErrorLog(comPort, "Send", "Send to D10.")
            AppSetting.WriteErrorLog(comPort, "D10", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function

    Public Function CalculateParameterTypeCA620(ByVal Type As String, ByVal Value As String) As Decimal
        Dim IntParameter As Integer
        If Integer.TryParse(Value, IntParameter) Then
            Dim Data As Decimal
            Select Case Type
                Case "1"
                    'XXXX.X XXXXX
                    Data = IntParameter * 0.1
                Case "3"
                    'XX.XX OXXXX
                    Data = IntParameter * 0.01
                Case "4"
                    'XX.XX OXXXX
                    Data = IntParameter * 0.01
                Case Else
            End Select
            Return Data
        End If

        Return 0
    End Function
    Public Function ExtractResult_CA620(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        'Dim ItemArr() As String
        'Dim itemArrOrd() As String
        'Dim itemArrRes() As String
        'Dim SpecimenId As String = ""
        'Dim SpecimenId_Pre As String = ""
        'Dim Sex As String = ""
        'Dim firstPat As Boolean = True
        'Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_CA620")
        End If
        'If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
        '    Return False
        'End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        '''''''''''''''''''
        Dim InstSpecID As String = ""
        Dim TextDistinctionCodeI As String = ""
        Dim TextDistinctionCodeII As String = ""
        Dim TextDistinctionCodeIII As String = ""
        Dim BlockNumber As String = ""
        Dim TotalNumberOfBlocks As String = ""
        Dim SampleDistinctionCode As String = ""
        Dim ItemDateTime As String = ""
        Dim Time As String = ""
        Dim RackNumber As String = ""
        Dim TubePositionNumber As String = ""
        Dim SampleIDNumber As String = ""
        Dim IDInformation As String = ""
        Dim PatientName As String = ""
        Dim ResultData As String = ""


        LineArr = DeviceStr.Split(Chr(3))
        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = "Y"
            ItemStr = LineArr(i)

            If LineArr.Count <= 0 Or ItemStr = "" Then
                Continue For
            End If
            Dim RemoveIndex = If(i = 0, 1, 2)
            ItemStr = ItemStr.Remove(0, RemoveIndex)
            TextDistinctionCodeI = ItemStr.Substring(0, 1)
            TextDistinctionCodeII = ItemStr.Substring(1, 1)
            TextDistinctionCodeIII = ItemStr.Substring(2, 2)
            BlockNumber = ItemStr.Substring(4, 2)
            TotalNumberOfBlocks = ItemStr.Substring(6, 2)
            SampleDistinctionCode = ItemStr.Substring(8, 1)
            ItemDateTime = ItemStr.Substring(9, 6)
            Time = ItemStr.Substring(15, 4)
            RackNumber = ItemStr.Substring(19, 4)
            TubePositionNumber = ItemStr.Substring(23, 2)
            Dim ItemStrData As String = ItemStr.Substring(25)
            Dim IndexResultSID As String = ""
            If ItemStrData.IndexOf("Q"c) Then
                IndexResultSID = ItemStrData.LastIndexOfAny(New Char() {"M"c, "A"c, "B"c, "C"c})
            Else
                IndexResultSID = ItemStrData.IndexOfAny(New Char() {"M"c, "A"c, "B"c, "C"c})
            End If

            SampleIDNumber = ItemStrData.Substring(0, IndexResultSID)
            IDInformation = ItemStrData.Substring(IndexResultSID, 1)
            PatientName = ItemStrData.Substring(IndexResultSID + 1, 11)
            ResultData = ItemStrData.Substring(IndexResultSID + 12)

            If SampleIDNumber = "" Then
                Continue For
            End If
            SampleIDNumber = SampleIDNumber.Trim()

            If My.Settings.TrackError Then
                log.Info(comPort & "ExtractResult_CA620 =>TextDistinctionCodeI " & TextDistinctionCodeI)
                log.Info(comPort & "ExtractResult_CA620 =>TextDistinctionCodeII " & TextDistinctionCodeII)
                log.Info(comPort & "ExtractResult_CA620 =>TextDistinctionCodeIII " & TextDistinctionCodeIII)
                log.Info(comPort & "ExtractResult_CA620 =>BlockNumber " & BlockNumber)
                log.Info(comPort & "ExtractResult_CA620 =>TotalNumberOfBlocks " & TotalNumberOfBlocks)
                log.Info(comPort & "ExtractResult_CA620 =>SampleDistinctionCode " & SampleDistinctionCode)
                log.Info(comPort & "ExtractResult_CA620 =>ItemDateTime " & ItemDateTime)
                log.Info(comPort & "ExtractResult_CA620 =>Time " & Time)
                log.Info(comPort & "ExtractResult_CA620 =>RackNumber " & RackNumber)
                log.Info(comPort & "ExtractResult_CA620 =>TubePositionNumber " & TubePositionNumber)
                log.Info(comPort & "ExtractResult_CA620 =>SampleIDNumber " & SampleIDNumber)
                log.Info(comPort & "ExtractResult_CA620 =>IDInformation " & IDInformation)
                log.Info(comPort & "ExtractResult_CA620 =>PatientName " & PatientName)
                log.Info(comPort & "ExtractResult_CA620 =>ResultData " & ResultData)
            End If

            Dim resultDate = Today
            Try
                Dim formatString As String = "ddMMyyHHmm"
                resultDate = DateTime.ParseExact(ItemDateTime + Time, formatString, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                log.Error(ex.Message)
            End Try

            If My.Settings.TrackError Then
                log.Info(comPort & "ExtractResult_CA620 => ItemDateTime + Time " & ItemDateTime + Time)
                log.Info(comPort & "ExtractResult_CA620 => " & resultDate.ToString("ddMMyyHHmm"))
            End If

            Dim ResultDataArrayList As ArrayList = New ArrayList()
            For r = 0 To (ResultData.Length / 9) - 1
                Dim SubResult = ResultData.Substring(r * 9, 9)
                Dim FlagResult = SubResult.Last
                If FlagResult = Chr(32) Then
                    Dim Parameter = SubResult.Substring(0, 3)
                    Dim SubResultData = SubResult.Substring(3, 5).Trim()
                    ' Logic
                    Dim IntParameter As Integer
                    If Parameter <> "" And Integer.TryParse(Parameter, IntParameter) And SubResultData <> "" Then
                        Dim ParameterType = Parameter.Last
                        Dim LastSubResultData = CalculateParameterTypeCA620(ParameterType, SubResultData)
                        ResultDataArrayList.Add(Parameter + "_" + LastSubResultData.ToString())
                    End If
                End If
            Next
            If ResultDataArrayList.Count <= 0 Then
                Continue For
            End If

            For Each Item In ResultDataArrayList
                Dim UpdateParameter = Item.Split("_")(0)
                Dim UpdateValue = Item.Split("_")(1)
                dtTestResult = GetResultCodeDilution(SampleIDNumber, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, UpdateParameter, "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()

                'dvTestResult.RowFilter = "alias_id = '" & UnivTestID & "'"
                'dvTestResult.RowFilter = "analyzer_ref_cd = '" & analyzer_ref_cd & "'"
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & UpdateParameter & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = UpdateValue
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If

                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            Next
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SampleIDNumber)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
    Public Function ReturnInquiryDataCA620(ByVal DeviceStr As String) As String
        Dim Result As String = ""
        Try
            Dim LineArr() As String
            Dim ItemStr As String
            If My.Settings.TrackError Then
                log.Info(comPort & "ExtractResult_CA620")
            End If

            Dim i, j As Integer
            Dim dtPatient As New DataTable
            '''''''''''''''''''
            Dim InstSpecID As String = ""
            Dim TextDistinctionCodeI As String = ""
            Dim TextDistinctionCodeII As String = ""
            Dim TextDistinctionCodeIII As String = ""
            Dim BlockNumber As String = ""
            Dim TotalNumberOfBlocks As String = ""
            Dim SampleDistinctionCode As String = ""
            Dim ItemDateTime As String = ""
            Dim Time As String = ""
            Dim RackNumber As String = ""
            Dim TubePositionNumber As String = ""
            Dim SampleIDNumber As String = ""
            Dim IDInformation As String = ""
            Dim PatientName As String = ""
            Dim ResultData As String = ""
            '''''''''''''''''''

            LineArr = DeviceStr.Split(Chr(3))
            j = LineArr.Count
            Dim ItemStrOld As String = ""
            Dim abnormalOld As String = ""
            Dim analyzer_ref_cdOld As String = ""
            Dim SpecimenIdOld As String = ""
            For i = 0 To j - 1
                Dim Dilution As String = ""
                Dim Dilution_flag As String = "Y"
                ItemStr = LineArr(i)

                If LineArr.Count <= 0 Or ItemStr = "" Then
                    Continue For
                End If
                Dim RemoveIndex = If(i = 0, 1, 2)
                ItemStr = ItemStr.Remove(0, RemoveIndex)
                TextDistinctionCodeI = ItemStr.Substring(0, 1)
                TextDistinctionCodeII = ItemStr.Substring(1, 1)
                TextDistinctionCodeIII = ItemStr.Substring(2, 2)
                BlockNumber = ItemStr.Substring(4, 2)
                TotalNumberOfBlocks = ItemStr.Substring(6, 2)
                SampleDistinctionCode = ItemStr.Substring(8, 1)
                ItemDateTime = ItemStr.Substring(9, 6)
                Time = ItemStr.Substring(15, 4)
                RackNumber = ItemStr.Substring(19, 4)
                TubePositionNumber = ItemStr.Substring(23, 2)
                Dim ItemStrData As String = ItemStr.Substring(25)
                Dim IndexResultSID As String = ItemStrData.IndexOfAny(New Char() {"M"c, "A"c, "B"c, "C"c})
                SampleIDNumber = ItemStrData.Substring(0, IndexResultSID)
                IDInformation = ItemStrData.Substring(IndexResultSID, 1)
                PatientName = ItemStrData.Substring(IndexResultSID + 1, 11)
                ResultData = ItemStrData.Substring(IndexResultSID + 12)

                If SampleIDNumber = "" Then
                    Continue For
                End If

                Dim ResultOrInformation As String = Chr(2)
                ResultOrInformation = ResultOrInformation & "S"
                ResultOrInformation = ResultOrInformation & TextDistinctionCodeII
                ResultOrInformation = ResultOrInformation & TextDistinctionCodeIII
                ResultOrInformation = ResultOrInformation & BlockNumber
                ResultOrInformation = ResultOrInformation & TotalNumberOfBlocks
                ResultOrInformation = ResultOrInformation & "U"
                Dim Now As Date = Today
                ResultOrInformation = ResultOrInformation & Now.ToString("ddMMyy")
                ResultOrInformation = ResultOrInformation & Now.ToString("HHmm")
                ResultOrInformation = ResultOrInformation & RackNumber
                ResultOrInformation = ResultOrInformation & TubePositionNumber
                ResultOrInformation = ResultOrInformation & SampleIDNumber
                ResultOrInformation = ResultOrInformation & IDInformation
                ResultOrInformation = ResultOrInformation & PatientName

                SampleIDNumber = SampleIDNumber.Trim()

                AppSetting.WriteErrorLog(comPort, "Send", "AnalyzerInterface: ItemStrData " & ItemStrData)

                'If My.Settings.TrackError Then
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>TextDistinctionCodeI " & TextDistinctionCodeI)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>TextDistinctionCodeII " & TextDistinctionCodeII)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>TextDistinctionCodeIII " & TextDistinctionCodeIII)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>BlockNumber " & BlockNumber)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>TotalNumberOfBlocks " & TotalNumberOfBlocks)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>SampleDistinctionCode " & SampleDistinctionCode)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>ItemDateTime " & ItemDateTime)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>Time " & Time)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>RackNumber " & RackNumber)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>TubePositionNumber " & TubePositionNumber)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>SampleIDNumber " & SampleIDNumber)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>IDInformation " & IDInformation)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>PatientName " & PatientName)
                AppSetting.WriteErrorLog(comPort, "Send", "ExtractResult_CA620 =>ResultData " & ResultData)
                'End If

                Try
                    Dim formatString As String = "ddMMyyHHmm"
                    Dim resultDate = DateTime.ParseExact(ItemDateTime + Time, formatString, Globalization.CultureInfo.InvariantCulture)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try
                Dim ResultDataArrayList As ArrayList = New ArrayList()
                For r = 0 To (ResultData.Length / 9) - 1
                    Dim SubResult = ResultData.Substring(r * 9, 9)
                    Dim FlagResult = SubResult.Last
                    If FlagResult = Chr(32) Then
                        Dim Parameter = SubResult.Substring(0, 3)
                        Dim SubResultData = SubResult.Substring(3, 5).Trim()
                        ' Logic
                        Dim IntParameter As Integer
                        If Parameter <> "" And Integer.TryParse(Parameter, IntParameter) Then
                            'If Parameter <> "" And (Parameter.StartsWith("04") Or Parameter.StartsWith("05")) And Integer.TryParse(Parameter, IntParameter) Then
                            Dim ParameterType = Parameter.Last
                            Dim LastSubResultData = CalculateParameterTypeCA620(ParameterType, SubResultData)
                            ResultDataArrayList.Add(Parameter + "_" + LastSubResultData.ToString())
                        End If
                    End If
                Next
                If ResultDataArrayList.Count <= 0 Then
                    Continue For
                End If
                Dim IsResult = False
                Dim ResultOrInformationData = ""
                For Each Item In ResultDataArrayList
                    Dim Parameter As String = Item.Split("_")(0)
                    'Dim UpdateValue = Item.Split("_")(1)
                    Dim dtTestResult = GetResultCodeDilution(SampleIDNumber, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, Parameter, "N").Copy
                    Dim dvTestResult = New DataView(dtTestResult)
                    dvTestResult.RowFilter = "analyzer_ref_cd LIKE '" & Parameter.Substring(0, 2) & "%'"
                    If dvTestResult.Count >= 1 Then
                        For Each row As DataRowView In dvTestResult
                            Dim value As String = row("analyzer_ref_cd").ToString()
                            If value <> "" Then
                                ResultOrInformationData = ResultOrInformationData & Parameter & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                                IsResult = True
                            Else
                                ResultOrInformationData = ResultOrInformationData & "000" & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                            End If
                        Next
                    Else
                        ResultOrInformationData = ResultOrInformationData & "000" & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                    End If
                    'dtTestResult = GetResultCode(SampleIDNumber, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
                    'Dim dvTestResult = New DataView(dtTestResult)
                    'Dim dtTestResultToUpdate = dtTestResult.Clone()
                Next

                If IsResult Then
                    ResultOrInformation = ResultOrInformation & ResultOrInformationData & Chr(3)
                Else
                    ResultOrInformation = ResultOrInformation & "999" & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32) & Chr(32)
                    ResultOrInformation = ResultOrInformation & Chr(3)
                End If

                Result = Result & ResultOrInformation
endLine:
            Next


            AppSetting.WriteErrorLog(comPort, "Send", "Send to CA620.")

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return Result
    End Function

    Private Class ElyteISE6000Class
        Public Property Key As String
        Public Property Value As String
    End Class
    Public Function ExtractResult_ElyteISE6000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        'Dim ItemArr() As String
        'Dim itemArrOrd() As String
        'Dim itemArrRes() As String
        'Dim SpecimenId As String = ""
        'Dim SpecimenId_Pre As String = ""
        'Dim Sex As String = ""
        'Dim firstPat As Boolean = True
        'Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_ElyteISE6000")
        End If
        'If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
        '    Return False
        'End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        '''''''''''''''''''
        Dim Running As String = ""
        'Dim PATIDText As String = "PAT ID:"
        Dim SID As String = ""
        Dim NumberText As String = ""
        Dim K As String = ""
        Dim Na As String = ""
        Dim Cl As String = ""
        Dim Ca As String = ""
        'Dim PH As String = ""
        Dim CO2 As String = ""
        Dim LastValue As String = ""

        LineArr = DeviceStr.Split(Chr(3))
        j = LineArr.Count
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = "Y"
            ItemStr = LineArr(i)

            If LineArr.Count <= 0 Or ItemStr = "" Then
                Continue For
            End If

            'ItemStr = ItemStr.Replace(PATIDText, String.Empty).Trim

            'Dim RemoveIndex = If(i = 0, 1, 2)
            'ItemStr = ItemStr.Remove(0, RemoveIndex)

            Dim Items() As String = ItemStr.Split(Chr(32))

            'Dim ResultDataArrayList As ArrayList = New ArrayList()
            Dim ElyteISE6000List As New List(Of ElyteISE6000Class)
            Dim index As Integer = 0
            For Each value As String In Items
                If value.Trim <> "" Then
                    Dim ElyteISE6000 As New ElyteISE6000Class()
                    Select Case index
                        Case 0
                            Running = value
                        Case 1
                            SID = Int(value)
                        Case 2
                            NumberText = value
                        Case 3
                            ElyteISE6000.Key = "K"
                            ElyteISE6000.Value = value
                            ElyteISE6000List.Add(ElyteISE6000)
                        Case 4
                            ElyteISE6000.Key = "Na"
                            ElyteISE6000.Value = value
                            ElyteISE6000List.Add(ElyteISE6000)
                        Case 5
                            ElyteISE6000.Key = "Cl"
                            ElyteISE6000.Value = value
                            ElyteISE6000List.Add(ElyteISE6000)
                        Case 6
                            'ElyteISE6000.Key = "Ca"
                            'ElyteISE6000.Value = value
                            'ElyteISE6000List.Add(ElyteISE6000)
                        Case 7
                            'ElyteISE6000.Key = "PH"
                            'ElyteISE6000.Value = value
                            'ElyteISE6000List.Add(ElyteISE6000)
                        Case 8
                            ElyteISE6000.Key = "CO2"
                            ElyteISE6000.Value = value
                            ElyteISE6000List.Add(ElyteISE6000)
                        Case Else
                    End Select
                    index = index + 1
                End If
            Next

            If My.Settings.TrackError Then
                log.Info(comPort & "ExtractResult_ElyteISE6000 =>Running " & Running)
                log.Info(comPort & "ExtractResult_ElyteISE6000 =>NumberText " & NumberText)
                log.Info(comPort & "ExtractResult_ElyteISE6000 =>SID " & SID)
                For Each element In ElyteISE6000List
                    log.Info(comPort & "ExtractResult_ElyteISE6000 =>ElyteISE6000List " & element.Key & "_" & element.Value)
                Next

                log.Info(comPort & "ExtractResult_ElyteISE6000 =>LastValue " & LastValue)
            End If

            'For Each element In ElyteISE6000List
            '    Console.WriteLine("Element: {0}", element)
            'Next

            Try
                'Dim formatString As String = "ddMMyyHHmm"
                Dim resultDate = Today 'DateTime.ParseExact(ItemDateTime + Time, formatString, Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                AppSetting.WriteErrorLog(comPort, "Error", "Get Date" & ex.Message)
                log.Error(ex.Message)
            End Try

            If My.Settings.TrackError Then
                log.Info(comPort & "ExtractResult_ElyteISE6000 => " & resultDate.ToString("dd-MM-yyyy HH:mm:ss"))
            End If

            If ElyteISE6000List.Count <= 0 Or SID = "" Then
                Continue For
            End If

            For Each Item In ElyteISE6000List
                Dim UpdateAnalyzerRefCD = Item.Key
                Dim UpdateValue = Item.Value
                If UpdateAnalyzerRefCD = "" Or UpdateValue = "" Then
                    Continue For
                End If
                dtTestResult = GetResultCodeDilution(SID, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, UpdateAnalyzerRefCD, "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()

                'dvTestResult.RowFilter = "alias_id = '" & UnivTestID & "'"
                'dvTestResult.RowFilter = "analyzer_ref_cd = '" & analyzer_ref_cd & "'"
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & UpdateAnalyzerRefCD & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = UpdateValue
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If

                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            Next

endLine:
        Next

        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SID)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailElitePro(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Dim codeTest As String = ""
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)
                SpecimenId = SpecimenIdArr(1).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim
            End If



            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|" & "\" & "^&||||||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|\^&||||||||ACL9000" 'Chr(2) = STX
            Str &= "||P|1|"
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)

            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            If dtTestResult.Rows.Count >= 0 Then
                Str = Chr(2) & CStr(Running) & "P|1||||"
                Str &= "^^^^"
                Str &= "|||||||||||||||||||||||||||||"
                Str &= Chr(13) & Chr(3)
                Str &= GetCheckSumValue(Str.Remove(0, 1))
                Str &= Chr(13) & Chr(10)
                StrArr.Add(Str)
                Running += 1
                Str = ""
                'New
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    If Running > 7 Then
                        Running = 0
                    End If
                    Str = Chr(2) & CStr(Running) & "O|" & i & "|"
                    If row.Item("analyzer_ref_cd").IndexOf("s") <> -1 Or row.Item("analyzer_ref_cd").IndexOf("R") <> -1 Or row.Item("analyzer_ref_cd").IndexOf("%") <> -1 Then
                        codeTest = row.Item("analyzer_ref_cd").Replace("s", "").Replace("INR", "").Replace("R", "").Replace("%", "").Trim()
                    End If
                    Str &= SpecimenId & "||"
                    Str &= "^^^" & codeTest
                    Str &= "|||||||||||^||||||||||O||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                Next

            End If
            Str = Chr(2) & CStr(Running) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to ElitePro.")
            AppSetting.WriteErrorLog(comPort, "ElitePro", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_ElitePro(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                Dim codeTest As String = ItemArr(4)
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim + codeTest
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    'dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    'resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                If ItemArr(6) = "A" Then
                    abnormalOld = ItemArr(6)
                    analyzer_ref_cdOld = ItemArr(2)
                    SpecimenIdOld = SpecimenId
                End If
                If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                    Dilution_flag = "Y"
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution3")
                    End If
                    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                    If My.Settings.TrackError Then
                        log.Info(comPort & "GetResultCodeDilution4")
                    End If
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                    abnormalOld = ""
                    analyzer_ref_cdOld = ""
                    SpecimenIdOld = ""
                End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailCobas6000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Dim S As String
        Dim No As Integer
        Dim Sample_No As String = ""
        Dim Rack_ID As String = ""
        Dim Position_No As String = ""
        Dim Rack_Type As String = ""
        Dim Container_Type As String = ""
        Dim CheckStat As String = "R"
        Dim ChecSpecimenDescriptor = "1"
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)
                S = SpecimenIdArr(7)
                No = SpecimenIdArr(5)
                SpecimenId = SpecimenIdArr(2).Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Replace("·", "").Trim
                Sample_No = SpecimenIdArr(3)
                Rack_ID = SpecimenIdArr(4)
                Position_No = SpecimenIdArr(5)
                Rack_Type = SpecimenIdArr(7)
                Container_Type = SpecimenIdArr(8)
                SpecimenIdArr(4) = SpecimenIdArr(4).Substring(0, 1)
                If SpecimenIdArr(4).IndexOf("4") <> -1 Then
                    CheckStat = "S"
                End If
                If Rack_Type = "S4" Then
                    ChecSpecimenDescriptor = "4"
                End If
                If Rack_Type = "S2" Then
                    ChecSpecimenDescriptor = "2"
                End If
            End If

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|\^&|||Host||||||TSDWN^REPLY" 'Chr(2) = STX
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'Str = "P|1|||||||M||||||40^Y"
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'If dtTestResult.Rows.Count >= 0 Then
            '    Str = "O|" & i & "|"
            '    Str &= SpecimenId & "|"
            '    Str &= Sample_No & "^"
            '    Str &= Rack_ID & "^"
            '    Str &= Position_No & "^^"
            '    Str &= Rack_Type & "^"
            '    Str &= Container_Type & "|"
            '    For Each row As DataRow In dtTestResult.Rows
            '        i += 1

            '        Str &= SepTest & row.Item("analyzer_ref_cd")
            '        If dtTestResult.Rows.Count = i Then
            '            Str &= "^"
            '        Else
            '            Str &= "^\"
            '        End If

            '    Next
            '    Str &= "|R||"
            '    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            '    Str &= "||||N||||"
            '    Str &= No
            '    Str &= "||||||||||O"
            '    Str &= Chr(13) & Chr(3)
            '    Str &= GetCheckSumValue(Str.Remove(0, 1))
            '    Str &= Chr(13) & Chr(10)
            '    StrArr.Add(Str)
            '    Running += 1
            '    Str = ""

            'End If
            'Str = "L|1|N" & Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)  'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)

            'Str = Chr(2) & "1H|\^&|||Host||||||TSDWN^REPLY" 'Chr(2) = STX (old)
            Str = Chr(2) & "1H|\^&|||Host||||||" 'Chr(2) = STXg
            Str &= Chr(13)
            'Str &= "P|1|||||||||||||"
            Str &= "O|1|"
            Str &= SpecimenId & "|"
            Str &= Sample_No & "^"
            Str &= Rack_ID & "^"
            Str &= Position_No & "^^"
            Str &= Rack_Type & "^"
            Str &= Container_Type & "|"
            If dtTestResult.Rows.Count >= 0 Then
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    If row.Item("analyzer_ref_cd") = "891" Then
                        Continue For
                    End If

                    If dtTestResult.Rows.Count = 1 And row.Item("analyzer_ref_cd") = "891" Then
                        Str &= "^^^871^\^^^881"
                    Else
                        'check 240 
                        'If i = 10 Then
                        '    Str &= SepTest & row.Item("analyzer_ref_cd")
                        '    Str &= Chr(23)
                        '    Str &= GetCheckSumValue(Str.Remove(0, 1))
                        '    Str &= Chr(13) & Chr(10)
                        '    Str &= Chr(3)
                        '    Str &= GetCheckSumValue(Str.Remove(0, 1))
                        'Else
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        'End If

                    End If

                    If dtTestResult.Rows.Count = i Then
                        Str &= "^"
                    Else
                        Str &= "^\"
                    End If
                Next
            End If
            Str &= "|" & CheckStat & "||"
            'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= "||||N||||"
            Str &= ChecSpecimenDescriptor
            Str &= "||||||||||O"
            Str &= Chr(13)
            'Str &= Chr(2) & "C|1|L|^^^^|G"
            Str &= "L|1|N"
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send To Cobas6000.")
            AppSetting.WriteErrorLog(comPort, "Cobas6000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_Cobas6000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(13))

        j = LineArr.Count

        For i = 0 To j - 1


            If LineArr(i).IndexOf("R") <> -1 And i > 2 Then

                If LineArr(i).IndexOf(Chr(23)) <> -1 Then
                    Dim check23 As Integer = InStr(1, LineArr(i), Chr(23), 1) - 1
                    LineArr(i) = LineArr(i).Substring(0, check23)
                End If

                LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(10), "").Replace("C13", "")
                Dim testLine = LineArr(i + 1)

                testLine = testLine.Replace(Chr(2), "").Replace(Chr(10), "").Replace("C13", "")
                If testLine.Length > 3 Then
                    testLine = testLine.Substring(1)
                End If
                'Dim check02Arr() As String = testLine.Split("^")
                'If check02Arr(2) = "" Then
                '    testLine = check02Arr(3)
                'End If

                Dim putLine As String = LineArr(i) + testLine
                LineArr(i) = putLine

                LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(2), "").Replace(Chr(10), "").Trim

            End If

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Or ItemArr(0).IndexOf("C") <> -1 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String

                If LineArr(i).IndexOf("|F") = -1 Then
                    LineArr(i) = LineArr(i) + LineArr(i + 1)
                End If
                SpecimenIdStr = LineArr(i).Replace("|||||||", "|").Replace("||||", "|").Replace("|||", "|").Replace("||", "|")

                    Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                    SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("·", "").Trim

                    If SpecimenId.Contains("AMM-N") Or SpecimenId.Contains("PCA1P") Or SpecimenId.Contains("PCCC1") Or SpecimenId.Contains("PCCC2") Or SpecimenId.Contains("AMM-P") Or SpecimenId.Contains("PCA1N") Then
                        Dim dt As String = Mid(SpecimenIdArr(7), 1, 14)
                        dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                        resultDate = DateTime.Parse(dt)
                    End If


                    If SpecimenId.Trim <> "" Then
                        'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                        dvTestResult = New DataView(dtTestResult)
                        dtTestResultToUpdate = dtTestResult.Clone()
                    End If

                End If

                If ItemArr(0).IndexOf("R") <> -1 Then

                If ItemArr(6) = "HH" Then
                    ItemArr(3) = "> " + ItemArr(3)
                End If
                If ItemArr(6) = "LL" Then
                    ItemArr(3) = "< " + ItemArr(3)
                End If
                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                ItemArr(2) = ItemArr(2).Substring(3)
                If ItemArr(2).IndexOf("^") <> -1 Or ItemArr(2).IndexOf("/") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("/")
                    assay_code = assay_codeArr(0).Replace("^", "").Trim
                    ItemArr(2) = assay_code
                End If


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_CaretiumXI(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim Code As String = ""
        Dim Str As String = ""
        Dim Result As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If

        Dim i, j As Integer
        Dim dvTestResult As New DataView
        Str = DeviceStr.Replace("  ", "x").Replace(" ", "x")
        LineArr = Str.Split("x")

        j = LineArr.Count

        For i = 0 To j - 1

            ItemStr = LineArr(i)

            If ItemStr = LineArr(0) Then
                Continue For
            End If
            SpecimenId = LineArr(1).Substring(12)

            If SpecimenId.Trim <> "" Then
                'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()
            End If



            If i > 2 Then
                Select Case i
                    Case 3
                        Code = "K"
                        Result = LineArr(i).Trim
                    Case 4
                        Code = "Na"
                        Result = LineArr(i).Trim
                    Case 5
                        Code = "Cl"
                        Result = LineArr(i).Trim
                    Case 6
                        Code = "Ca"
                        Result = LineArr(i).Trim
                    Case 7
                        Code = "pH"
                        Result = LineArr(i).Trim
                    Case 8
                        Code = "CO2"
                        Result = LineArr(i).Trim

                End Select


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & Code & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = Result.ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next

                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue1")
                    End If
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    If My.Settings.TrackError Then
                        log.Info(comPort & " UpdateTestResultValue2")
                    End If

                    dtTestResultToUpdate.Clear()

                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
                End If
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailAccess2(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Dim S As String
        Dim No As Integer
        Dim Sample_No As String = ""
        Dim Rack_ID As String = ""
        Dim Position_No As String = ""
        Dim Rack_Type As String = ""
        Dim Container_Type As String = ""
        Dim CheckStat As String = "R"
        Dim ChecSpecimenDescriptor = "1"
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)

                SpecimenId = SpecimenIdArr(1).Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Replace("·", "").Trim

            End If

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|\^&|||Host||||||TSDWN^REPLY" 'Chr(2) = STX
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'Str = "P|1|||||||M||||||40^Y"
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'If dtTestResult.Rows.Count >= 0 Then
            '    Str = "O|" & i & "|"
            '    Str &= SpecimenId & "|"
            '    Str &= Sample_No & "^"
            '    Str &= Rack_ID & "^"
            '    Str &= Position_No & "^^"
            '    Str &= Rack_Type & "^"
            '    Str &= Container_Type & "|"
            '    For Each row As DataRow In dtTestResult.Rows
            '        i += 1

            '        Str &= SepTest & row.Item("analyzer_ref_cd")
            '        If dtTestResult.Rows.Count = i Then
            '            Str &= "^"
            '        Else
            '            Str &= "^\"
            '        End If

            '    Next
            '    Str &= "|R||"
            '    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            '    Str &= "||||N||||"
            '    Str &= No
            '    Str &= "||||||||||O"
            '    Str &= Chr(13) & Chr(3)
            '    Str &= GetCheckSumValue(Str.Remove(0, 1))
            '    Str &= Chr(13) & Chr(10)
            '    StrArr.Add(Str)
            '    Running += 1
            '    Str = ""

            'End If
            'Str = "L|1|N" & Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)  'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)

            Str = Chr(2) & "H|\^&|||Host LIS|||||ACCESS||P|1|" 'Chr(2) = STX
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(10)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)
            Str &= "P|1|Abel, Cindy"
            Str &= Chr(13) & Chr(10)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13)
            If dtTestResult.Rows.Count >= 0 Then
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= "O|01|"
                    Str &= SpecimenId & "||"
                    Str &= SepTest
                    Str &= row.Item("analyzer_ref_cd")
                    Str &= "|R||||||A||||Serum|"
                Next
            End If
            Str &= Chr(13) & Chr(10)
            Str &= "L|1|N"
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send To Access2.")
            AppSetting.WriteErrorLog(comPort, "Access2", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_Access2(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(13))

        j = LineArr.Count

        For i = 0 To j - 1


            If LineArr(i).IndexOf("R") <> -1 And i > 3 Then

                If LineArr(i).IndexOf(Chr(23)) <> -1 Then
                    Dim check23 As Integer = InStr(1, LineArr(i), Chr(23), 1) - 1
                    LineArr(i) = LineArr(i).Substring(0, check23)
                End If

                LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(10), "").Replace("C13", "")
                Dim testLine = LineArr(i + 1)

                testLine = testLine.Replace(Chr(2), "").Replace(Chr(10), "").Replace("C13", "")
                If testLine.Length > 3 Then
                    testLine = testLine.Substring(1)
                End If
                'Dim check02Arr() As String = testLine.Split("^")
                'If check02Arr(2) = "" Then
                '    testLine = check02Arr(3)
                'End If

                Dim putLine As String = LineArr(i) + testLine
                LineArr(i) = putLine

                LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(2), "").Replace(Chr(10), "").Trim

            End If

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Or ItemArr(0).IndexOf("C") <> -1 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String

                If LineArr(i).IndexOf("|F") = -1 Then
                    LineArr(i) = LineArr(i) + LineArr(i + 1)
                End If
                SpecimenIdStr = LineArr(i).Replace("|||||||", "|").Replace("||||", "|").Replace("|||", "|").Replace("||", "|")

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("·", "").Trim

                If SpecimenId.Contains("AMM-N") Or SpecimenId.Contains("PCA1P") Or SpecimenId.Contains("PCCC1") Or SpecimenId.Contains("PCCC2") Or SpecimenId.Contains("AMM-P") Or SpecimenId.Contains("PCA1N") Then
                    Dim dt As String = Mid(SpecimenIdArr(7), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                End If


                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                If ItemArr(6) = "HH" Then
                    ItemArr(3) = "> " + ItemArr(3)
                End If
                If ItemArr(6) = "LL" Then
                    ItemArr(3) = "< " + ItemArr(3)
                End If
                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                ItemArr(2) = ItemArr(2).Substring(3)
                If ItemArr(2).IndexOf("^") <> -1 Or ItemArr(2).IndexOf("/") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(0).Replace("/", "").Trim
                    ItemArr(2) = assay_code
                End If


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailA1000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim ControlID As String = "" 'Pong'
        Dim SequenceNumbert As String = "" 'Pong'
        Dim QRD As String = "" 'Pong'
        Dim QRF As String = "" 'Pong'
        Dim STAT As String = "0"
        Dim Gender As String = "0"
        Dim Age As String = "0"
        Dim Name As String = ""
        Dim DateSystem As String = ""
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim ii As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If

            Dim SpecimenIdListArray As String() = SpecimenId.Split(Chr(2))
            SpecimenId = SpecimenIdListArray(0)
            Dim SequenceNumbertArr As String() = SpecimenIdListArray(1).Split("|")
            SequenceNumbert = SequenceNumbertArr(12)
            ControlID = SequenceNumbertArr(12)
            DateSystem = SequenceNumbertArr(6)
            QRD = SequenceNumbertArr(2)
            QRF = SequenceNumbertArr(3)

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            Dim CountItem As String = dtTestResult.Rows.Count
            Select Case ControlID
                Case "1"
                    Str = Chr(2) & "MSH|^~\&|Autobio|LIS|||" & DateSystem & "||DSR^Q01|1|P|2.3.1|" & SequenceNumbert 'Chr(2) = STX (start-of-text)
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "MSA|AA|" & ControlID
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "ERR|^^^0"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "QAK|SR|OK"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRD
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRF
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|1||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|2||" & STAT
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|3||serum"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|4||"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|5||"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|6||" & Gender
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|7||" & Age
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|8||apartment"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|9||ward"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|10||bed"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|11||"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|12||2017/2/17 11:09:05"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|13||" & Name
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|14||null"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|15||" & CountItem
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    ii = 16
                    For Each row As DataRow In dtTestResult.Rows
                        Str &= "DSP|" & ii & "||" & row.Item("analyzer_ref_cd")
                        Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Next
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                Case "9"
                    Str = Chr(2) & "MSH|^~\&|Autobio|LIS|||20170629143343||DSR^Q01|9|P|2.3.1|" & SequenceNumbert 'Chr(2) = STX (start-of-text)
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "MSA|AA|9"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "ERR|^^^0"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "QAK|SR|OK"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRD
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRF
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|1||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|2||0" & STAT
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|3||serum"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|4||BL25"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|5||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|6||167" & Gender
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|7||1" & Age
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                Case "15"
                    Str = Chr(2) & "MSH|^~\&|Autobio|LIS|||20170629143343||DSR^Q01|9|P|2.3.1|" & SequenceNumbert 'Chr(2) = STX (start-of-text)
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "MSA|AA|9"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "ERR|^^^0"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str = "QAK|SR|OK"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRD
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= QRF
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|1||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|2||14081901" & STAT
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|3||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|4||0"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|5||BL20140226"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|6||serum"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|7||2" & Age
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|8||122"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|9||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|10||100"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    Str &= "DSP|11||1"
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
            End Select


            AppSetting.WriteErrorLog(comPort, "Send", "Send To A1000.")
            AppSetting.WriteErrorLog(comPort, "A1000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr

    End Function
    Public Function ExtractResult_A1000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""


        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))

        j = LineArr.Count

        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then

                Dim SpecimenIdStr As String
                SpecimenIdStr = ItemArr(3)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("^")

                SpecimenId = SpecimenIdArr(1).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(1).Trim
                End If


                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then
                Dim resultArr() As String = ItemArr(3).Split("^")
                ItemArr(3) = resultArr(1).Trim 'Result

                Dim assyArr() As String = ItemArr(2).Split("^")
                Dim assay_code As String = ""

                assay_code = assyArr(1).Trim

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & assay_code & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailSF8050(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim SepTest As String = "^^^"
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)

                SpecimenId = SpecimenIdArr(1).Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Replace("·", "").Trim

            End If

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|\^&|||Host||||||TSDWN^REPLY" 'Chr(2) = STX
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'Str = "P|1|||||||M||||||40^Y"
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'If dtTestResult.Rows.Count >= 0 Then
            '    Str = "O|" & i & "|"
            '    Str &= SpecimenId & "|"
            '    Str &= Sample_No & "^"
            '    Str &= Rack_ID & "^"
            '    Str &= Position_No & "^^"
            '    Str &= Rack_Type & "^"
            '    Str &= Container_Type & "|"
            '    For Each row As DataRow In dtTestResult.Rows
            '        i += 1

            '        Str &= SepTest & row.Item("analyzer_ref_cd")
            '        If dtTestResult.Rows.Count = i Then
            '            Str &= "^"
            '        Else
            '            Str &= "^\"
            '        End If

            '    Next
            '    Str &= "|R||"
            '    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            '    Str &= "||||N||||"
            '    Str &= No
            '    Str &= "||||||||||O"
            '    Str &= Chr(13) & Chr(3)
            '    Str &= GetCheckSumValue(Str.Remove(0, 1))
            '    Str &= Chr(13) & Chr(10)
            '    StrArr.Add(Str)
            '    Running += 1
            '    Str = ""

            'End If
            'Str = "L|1|N" & Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)  'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)

            Str = Chr(2) & "1H|P|1|" 'Chr(2) = STX
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str &= Chr(2) & "2P|" & SpecimenId
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str &= Chr(2) & "3O|1|^" & SpecimenId & "||"
            If dtTestResult.Rows.Count >= 0 Then
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    Str &= SepTest
                    Str &= row.Item("analyzer_ref_cd")
                    If dtTestResult.Rows.Count = i Then
                        Str &= "-"
                    Else
                        Str &= "\"
                    End If
                Next
            End If
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str &= Chr(2) & "4L|1"
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send To SF8050.")
            AppSetting.WriteErrorLog(comPort, "SF8050", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_SF8050(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView


        LineArr = DeviceStr.Split(Chr(10))


        j = LineArr.Count


        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("P") <> -1 Then
                Dim SpecimenIdStr As String
                SpecimenIdStr = LineArr(i + 1)

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")

                SpecimenId = SpecimenIdArr(2).Trim
                If SpecimenId.IndexOf("^") <> -1 Then
                    Dim SpecimenIdArr2() As String = SpecimenId.Split("^")
                    SpecimenId = SpecimenIdArr2(1).Trim
                End If

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If



            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""

                ItemArr(2) = ItemArr(2).Trim



                'Try
                '    Dim dt As String = Mid(ItemArr(12), 1, 14)
                '    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                '    resultDate = DateTime.Parse(dt)
                'Catch ex As Exception
                '    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                '    log.Error(ex.Message)
                'End Try


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailDxI800(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Dim Priority As String = "R"
        Dim ActionCode As String = "A"
        Dim SpecimenType As String = "Serum"
        Dim ReportTypes As String = "F"
        Dim RackNumber As String = "1"
        Dim SamplePosition As String = "4"

        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("^")
                'SpecimenId = SpecimenIdListArray(1)

                SpecimenId = SpecimenIdArr(1).Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Replace("·", "").Trim

            End If

            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|\^&|||Host||||||TSDWN^REPLY" 'Chr(2) = STX
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'Str = "P|1|||||||M||||||40^Y"
            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            ''New
            'If dtTestResult.Rows.Count >= 0 Then
            '    Str = "O|" & i & "|"
            '    Str &= SpecimenId & "|"
            '    Str &= Sample_No & "^"
            '    Str &= Rack_ID & "^"
            '    Str &= Position_No & "^^"
            '    Str &= Rack_Type & "^"
            '    Str &= Container_Type & "|"
            '    For Each row As DataRow In dtTestResult.Rows
            '        i += 1

            '        Str &= SepTest & row.Item("analyzer_ref_cd")
            '        If dtTestResult.Rows.Count = i Then
            '            Str &= "^"
            '        Else
            '            Str &= "^\"
            '        End If

            '    Next
            '    Str &= "|R||"
            '    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            '    Str &= "||||N||||"
            '    Str &= No
            '    Str &= "||||||||||O"
            '    Str &= Chr(13) & Chr(3)
            '    Str &= GetCheckSumValue(Str.Remove(0, 1))
            '    Str &= Chr(13) & Chr(10)
            '    StrArr.Add(Str)
            '    Running += 1
            '    Str = ""

            'End If
            'Str = "L|1|N" & Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)  'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            'StrArr.Add(Str)

            Str = Chr(2) & CStr(Running) & "H|\^&|||Host LIS|||||||P|1|" 'Chr(2) = STX
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            Str &= Chr(2) & CStr(Running) & "P|1|" & dtPatient.Rows(0).Item("middle_name")
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New

            If dtTestResult.Rows.Count >= 0 Then
                For Each row As DataRow In dtTestResult.Rows
                    i += 1
                    'Str &= Chr(2) & "3O|" & i & "|" ของเก่า
                    Str &= Chr(2) & CStr(Running) & "O|" & i & "|"
                    Str &= SpecimenId & "||"
                    Str &= SepTest
                    Str &= row.Item("analyzer_ref_cd")
                    Str &= "|R||||||A||||Serum"
                    Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
                    StrArr.Add(Str)
                    Running += 1
                    If Running = 8 Then
                        Running = 0
                    End If
                    Str = ""
                Next
            End If

            Str &= Chr(2) & CStr(Running) & "L|1|F"
            Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-return) , Chr(3) = ETX (end-of-text)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send To DxI800.")
            AppSetting.WriteErrorLog(comPort, "DxI800", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_DxI800(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(13))

        j = LineArr.Count

        For i = 0 To j - 1


            'If LineArr(i).IndexOf("R") <> -1 And i > 3 Then

            '    If LineArr(i).IndexOf(Chr(23)) <> -1 Then
            '        Dim check23 As Integer = InStr(1, LineArr(i), Chr(23), 1) - 1
            '        LineArr(i) = LineArr(i).Substring(0, check23)
            '    End If

            '    LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(10), "").Replace("C13", "")

            '    Dim testLine = LineArr(i + 1)

            '    testLine = testLine.Replace(Chr(2), "").Replace(Chr(10), "").Replace("C13", "")
            '    If testLine.Length > 3 Then
            '        testLine = testLine.Substring(1)
            '    End If
            '    'Dim check02Arr() As String = testLine.Split("^")
            '    'If check02Arr(2) = "" Then
            '    '    testLine = check02Arr(3)
            '    'End If

            '    Dim putLine As String = LineArr(i) + testLine
            '    LineArr(i) = putLine

            '    LineArr(i) = LineArr(i).Replace(Chr(23), "").Replace(Chr(2), "").Replace(Chr(10), "").Trim

            'End If

            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Or ItemArr(0).IndexOf("C") <> -1 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then
                Dim SpecimenIdStr As String

                If LineArr(i).IndexOf("|F") = -1 Then
                    LineArr(i) = LineArr(i) + LineArr(i + 1)
                End If
                SpecimenIdStr = LineArr(i).Replace("|||||||", "|").Replace("||||", "|").Replace("|||", "|").Replace("||", "|")

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                SpecimenId = SpecimenIdArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("·", "").Trim

                If SpecimenId.Contains("AMM-N") Or SpecimenId.Contains("PCA1P") Or SpecimenId.Contains("PCCC1") Or SpecimenId.Contains("PCCC2") Or SpecimenId.Contains("AMM-P") Or SpecimenId.Contains("PCA1N") Then
                    Dim dt As String = Mid(SpecimenIdArr(7), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                End If


                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If ItemArr(3).IndexOf("^") <> -1 Then
                    Dim textArr() As String = ItemArr(3).Split("^")
                    ItemArr(3) = textArr(0)
                Else
                    ItemArr(3) = ItemArr(3).Trim 'Result
                End If

                Dim assay_code As String = ""
                Dim result_type As String = ""
                ItemArr(2) = ItemArr(2).Substring(3)
                If ItemArr(2).IndexOf("^") <> -1 Or ItemArr(2).IndexOf("/") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(0).Replace("/", "").Trim
                    ItemArr(2) = assay_code
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try

                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ReturnOrderDetailXL1000(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Debug.Write(SpecimenId)
        Debug.Write(Running)
        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim PacketSpecimenTypeID As String = SpecimenId
        Try

            If AppSetting.TestLisConnection = False Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: Cannot connect database !")
                Return Nothing
            End If
            If SpecimenId.IndexOf("^") <> -1 Then
                Dim SpecimenIdArr As String() = SpecimenId.Split("\")
                'SpecimenId = SpecimenIdListArray(1)
                SpecimenId = SpecimenIdArr(0).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim
            End If



            dtPatient = GetPatientInformation(SpecimenId).Copy

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1", "Y").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            'Str = Chr(2) & "1H|" & "\" & "^&||||||||||P|1|" 'Chr(2) = STX (start-of-text)
            Str = Chr(2) & "1H|\^&||||||||||P|E 1394-97|" 'Chr(2) = STX 
            Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")

            'Str &= Chr(13) & Chr(3) 'Chr(13) = CR (carriage-Return) , Chr(3) = ETX (End-Of-text)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10) 'Chr(13) = CR (carriage-return) , Chr(10) = LF (line-feed)
            Str &= Chr(13)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            Str = "P|1|||||||"
            'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("middle_name").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("firstname").ToString) & "^" & AppSetting.chkDBNull(dtPatient.Rows(0).Item("lastname").ToString)
            'Str &= "||"
            'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("birthday").ToString)
            'Str &= "|"
            'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("sex").ToString)
            'Str &= "|"
            'Str &= "||||"
            'Str &= "^Dr.1||||||||||||^^^WEST"
            'Str &= Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)
            Str &= Chr(13)

            StrArr.Add(Str)
            Running += 1
            Str = ""
            'New
            i = 0
            Str = "O|1|"
            Str &= SpecimenId & "^01||"

            For Each row As DataRow In dtTestResult.Rows
                i += 1
                If i = 1 Then
                    Str &= SepTest
                Else
                    Str &= "\" & SepTest
                End If

                Str &= row.Item("analyzer_ref_cd")

            Next

            Str &= "|R||||||A||||Serum"
            'Str &= Chr(23)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)
            'StrArr.Add(Str)
            'Running += 1
            'Str = ""
            'New

            'Str = Chr(2) & CStr(Running) & "^^PCT"
            'Str &= "||"
            'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
            'Str &= "|||||"
            'Str &= "N"
            'Str &= "||||||||||||||"
            'Str &= "Q"
            'Str &= Chr(13) & Chr(3)
            'Str &= GetCheckSumValue(Str.Remove(0, 1))
            'Str &= Chr(13) & Chr(10)
            Str &= Chr(13)
            StrArr.Add(Str)
            Running += 1
            Str = ""
            ''New


            Str = Chr(2) & "L|1|N" & Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)

            AppSetting.WriteErrorLog(comPort, "Send", "Send to XL1000.")
            AppSetting.WriteErrorLog(comPort, "XL1000", String.Join("", StrArr.ToArray()))

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
            log.Error(ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_ex: " & ex.Message)
            log.Error(ex.Message)
        End Try
        Return StrArr
    End Function
    Public Function ExtractResult_XL1000(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim itemArrOrd() As String
        Dim itemArrRes() As String
        Dim SpecimenId As String = ""
        Dim SpecimenId_Pre As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Chem1")
        End If
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(13))

        j = LineArr.Count
        Dim ItemStrOld As String = ""
        Dim abnormalOld As String = ""
        Dim analyzer_ref_cdOld As String = ""
        Dim SpecimenIdOld As String = ""
        For i = 0 To j - 1
            Dim Dilution As String = ""
            Dim Dilution_flag As String = ""
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")


            If ItemArr.Count <= 0 Then
                Continue For
            End If

            If ItemArr(0).IndexOf("O") <> -1 And SpecimenId.Length <= 1 Then
                'Dim SpecimenIdStr As String
                'SpecimenIdStr = LineArr(i)
                'SpecimenIdStr = SpecimenIdStr.Replace("||", "|")
                'Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")
                'Dim SpecimenIdArr2() As String = SpecimenIdArr(2).Split("^")
                SpecimenId = ItemArr(2).Trim.Replace(" ", "").Replace("^", "").Replace("M", "").Replace("B", "").Trim

                If SpecimenId.Trim <> "" Then
                    'dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "R", "", "N").Copy()
                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()
                End If

            End If

            If ItemArr(0).IndexOf("R") <> -1 Then

                ItemArr(3) = ItemArr(3).Trim 'Result
                Dim assay_code As String = ""
                Dim result_type As String = ""
                If ItemArr(2).IndexOf("^") <> -1 Then
                    Dim assay_codeArr() As String = ItemArr(2).Split("^")
                    assay_code = assay_codeArr(3).Trim
                    ItemArr(2) = assay_codeArr(3).Trim
                    'result_type = assay_codeArr(10).Trim
                End If

                Try
                    Dim dt As String = Mid(ItemArr(12), 1, 14)
                    dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                    resultDate = DateTime.Parse(dt)
                Catch ex As Exception
                    AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    log.Error(ex.Message)
                End Try


                'Golf 2017-08-17 start
                'If ItemArr(6) = "A" Then
                '    abnormalOld = ItemArr(6)
                '    analyzer_ref_cdOld = ItemArr(2)
                '    SpecimenIdOld = SpecimenId
                'End If
                'If ItemArr(6).Trim = "N" And abnormalOld.Trim = "A" And analyzer_ref_cdOld.Trim = ItemArr(2) And SpecimenIdOld.Trim = SpecimenId.Trim Then
                '    Dilution_flag = "Y"
                '    If My.Settings.TrackError Then
                '        log.Info(comPort & "GetResultCodeDilution3")
                '    End If
                '    dtTestResult = GetResultCodeDilution(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0", Dilution_flag, ItemArr(2), "N").Copy
                '    If My.Settings.TrackError Then
                '        log.Info(comPort & "GetResultCodeDilution4")
                '    End If
                '    dvTestResult = New DataView(dtTestResult)
                '    dtTestResultToUpdate = dtTestResult.Clone()
                '    abnormalOld = ""
                '    analyzer_ref_cdOld = ""
                '    SpecimenIdOld = ""
                'End If
                'Golf 2017-08-17 End


                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
                If dvTestResult.Count >= 1 Then
                    For ii = 0 To dvTestResult.Count - 1
                        Dim row As DataRow = dtTestResultToUpdate.NewRow
                        row("order_skey") = dvTestResult.Item(ii)("order_skey")
                        row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                        row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                        row("alias_id") = dvTestResult.Item(ii)("alias_id")
                        row("result_value") = ItemArr(3).ToString()
                        row("order_id") = dvTestResult.Item(ii)("order_id")
                        row("analyzer") = dvTestResult.Item(ii)("analyzer")
                        row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                        row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                        row("result_date") = resultDate
                        row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                        row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                        row("rerun") = dvTestResult.Item(ii)("rerun")
                        row("sample_type") = dvTestResult.Item(ii)("sample_type")
                        dtTestResultToUpdate.Rows.Add(row)
                    Next
                End If


            End If

            If ItemArr(0).IndexOf("R") <> -1 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()

                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function

    Public Function ExtractResult_Q4lyte(ByVal DeviceStr As String) As Boolean
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim Sex As String = ""
        Dim firstPat As Boolean = True
        Dim ResultStatus As Boolean = False
        Dim ref_cd As String = ""
        Dim ref_value As String = ""


        If My.Settings.TrackError Then
            log.Info(comPort & "ExtractResult_Q4lyte")
        End If
        If DeviceStr.Length <= 0 Then
            Return False
        End If
        Dim i, j As Integer
        Dim dvTestResult As New DataView
        Dim ReadStatus As Boolean = False
        DeviceStr = DeviceStr.Replace("  ", " ")
        LineArr = DeviceStr.Split(Chr(32)) '//<FS>

        SpecimenId = LineArr(1).Substring(8)

        If SpecimenId.Length > 0 Then

            If SpecimenId.Trim <> "" Then
                dtTestResult = New DataTable
                dtTestResult = GetResultCode(SpecimenId.Trim, analyzerModel, analyzerSkey, analyzerDate, "0", "N").Copy
                dvTestResult = New DataView(dtTestResult)
                dtTestResultToUpdate = dtTestResult.Clone()
            End If
        End If

        For i = 0 To LineArr.Count - 1
            Select Case i
                Case "3"
                    ref_cd = "A"
                    ref_value = LineArr(i).Trim
                Case "4"
                    ref_cd = "B"
                    ref_value = LineArr(i).Trim
                Case "5"
                    ref_cd = "C"
                    ref_value = LineArr(i).Trim
                Case "6"
                    ref_cd = "D"
                    ref_value = LineArr(i).Trim
                Case "7"
                    ref_cd = "E"
                    ref_value = LineArr(i).Trim
                Case "8"
                    ref_cd = "F"
                    ref_value = LineArr(i).Trim
            End Select


            dvTestResult.RowFilter = "analyzer_ref_cd = '" & ref_cd & "'" 'analyzer_ref_cd
            If dvTestResult.Count >= 1 Then
                For ii = 0 To dvTestResult.Count - 1
                    Dim row As DataRow = dtTestResultToUpdate.NewRow
                    row("order_skey") = dvTestResult.Item(ii)("order_skey")
                    row("result_item_skey") = dvTestResult.Item(ii)("result_item_skey")
                    row("analyzer_skey") = dvTestResult.Item(ii)("analyzer_skey")
                    row("alias_id") = dvTestResult.Item(ii)("alias_id")
                    row("result_value") = ref_value 'value
                    row("order_id") = dvTestResult.Item(ii)("order_id")
                    row("analyzer") = dvTestResult.Item(ii)("analyzer")
                    row("analyzer_ref_cd") = dvTestResult.Item(ii)("analyzer_ref_cd")
                    row("lot_skey") = dvTestResult.Item(ii)("lot_skey")
                    row("result_date") = resultDate
                    row("specimen_type_id") = dvTestResult.Item(ii)("specimen_type_id")
                    row("analyzer_cd") = dvTestResult.Item(ii)("analyzer_cd")
                    row("rerun") = dvTestResult.Item(ii)("rerun")
                    row("sample_type") = dvTestResult.Item(ii)("sample_type")
                    dtTestResultToUpdate.Rows.Add(row)
                Next
            End If

            If i > 2 Then
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue1")
                End If
                UpdateTestResultValue(dtTestResultToUpdate, False)
                If My.Settings.TrackError Then
                    log.Info(comPort & " UpdateTestResultValue2")
                End If

                dtTestResultToUpdate.Clear()
                dtTestResult.Dispose()
                dtTestResultToUpdate.Dispose()
            End If
endLine:
        Next

        'Tui add 2015-09-22  auto update formula
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula Before")
        End If

        UpdateFormula(SpecimenId)
        If My.Settings.TrackError Then
            log.Info(analyzerModel & "UpdateFormula end")
        End If
        Return True
    End Function
End Class





