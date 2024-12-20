Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports ValueSoft.DALManage
Imports DeviceControlApp.ComPortMonitor
 

Public Class AnalyzerInterface
    'Dim regis As Core.LIS.Registration

    'Tui add 2013-09-18
    Public Property builder As StringBuilder
    Public Property parm As IDbDataParameter()
    Public Property SqlProvider As SqlDataProvider
    Dim dtTestResult As New DataTable
    Dim dtTestResultToUpdate As New DataTable
    Dim conString As String = MyGlobal.myConnectionString
    Dim resultDate As DateTime = DateTime.Now()
    Dim analyzerDate As DateTime = DateTime.Now()
    Dim comPort As String
    Dim analyzerSkey As Int32
    Dim analyzerModel As String

    Sub New(ByVal argComport As String, ByVal argAnalyzerSkey As Int32, ByVal argAnalyzerModel As String)
        'regis = New Core.LIS.Registration(_lis)
        comPort = argComport
        analyzerSkey = argAnalyzerSkey
        analyzerModel = argAnalyzerModel
    End Sub

    Public Function ExtractResult(ByVal DeviceStr As String) As Boolean

        'Debug.WriteLine("ExtractResult")
        'Debug.WriteLine(DeviceStr)
        'Debug.WriteLine(analyzerModel)
        'Debug.WriteLine(analyzerSkey)

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim ResultStatus As Boolean = False

        Try

            If AppSetting.TestLisConnection = False Then
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
            ElseIf analyzerModel = "Daytona" Or analyzerModel = "Imola" Or analyzerModel = "Suzuka" Or analyzerModel = "DX300" Or analyzerModel = "XS-Serie" Or analyzerModel = "Liaison" Or analyzerModel = "COBAS_e411" Or analyzerModel = "SUZUKA" Or analyzerModel = "RXMODENA" Or analyzerModel = "Modena" Then
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
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
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
        End Try

        Return True
    End Function
    Public Function ExtractResult_Micro(ByVal DeviceStr As String) As Boolean

        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim SpecimenId As String = ""
        Dim Sex As String = ""
        Dim ResultStatus As Boolean = False
        Dim resultCode As String = ""
        Dim resultValue As String = ""
        Dim isFoundHbA1c As Boolean = False

        If analyzerModel = "AS300" Or analyzerModel = "AS720" Then
            If DeviceStr.IndexOf("LOT") = -1 Then
                Return False
            End If
        ElseIf analyzerModel = "H-500" Then
            'If DeviceStr.IndexOf("ALB") = -1 Then
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
                tempQcDate = LineArr(i)
                tempQcDate = tempQcDate.Replace("~", String.Empty)
                tempQcDate = tempQcDate.Trim()
                If tempQcDate.Length = 20 Then
                    DateTime.TryParse(tempQcDate, analyzerDate)
                    Continue For
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
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
        Dim dvTestResult As New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

        '##2## Get test result by test code
        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemStr = LineArr(i).TrimStart
            ItemStr = ItemStr.TrimEnd
            'Golf 2016-04-04 H-500 Remove value+
            If analyzerModel = "H-500" Then
                If ItemStr.IndexOf("+") > -1 Then 'Golf mark
                    ItemStr = ItemStr.Replace(ItemStr.Substring(ItemStr.IndexOf(" "), ItemStr.IndexOf("+")).Trim, " ")
                End If
                If ItemStr.IndexOf("Normal") > -1 Then 'Golf mark 2016-11-24
                    ItemStr = ItemStr.Replace("Normal", " ")
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

            If analyzerModel = "H-500" Then 
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
            If analyzerModel = "H-500" Then

                'Golfmark

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


                    Case Else
                        If ItemArr(k) = "Pos" Or ItemArr(k) = "Neg" Or ItemArr(k) = "Trace" Or ItemArr(k) = "Few" Then
                            resultValue = """" + ItemArr(k) + """"
                        Else
                            resultValue = GetResultMicro(ItemArr(k))
                        End If
                End Select


            Else
                resultValue = GetResultMicro(ItemArr(k))
            End If

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
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
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

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
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
        
        If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
            Return False
        End If

        'If analyzerModel = "Imola" Then

        '    If DeviceStr.IndexOf(Chr(4)) = -1 Then  '****** ENQ *******
        '        Return False
        '    End If
        '    'Golf 2016-04-05
        'ElseIf analyzerModel = "DX300" Then
        '    Dim ttttt As Integer = DeviceStr.IndexOf(Chr(3) + "06")
        '    Dim aaaaa As Integer = DeviceStr.Length
        '    Dim sssss As String = Asc(DeviceStr.IndexOf(Chr(4)) - 1)
        '    If DeviceStr.IndexOf(Chr(3) + "06") = -1 Then
        '        Return False
        '    End If
        'End If
        'Golf 2016-04-05




        Dim i, j As Integer
        'Tui Declare
        Dim dvTestResult As New DataView

        LineArr = DeviceStr.Split(Chr(10))
        j = LineArr.Count

        For i = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")

            If ItemArr.Count <= 0 Then
                Continue For
            End If
            ' Dim strTest As String = ItemArr(0).Substring(2, 1)
            If ItemArr(0).IndexOf("P") <> -1 Then

                Dim SpecimenIdStr As String
                'Get Lab No
                'Tui Add Imola
                If analyzerModel.Equals("Imola") Or analyzerModel.Equals("RXMODENA") Or analyzerModel.Equals("Modena") Then 'Golf add DX300, RXMODENA
                    SpecimenIdStr = LineArr(i + 2)
                    'End Tui Add Emola
                ElseIf analyzerModel.Equals("DX300") Then

                    SpecimenIdStr = LineArr(i + 1)
                Else

                    SpecimenIdStr = LineArr(i + 1)
                End If

                Dim SpecimenIdArr() As String = SpecimenIdStr.Split("|")


                'SpecimenId = SpecimenIdArr(2).Trim Golf comment

                'Golf create c ase1
                Select Case analyzerModel
                    Case "Imola", "RXMODENA", "Modena"
                        SpecimenId = SpecimenIdArr(2).Trim
                    Case "DX300", "SUZUKA"
                        'Debug.WriteLine("Between 6 and 8, inclusive")
                        SpecimenId = SpecimenIdArr(2).Substring(0, SpecimenIdArr(2).IndexOf(Chr(94))).Trim
                    Case Else
                        SpecimenId = SpecimenIdArr(2).Trim
                End Select
                'Golf create c ase
                ' AppSetting.WriteErrorLog(comPort, "Error", "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa" & SpecimenId)

                If SpecimenId.Trim <> "" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                    dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
                    dvTestResult = New DataView(dtTestResult)
                    dtTestResultToUpdate = dtTestResult.Clone()

                End If

                'Continue For

                '<<<< ******* Order Line ******** >>>>
                ' Order
            End If

            If ItemArr(0).IndexOf("O") <> -1 Then

                If analyzerModel = "XS-Serie" Then
                    'Example 4O|1||^^    11070200022^M|.......
                    itemArrOrd = ItemArr(3).Split("^") 'Split ที่ ^ และใช้ตำแหน่งที่ 2
                    itemArrOrd(2) = itemArrOrd(2).Trim
                    SpecimenId = itemArrOrd(2).Trim
                Else
                    'DX300,Imola,Daytona,Pentra60,Liason... ใช้ 10 หลักแรก
                    If analyzerModel.Equals("Daytona") Then
                        ItemArr(2) = ItemArr(2).Trim
                        SpecimenId = ItemArr(2).Trim
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

                            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
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

                            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
                            
                            dvTestResult = New DataView(dtTestResult)
                            dtTestResultToUpdate = dtTestResult.Clone()

                        End If
                    End If
                End If

                'Golf 2016-04-18

                'If SpecimenId_Pre = "" Then 'SpecimenId อันแรกของ Batch
                '    SpecimenId_Pre = SpecimenId
                '    Continue For
                'ElseIf SpecimenId = SpecimenId_Pre Then 'ยังอยู่ใน loop ของ SpecimenId เดิม
                '    SpecimenId_Pre = SpecimenId
                '    Continue For
                'ElseIf SpecimenId <> SpecimenId_Pre Then 'ขึ้น SpecimenId ใหม่   >>  บันทึก SpecimenId เก่า
                '    If SpecimenId <> "" Then

                '        'Tui comments 2013-09-24
                '        'AddResult(SpecimenId_Pre, analyzerModel, Now, dt_result) 'บันทึกข้อมูล SpecimenId ก่อนหน้า  
                '        'dt_result.Rows.Clear() 'Clear ข้อมูล Data table เพื่อรอรับ Result ชุดใหม่
                '    End If
                '    SpecimenId_Pre = SpecimenId
                '    Continue For
                'End If
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
                If analyzerModel = "COBAS_e411" OrElse analyzerModel = "Imola" OrElse analyzerModel = "Daytona" OrElse analyzerModel = "RXMODENA" OrElse analyzerModel = "Modena" Then

                    Try
                        If ItemArr.Count = 13 Then
                            Dim dt As String = Mid(ItemArr(12), 1, 14)
                            dt = Mid(dt, 1, 4) & "-" & Mid(dt, 5, 2) & "-" & Mid(dt, 7, 2) & " " & Mid(dt, 9, 2) & ":" & Mid(dt, 11, 2) & ":" & Mid(dt, 13, 2)
                            resultDate = DateTime.Parse(dt)
                        End If
                    Catch ex As Exception
                        AppSetting.WriteErrorLog(comPort, "Error", "Get IQC Date" & ex.Message)
                    End Try

                End If
                'End Tui Add 2015-09-15 Get IQC Date
                Debug.WriteLine(SpecimenId)
                dvTestResult.RowFilter = "analyzer_ref_cd = '" & ItemArr(2) & "'"
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
                    dtTestResultToUpdate.Rows.Add(row)
                End If

            End If

            'Imola
            If analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Then
                GoTo endLine
            End If
            If analyzerModel = "Imola" Or analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Or analyzerModel = "Liaison" Or analyzerModel = "RXMODENA" Or analyzerModel = "Modena" Then  'Golf 2016-04-18 add dx300 'Golf 2016-10-19 add Liaison
                'ถ้า ผลของ DX300 และ SUZUKA เบิ้ล ต้องย้ายลงมาไว้หลัง Loop และเปิด if ด้านบนด้วย
                If ItemArr(0).IndexOf("R") <> -1 Then
                    UpdateTestResultValue(dtTestResultToUpdate, False)
                    dtTestResult.Dispose()
                    dtTestResultToUpdate.Dispose()
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
        If analyzerModel = "DX300" Or analyzerModel = "SUZUKA" Then
            UpdateTestResultValue(dtTestResultToUpdate, False)
            dtTestResult.Dispose()
            dtTestResultToUpdate.Dispose()
        End If

        'Tui add 2015-09-22  auto update formula
        UpdateFormula(SpecimenId)

        Return True

    End Function 
    Public Function ExtractResult_HbA1c(ByVal DeviceStr As String) As Boolean
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
            'resultValue = ItemArrR(2).ToString.Replace("HbA0", "") 'Golf Comments
            ' resultCode = ItemArrR(2).ToString.Substring(1, ItemArrR(2).ToString.IndexOf("=") - 1)
            resultValue = ItemArrR(1).ToString.Substring(ItemArrR(2).ToString.IndexOf("=") + 2)
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
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
            ElseIf analyzerModel = "Liaison" Then
                Return ReturnOrderDetail_LIASON(SpecimenId, Running) 'Golf 2016-10-06

                ''Tui Add COBAS_e411, Feb 24, 2014
            ElseIf analyzerModel = "COBAS_e411" Then
                Return ReturnOrderDetail_COBAS_e411(SpecimenId)
                ''End Tui Add COBAS_e411, Feb 24, 2014
            ElseIf analyzerModel = "SUZUKA" Then
                SpecimenId = SpecimenId.Substring(0, SpecimenId.IndexOf(Chr(94))).Trim 'Golf 
            End If
            dtPatient = GetPatientInformation(SpecimenId).Copy
            'GolfMark 2016-05-17
            'AppSetting.WriteErrorLog(comPort, "Error", "recheck7 ")
            ' AppSetting.WriteErrorLog(comPort, "Error", "recheck8 " & SpecimenId)
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
            '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10 ")
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()

            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck10.1 " & analyzerModel)

            ''Order
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
                        ' StrArr.Add(Chr(5))
                        'Header
                        'Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
                        'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                        'Str &= Chr(13) & Chr(3)

                        'Str &= GetCheckSumValue(Str.Remove(0, 1))

                        'Str &= Chr(13) & Chr(10)
                        'StrArr.Add(Str)
                        'Running += 1

                        ''Patient
                        'Str = ""
                        'Str = Chr(2) & CStr(Running) & "P|1|"
                        'Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                        ''Str &= "Test"
                        'Str &= "||||||||||||"
                        'Str &= Chr(13) & Chr(3)
                        'Str &= GetCheckSumValue(Str.Remove(0, 1))

                        'Str &= Chr(13) & Chr(10)
                        'StrArr.Add(Str)
                        Running += 1


                        i += 1
                        Str &= SpecimenIdReturn & "||"

                        Str = ""

                        Str = Chr(2) & CStr(Running) & "O|" & i & "|"

                        Str &= SpecimenIdReturn & "||"
                        'Str &= "160037143^1^1^3^N" & "||"


                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "|"

                        ''GolfTest
                        'Str &= "R|"
                        'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                        'Str &= "||||||||" & "|1||||||||||"
                        'End If 

                        Str &= Chr(13) & Chr(3)
                        Str &= GetCheckSumValue(Str.Remove(0, 1))

                        Str &= Chr(13) & Chr(10)
                        StrArr.Add(Str)



                        ''EndLine
                        'Str = ""
                        'Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                        'Str &= GetCheckSumValue(Str.Remove(0, 1))

                        'Str &= Chr(13) & Chr(10) 
                        'StrArr.Add(Str)

                        'StrArr.Add(Chr(4))

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

                    'Golf 2016-04-11
                Case "Modena", "RXMODENA"
                    'Header
                    StrArr.Add(Chr(5))
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    '   AppSetting.WriteErrorLog(comPort, "Error", "recheck10.3 ")
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.4 " & Str)
                    'Patient
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    'Str = CStr(Running).PadLeft(2, " ") & "P|1|"
                    'Str &= dt_patient.Item(0).HN
                    'Str &= GetPatientInformation(SpecimenId).Rows(0).Item(0).ToString
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "Test"
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    'Old
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
                    'Old

                    'New
                    For Each row As DataRow In dtTestResult.Rows
                        Running += 1
                        i += 1
                        Str = ""
                        Str = Chr(2) & CStr(Running) & "O|" & i & "|"

                        Str &= SpecimenId & "||"

                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        'Str &= "|" 
                        ' ''GolfTest
                        'Str &= "R|"
                        'Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                        'Str &= "||||||||" & "|1||||||||||"
                        ''End If 

                        Str &= Chr(13) & Chr(3)
                        Str &= GetCheckSumValue(Str.Remove(0, 1)) 
                        Str &= Chr(13) & Chr(10)
                        StrArr.Add(Str) 
                    Next
                    'New
                    
                    Running += 1 
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10) 
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))
                Case Else

                    ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.2 ")
                    'Header
                    StrArr.Add(Chr(5))
                    Str = Chr(2) & "1H|" & "\" & "^&|||Host|||||||||"
                    Str &= Now.Year & CStr(Now.Month).PadLeft(2, "0") & CStr(Now.Day).PadLeft(2, "0") & CStr(Now.Hour).PadLeft(2, "0") & CStr(Now.Minute).PadLeft(2, "0") & CStr(Now.Second).PadLeft(2, "0")
                    Str &= Chr(13) & Chr(3)

                    Str &= GetCheckSumValue(Str.Remove(0, 1))
                    '   AppSetting.WriteErrorLog(comPort, "Error", "recheck10.3 ")
                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.4 " & Str)
                    'Patient
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "P|1|"
                    'Str = CStr(Running).PadLeft(2, " ") & "P|1|"
                    'Str &= dt_patient.Item(0).HN
                    'Str &= GetPatientInformation(SpecimenId).Rows(0).Item(0).ToString
                    Str &= AppSetting.chkDBNull(dtPatient.Rows(0).Item("hn").ToString)
                    'Str &= "Test"
                    Str &= "||||||||||||"
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    Str = ""
                    i = 0

                    Str = Chr(2) & CStr(Running) & "O|1|"
                    Str &= SpecimenId & "||"    'Tui Comment 2013-11-05
                    '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.7 ")
                    For Each row As DataRow In dtTestResult.Rows
                        i += 1
                        Str &= SepTest & row.Item("analyzer_ref_cd")
                        Str &= "\"
                    Next

                    


                    '     AppSetting.WriteErrorLog(comPort, "Error", "recheck10.10 " & Str)
                    If Str.Substring(Str.Length - 1, 1) = "\" And analyzerModel <> "DX300" And analyzerModel <> "SUZUKA" Then

                        Str = Str.Substring(0, Str.Length - 1)
                    End If
                    Str &= Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    StrArr.Add(Str)
                    Running += 1
                    '    AppSetting.WriteErrorLog(comPort, "Error", "recheck10.11 " & Str)
                    'EndLine
                    Str = ""
                    Str = Chr(2) & CStr(Running) & "L|1" & Chr(13) & Chr(3)
                    Str &= GetCheckSumValue(Str.Remove(0, 1))

                    Str &= Chr(13) & Chr(10)
                    'Str = CStr(Running).PadLeft(2, " ") & "L|1" & Chr(13)
                    StrArr.Add(Str)

                    StrArr.Add(Chr(4))
                    '    AppSetting.WriteErrorLog(comPort, "Error", "recheck10.12 " & Str)
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
            End Select

            ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.15 ")

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail: " & ex.Message)
        End Try
        ' AppSetting.WriteErrorLog(comPort, "Error", "recheck10.16 ")
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck10.17 " & (String.Join("", StrArr.ToArray())))
        Return StrArr

    End Function
    Public Function ReturnOrderDetail_COBAS_e411(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList

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

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
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
            'Str = CStr(Running).PadLeft(2, " ") & "P|1|"
            'Str &= dt_patient.Item(0).HN
            'Str &= dtPatient.Rows(0).Item(1).ToString
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
            'Str = CStr(Running).PadLeft(2, " ") & "O|"
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
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.ReturnOrderDetail_COBAS_e411: " & ex.Message)
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
         
        dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
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

            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
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

    Public Function GetCheckSumValue(ByVal frame As String) As String
        Dim checksum As String = "00"

        Dim byteVal As Integer = 0
        Dim sumOfChars As Integer = 0
        Dim complete As Boolean = False

        Try

            'take each byte in the string and add the values
            For idx As Integer = 0 To frame.Length - 1
                byteVal = Convert.ToInt32(frame(idx))

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
                checksum = Convert.ToString(sumOfChars Mod 256, 16).ToUpper()
            End If

        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetCheckSumValue: " & ex.Message)
        End Try

        'if checksum is only 1 char then prepend a 0
        Return DirectCast(If(checksum.Length = 1, "0" & checksum, checksum), String)
    End Function

    Function GetResultCode(ByVal SpecimenId As String, ByVal analyzer As String, ByVal analyzerSkey As Int32, ByVal analyzerDate As DateTime, ByVal flagExist As String) As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable
        '  AppSetting.WriteErrorLog(comPort, "Error", "recheck9 ")
        Try
            parm = SqlProvider.GetParameterArray(5)
            parm(0) = SqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            parm(1) = SqlProvider.GetParameter("analyzer_model", DbType.String, analyzer, ParameterDirection.Input)
            parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, analyzerSkey, ParameterDirection.Input)
            parm(3) = SqlProvider.GetParameter("analyzer_date", DbType.DateTime, analyzerDate, ParameterDirection.Input)
            parm(4) = SqlProvider.GetParameter("flag_exist", DbType.String, flagExist, ParameterDirection.Input)
            SqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_test_result_model", parm)
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.1 ")
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.2 " & SpecimenId)
        Catch ex As ValueSoft.CoreControlsLib.CustomException
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.21 ")
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
        Catch ex As Exception
            '   AppSetting.WriteErrorLog(comPort, "Error", "recheck9.22 ")
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetResultCode: " & ex.Message)
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

            SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
            SqlProvider.BeginTransaction()

            For Each row As DataRow In dtResult.Rows
                parm = SqlProvider.GetParameterArray(6)
                parm(0) = SqlProvider.GetParameter("order_skey", DbType.Int32, row("order_skey"), ParameterDirection.Input)
                parm(1) = SqlProvider.GetParameter("result_item_skey", DbType.Int32, row("result_item_skey"), ParameterDirection.Input)
                parm(2) = SqlProvider.GetParameter("analyzer_skey", DbType.Int32, row("analyzer_skey"), ParameterDirection.Input)
                parm(3) = SqlProvider.GetParameter("result_value", DbType.String, row("result_value"), ParameterDirection.Input)
                parm(4) = SqlProvider.GetParameter("lot_skey", DbType.Int32, row("lot_skey"), ParameterDirection.Input)
                parm(5) = SqlProvider.GetParameter("result_date", DbType.DateTime, row("result_date"), ParameterDirection.Input)
                SqlProvider.ExecuteNonQuerySP("sp_lis_interface_lis_lab_result_value_model_ins", parm)

                row("sending_status") = "Sent"
                dtTemp.Rows.Add(row("order_skey").ToString, row("result_item_skey").ToString, row("analyzer_skey").ToString, row("analyzer_ref_cd"), row("alias_id"), row("result_value"), row("specimen_type_id").ToString, row("analyzer").ToString, row("sending_status"), 0, row("result_date"), row("order_id"), row("analyzer_cd").ToString)

                If dtTemp.Rows.Count > 0 And fm IsNot Nothing Then
                    fm.AppendResult(dtTemp)
                End If

                dtTemp.Clear()

            Next

            SqlProvider.CommitTransaction()

        Catch ex As SqlException
            If ex.Number = 1205 Then

                SqlProvider.RollbackTransaction()
                AppSetting.WriteErrorLog(comPort, "Error", "Deadlock: " & ex.Message)
                retryCounter += 1

                If retryCounter >= 3 Then
                    SqlProvider.Dispose()
                    Exit Function
                End If

                System.Threading.Thread.Sleep(100)
                GoTo RETRY

            Else

                SqlProvider.RollbackTransaction()
                AppSetting.WriteErrorLog(comPort, "Error", "UpdateTestResultValue: " & ex.Message)

            End If
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "UpdateTestResultValue: " & ex.Message)
            SqlProvider.RollbackTransaction()
        Finally
            SqlProvider.Dispose()
        End Try

    End Function

    Function UpdateFormula(ByVal specimenTypeId As String) As Int16

        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)

        Try
            SqlProvider.BeginTransaction()
            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("specimen_type_id", DbType.String, specimenTypeId, ParameterDirection.Input)
            SqlProvider.ExecuteNonQuerySP("sp_lis_interface_update_formula", parm)
            SqlProvider.CommitTransaction()
        Catch ex As Exception
            AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.UpdateFormula: " & ex.Message)
            SqlProvider.RollbackTransaction()
        Finally
            SqlProvider.Dispose()
        End Try

    End Function

    Public Function GetResultMicro(value As String) As String
        Dim output As StringBuilder = New StringBuilder

        Try

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
        Catch ex As Exception
            If patientSqlProvider.Transaction IsNot Nothing AndAlso patientSqlProvider.GetTransactionCount > 0 Then
                AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetPatientInformation: " & ex.Message)
            End If
        Finally
            patientSqlProvider.Dispose()
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
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "0").Copy
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

End Class





