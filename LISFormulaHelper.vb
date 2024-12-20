Imports System.IO
Imports System.Drawing.Drawing2D
Imports System.Drawing.Imaging
Imports System.Drawing
Imports System.ServiceModel.Channels
Imports System.Net
Imports System.Text
Imports DevExpress.Spreadsheet
Imports ValueSoft.DALManage


Public Class LISFormulaHelper

#Region "Formula"
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Shared Function GetLISParameter(ByVal SqlProvider As SqlDataProvider) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            builder = New StringBuilder
            builder.AppendLine("select * from lis_parameter")

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString()).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetLISFormulaParameter(ByVal SqlProvider As SqlDataProvider) As System.Data.DataTable
        Try
            Dim builder As StringBuilder

            builder = New StringBuilder
            builder.AppendLine("select * from lis_formula_parameter")

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString()).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetLabOrder(ByVal SqlProvider As SqlDataProvider, ByVal OrderID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder
            builder.AppendLine("select lis_lab_order.*,lis_patient.*,")
            builder.AppendLine("case when lis_lab_order.anonymous is null or len(lis_lab_order.anonymous) = 0 then patho_prefix.prefix_desc + ' ' + lis_patient.firstname + ' ' + lis_patient.lastname else lis_lab_order.anonymous end as patient_name,")
            builder.AppendLine("patho_prefix.prefix_desc,patho_sex.sex_desc,")
            builder.AppendLine("lis_visit.visit_id,customer.cust_no,customer.name as cust_name")
            builder.AppendLine("from lis_lab_order inner join lis_visit on")
            builder.AppendLine("lis_lab_order.visit_skey = lis_visit.visit_skey inner join lis_patient on")
            builder.AppendLine("lis_visit.patient_skey = lis_patient.patient_skey inner join patho_sex on")
            builder.AppendLine("lis_patient.sex = patho_sex.sex_cd inner join patho_prefix on")
            builder.AppendLine("lis_patient.prefix = patho_prefix.prefix_cd inner join customer on")
            builder.AppendLine("lis_lab_order.cust_skey = customer.cust_skey")
            builder.AppendLine("where lis_lab_order.order_id = @order_id")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("order_id", DbType.String, OrderID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetResultItem(ByVal SqlProvider As SqlDataProvider, ByVal ResultItemID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder
            builder.AppendLine("select * from lis_result_item")
            builder.AppendLine("where result_item_id = @result_item_id")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("result_item_id", DbType.String, ResultItemID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetLabResultItem(ByVal SqlProvider As SqlDataProvider, ByVal OrderID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder
            builder.AppendLine("select lis_lab_result_item.*,lis_result_item.result_item_id,lis_result_item.result_item_desc")
            builder.AppendLine("from lis_lab_order inner join lis_lab_result_item on")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_result_item.order_skey inner join lis_result_item on")
            builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey")
            builder.AppendLine("where lis_lab_order.order_id = @order_id")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("order_id", DbType.String, OrderID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetLabResultItemFormula(ByVal SqlProvider As SqlDataProvider, ByVal OrderID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            builder = New StringBuilder
            builder.AppendLine("select lis_lab_result_item.*,lis_result_item.result_item_id,lis_result_item.result_item_desc")
            builder.AppendLine("from lis_lab_order inner join lis_lab_result_item on")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_result_item.order_skey inner join lis_result_item on")
            builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey")
            builder.AppendLine("where lis_lab_order.order_id = @order_id and")
            builder.AppendLine("lis_result_item.result_type = 'F'")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("order_id", DbType.String, OrderID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetLabResultItemFormulaBySpecimenTypeID(ByVal SqlProvider As SqlDataProvider, ByVal SpecimenTypeID As String) As System.Data.DataTable
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()

            'builder = New StringBuilder
            'builder.AppendLine("select lis_lab_result_item.*,lis_result_item.result_item_id,lis_result_item.result_item_desc")
            'builder.AppendLine("from lis_lab_specimen_type inner join lis_lab_result_item on")
            'builder.AppendLine("lis_lab_specimen_type.order_skey = lis_lab_result_item.order_skey inner join lis_result_item on")
            'builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey")
            'builder.AppendLine("where lis_lab_specimen_type.specimen_type_id = @specimen_type_id and")
            'builder.AppendLine("lis_result_item.result_type = 'F'")

            builder = New StringBuilder
            builder.AppendLine("select lis_lab_result_item.order_skey,lis_lab_result_item.result_item_skey,lis_result_item.result_item_id,lis_result_item.result_item_desc")
            builder.AppendLine("from  lis_lab_order inner join lis_lab_specimen_type on")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_specimen_type.order_skey inner join lis_lab_specimen_type_test_item on")
            builder.AppendLine("lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey and")
            builder.AppendLine("lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id inner join lis_lab_test_result_item on")
            builder.AppendLine("lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey and")
            builder.AppendLine("lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey inner join lis_lab_result_item on")
            builder.AppendLine("lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey and")
            builder.AppendLine("lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey inner join lis_result_item on")
            builder.AppendLine("lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey")
            builder.AppendLine("where lis_lab_specimen_type.specimen_type_id = @specimen_type_id and")
            builder.AppendLine("lis_result_item.result_type = 'F'")
            builder.AppendLine("and lis_lab_order.status <> 'CA'") '2018-12-18
            builder.AppendLine("group by lis_lab_result_item.order_skey,lis_lab_result_item.result_item_skey,lis_result_item.result_item_id,lis_result_item.result_item_desc")

            parm = SqlProvider.GetParameterArray(1)
            parm(0) = SqlProvider.GetParameter("specimen_type_id", DbType.String, SpecimenTypeID, ParameterDirection.Input)

            Return SqlProvider.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function GetFormulaValue(Formula As String) As String
        Try
            Dim DevSpreadsheetControl As New DevExpress.XtraSpreadsheet.SpreadsheetControl

            Dim workbook As IWorkbook = DevSpreadsheetControl.Document
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            workbook.Worksheets(0).Cells("B2").Formula = String.Format("= {0}", Formula)

            Return workbook.Worksheets(0).Cells("B2").Value.ToString()
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function CalculateResultItemFormulaBySpecimenTypeID(SqlProviderTarget As SqlDataProvider, SpecimenTypeID As String, ResultItemID As String) As String
        Try
            Dim builder As StringBuilder
            Dim parm As IDbDataParameter()
            If My.Settings.TrackError Then
                log.Info(SpecimenTypeID & "Cal Formula ")
            End If 

            builder = New StringBuilder
            builder.AppendLine("select lis_lab_order.order_skey,lis_lab_order.order_id")
            builder.AppendLine("from lis_lab_order inner join lis_lab_specimen_type on")
            builder.AppendLine("lis_lab_order.order_skey = lis_lab_specimen_type.order_skey")
            builder.AppendLine("where lis_lab_specimen_type.specimen_type_id = @specimen_type_id")
            builder.AppendLine("and lis_lab_order.status <> 'CA'")
            builder.AppendLine("group by lis_lab_order.order_skey,lis_lab_order.order_id")

            parm = SqlProviderTarget.GetParameterArray(1)
            parm(0) = SqlProviderTarget.GetParameter("specimen_type_id", DbType.String, SpecimenTypeID, ParameterDirection.Input)

            Dim dtLabOrder As System.Data.DataTable = SqlProviderTarget.ExecuteDataSet(CommandType.Text, builder.ToString(), parm).Tables(0)

            If dtLabOrder.Rows.Count > 0 Then
                Return CalculateResultItemFormula(SqlProviderTarget, dtLabOrder.Rows(0)("order_id"), ResultItemID)
            End If

            Throw New Exception(String.Format("Not Found Specimen Type ID {0}", SpecimenTypeID))
        Catch ex As Exception
            Throw
        End Try
    End Function

    Public Shared Function CalculateResultItemFormula(SqlProviderTarget As SqlDataProvider, OrderID As String, ResultItemID As String) As String
        Try
            If My.Settings.TrackError Then
                log.Info(OrderID & "CalculateResultItemFormula start ")
            End If
            Dim dtLabOrder As System.Data.DataTable = GetLabOrder(SqlProviderTarget, OrderID)
            Dim dtLabResultItem As System.Data.DataTable = GetLabResultItem(SqlProviderTarget, OrderID)
            Dim dtResultItem As System.Data.DataTable = GetResultItem(SqlProviderTarget, ResultItemID)
            Dim dtFormulaParameter As System.Data.DataTable = GetLISFormulaParameter(SqlProviderTarget)
            Dim dtParameter As System.Data.DataTable = GetLISParameter(SqlProviderTarget)

            If dtLabOrder.Rows.Count = 0 Then
                Throw New Exception(String.Format("No Found Lab Order {0}", OrderID))
            End If

            If dtResultItem.Rows.Count = 0 Then
                Throw New Exception(String.Format("No Found Result Item {0}", ResultItemID))
            End If

            If dtResultItem.Rows(0)("result_type") <> "F" Then
                Throw New Exception(String.Format("Result Item {0} isnot Result formula", ResultItemID))
            End If

            If IsDBNull(dtResultItem.Rows(0)("formula")) OrElse String.IsNullOrEmpty(dtResultItem.Rows(0)("formula")) Then
                Throw New Exception(String.Format("Result Item {0} not found formula", ResultItemID))
            End If

            Dim Formula As String = dtResultItem.Rows(0)("formula")
            Dim FormulaArr() As String = Formula.Split(" ")
            Dim DictResultItem As New SortedDictionary(Of String, String)
            Dim DictAttibuteParameter As New SortedDictionary(Of String, String)
            Dim DictParameter As New SortedDictionary(Of String, String)
            Dim StartPointResultItem As Integer = -1
            Dim StopPointResultItem As Integer = -1
            Dim StartPointAttibuteParameter As Integer = -1
            Dim StopPointAttibuteParameter As Integer = -1
            Dim StartPointParameter As Integer = -1
            Dim StopPointParameter As Integer = -1
            Dim DecimalDigit As Integer

            If IsDBNull(dtResultItem.Rows(0)("decimal_digit")) Then
                DecimalDigit = 0
            Else
                DecimalDigit = dtResultItem.Rows(0)("decimal_digit")
            End If



            For i As Integer = 0 To Formula.Length - 1
                Application.DoEvents() 'Golf  Application.DoEvents()
                Select Case Formula(i)
                    Case "$"
                        If StartPointResultItem = -1 Then
                            StartPointResultItem = i
                        Else
                            StopPointResultItem = i

                            Dim ComponentResultItemID As String = Formula.Substring(StartPointResultItem + 1, StopPointResultItem - StartPointResultItem - 1)
                            Dim ResultValue As String = Nothing

                            Dim foundrow() As System.Data.DataRow = dtLabResultItem.Select(String.Format("result_item_id = '{0}'", ComponentResultItemID))
                            If foundrow.Length > 0 Then
                                If Not IsDBNull(foundrow(0)("result_value")) AndAlso Not String.IsNullOrEmpty(foundrow(0)("result_value")) Then
                                    ResultValue = foundrow(0)("result_value")
                                End If
                            End If

                            If Not DictResultItem.ContainsKey(ComponentResultItemID) Then
                                DictResultItem.Add(ComponentResultItemID, ResultValue)
                            End If

                            StartPointResultItem = -1
                            StopPointResultItem = -1
                        End If
                    Case "@"
                        If StartPointAttibuteParameter = -1 Then
                            StartPointAttibuteParameter = i
                        Else
                            StopPointAttibuteParameter = i

                            Dim ParameterName As String = Formula.Substring(StartPointAttibuteParameter + 1, StopPointAttibuteParameter - StartPointAttibuteParameter - 1)
                            Dim ParameterValue As String = Nothing
                            Dim foundrow() As System.Data.DataRow = dtFormulaParameter.Select(String.Format("parameter_name = '{0}'", ParameterName))

                            If foundrow.Length = 0 Then
                                Throw New Exception(String.Format("{0} isn't attibute parameter For Result Item {1}", ParameterName, dtResultItem.Rows(0)("result_item_desc")))
                            Else
                                Dim FormulaParameter As String = foundrow(0)("parameter_formula")
                                Dim FieldValue As String = String.Empty

                                If dtLabOrder.Columns(foundrow(0)("parameter_field")) IsNot Nothing Then
                                    If CType(dtLabOrder.Columns(foundrow(0)("parameter_field")), System.Data.DataColumn).DataType Is Type.GetType("System.DateTime") Then
                                        FieldValue = String.Format("{0:yyyy-MM-dd}", dtLabOrder.Rows(0)(foundrow(0)("parameter_field")))
                                    Else
                                        FieldValue = dtLabOrder.Rows(0)(foundrow(0)("parameter_field"))
                                    End If
                                End If

                                FormulaParameter = FormulaParameter.Replace(String.Format("@{0}@", ParameterName), FieldValue)

                                Try
                                    ParameterValue = LISFormulaHelper.GetFormulaValue(FormulaParameter)
                                Catch ex As Exception
                                    Throw New Exception(ex.Message)
                                End Try
                            End If

                            If Not DictAttibuteParameter.ContainsKey(ParameterName) Then
                                DictAttibuteParameter.Add(ParameterName, ParameterValue)
                            End If

                            StartPointAttibuteParameter = -1
                            StopPointAttibuteParameter = -1
                        End If
                    Case "#"
                        If StartPointParameter = -1 Then
                            StartPointParameter = i
                        Else
                            StopPointParameter = i

                            Dim ParameterName As String = Formula.Substring(StartPointParameter + 1, StopPointParameter - StartPointParameter - 1)
                            Dim ParameterValue As String = Nothing
                            Dim foundrow() As System.Data.DataRow = dtParameter.Select(String.Format("parameter_cd = '{0}'", ParameterName))

                            If foundrow.Length = 0 Then
                                Throw New Exception(String.Format("{0} isn't parameter For Result Item {1}", ParameterName, dtResultItem.Rows(0)("result_item_desc")))
                            Else
                                ParameterValue = foundrow(0)("parameter_value")
                            End If

                            If Not DictParameter.ContainsKey(ParameterName) Then
                                DictParameter.Add(ParameterName, ParameterValue)
                            End If

                            StartPointParameter = -1
                            StartPointParameter = -1
                        End If
                End Select
            Next

            For Each entry As KeyValuePair(Of String, String) In DictResultItem
                Application.DoEvents() 'Golf  Application.DoEvents()
                If entry.Value Is Nothing Then
                    Dim ComponentResultItemID As String = entry.Key
                    Dim ResultItemDesc As String = Nothing
                    Dim foundrow() As System.Data.DataRow = dtLabResultItem.Select(String.Format("result_item_id = '{0}'", ComponentResultItemID))
                    If foundrow.Length > 0 Then
                        ResultItemDesc = foundrow(0)("result_item_desc")
                    End If

                    Throw New Exception(String.Format("Not Found Value Of Result Item : {0}", ResultItemDesc))
                End If

                Formula = Formula.Replace(String.Format("${0}$", entry.Key), entry.Value)
            Next

            For Each entry As KeyValuePair(Of String, String) In DictAttibuteParameter
                Application.DoEvents() 'Golf  Application.DoEvents()
                If entry.Value Is Nothing Then
                    Dim ParameterName As String = entry.Key

                    Throw New Exception(String.Format("Not Found Value Of Attibute Parameter {0}", ParameterName))
                End If

                Formula = Formula.Replace(String.Format("@{0}@", entry.Key), entry.Value)
            Next

            For Each entry As KeyValuePair(Of String, String) In DictParameter
                Application.DoEvents() 'Golf  Application.DoEvents()
                If entry.Value Is Nothing Then
                    Dim ParameterName As String = entry.Key

                    Throw New Exception(String.Format("Not Found Value Of Parameter {0}", ParameterName))
                End If

                Formula = Formula.Replace(String.Format("#{0}#", entry.Key), entry.Value)
            Next

            Dim FormulaValue As String
            Try
                FormulaValue = GetFormulaValue(Formula)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

            If IsNumeric(FormulaValue) AndAlso FormulaValue.IndexOf("+") = -1 Then
                FormulaValue = System.Math.Round(CDec(FormulaValue), DecimalDigit)
            End If

            Return FormulaValue
        Catch ex As Exception
            Throw
        End Try
        If My.Settings.TrackError Then
            log.Info(OrderID & "CalculateResultItemFormula end ")
        End If
    End Function

#End Region

End Class
