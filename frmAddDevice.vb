Imports ValueSoft.DALManage
Imports System.Data.SqlClient
Imports System.Text

Public Class frmAddDevice

    Dim conString As String = MyGlobal.myConnectionString
    Public Property builder As StringBuilder
    Public Property parm As IDbDataParameter()
    Public Property SqlProvider As SqlDataProvider
    Public comPorts() As String
    Private dtAnalyzer As DataTable

    Private Sub frmAddDevice_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        For Each COMString As String In My.Computer.Ports.SerialPortNames ' Load all available COM ports.
            cbComPort.Items.Add(COMString)
        Next

        btnAdd.Enabled = False
        dtAnalyzer = GetAnalyzer()

        cbComPort.Sorted = True
        With cbAnalyzerCd
            .DataSource = dtAnalyzer
            .DisplayMember = "analyzer_cd"
            .ValueMember = "analyzer_skey"
        End With

        Try
            If comPorts.Count > 0 Then
                For Each comPort As String In comPorts
                    cbComPort.Items.Remove(comPort)
                Next
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        If cbComPort.Items.Count <= 0 Or cbAnalyzerCd.Items.Count <= 0 Then
            MsgBox("Require Comport and Analyzer !")
        Else
            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
    End Sub
    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Function GetAnalyzer() As DataTable
        SqlProvider = New SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable

        Try
            SqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_analyzer_model")

        Catch ex As ValueSoft.CoreControlsLib.CustomException
            MsgBox(ex.Message)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            SqlProvider.Dispose()
        End Try

        Return dt

    End Function

    Private Sub cbAnalyzer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbAnalyzerCd.SelectedIndexChanged
        Dim tempAnalyzerSkey As Int32
        Dim foundRows() As Data.DataRow

        Try
            tempAnalyzerSkey = cbAnalyzerCd.SelectedValue()
            foundRows = dtAnalyzer.Select("analyzer_skey = '" & tempAnalyzerSkey & "'")
            If foundRows.Count > 0 Then
                lbAnalyzerModel.Text = foundRows(0).Item("analyzer_model").ToString()
                lbSerialNo.Text = foundRows(0).Item("serial_no").ToString()
                lbAnalyzerSkey.Text = foundRows(0).Item("analyzer_skey").ToString()
                txtAnalyzerDesc.Text = foundRows(0).Item("analyzer_desc").ToString()
                cbComPort.SelectedIndex = cbComPort.FindStringExact(foundRows(0).Item("analyzer_com_port").ToString())
                'MessageBox.Show(foundRows(0).Item("analyzer_com_port").ToString())
            End If

            If cbAnalyzerCd.SelectedIndex = 0 Or cbComPort.Text.Equals(String.Empty) Then
                btnAdd.Enabled = False
            Else
                btnAdd.Enabled = True
            End If

        Catch ex As Exception

        End Try
    End Sub

End Class