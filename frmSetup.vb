Imports DeviceControlApp.AppSetting
Imports System.Data.SqlClient
Imports System.Text
Imports ValueSoft.DALManage
Imports System.IO

Imports ValueSoft.Common
Imports ValueSoft.CoreControlsLib 

Public Class frmSetup
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public Property builder As StringBuilder
    Public Property parm As IDbDataParameter()

    Dim dsSetup As New DataSet
    Dim dsTemp As New DataSet

    Dim blHisConnected As Boolean
    Dim blLisConnected As Boolean

    Dim strSqlCon As String
    Dim objSqlCon As System.Data.SqlClient.SqlConnection
    Dim ojbSqlCmd As System.Data.SqlClient.SqlCommand
    Dim setting As New AppSetting

    Sub InitDataset()

        dsSetup.Tables.Add("LisConnection")
        dsSetup.Tables("LisConnection").Columns.Add("Server")
        dsSetup.Tables("LisConnection").Columns.Add("Database")
        dsSetup.Tables("LisConnection").Columns.Add("User")
        dsSetup.Tables("LisConnection").Columns.Add("Password")
        dsSetup.Tables("LisConnection").Columns.Add("Connected")

        dsSetup.Tables.Add("Setup")
        dsSetup.Tables("Setup").Columns.Add("AutoStart")
        dsSetup.Tables("Setup").Columns.Add("WriteLog") 
    End Sub

    Sub UpdateSetup()
         
        Try

           

            If AppSetting.GetTestLisConnection(txtLisServer.Text, txtLisDatabase.Text, txtLisUser.Text, txtLisPassword.Text) Then
                blLisConnected = True
                picLis.Image = My.Resources.green_light32
            Else
                blLisConnected = False
                picLis.Image = My.Resources.red_light32
            End If

            My.Settings.Server = txtLisServer.Text
            My.Settings.Database = txtLisDatabase.Text
            My.Settings.User = txtLisUser.Text
            My.Settings.Password = txtLisPassword.Text

            My.Settings.AutoStart = chkAutoStart.Checked
            My.Settings.WriteLog = chkWriteLog.Checked
            My.Settings.WaitConnectSecond = txtWaitConnect.Text
            My.Settings.Connected = blLisConnected
            My.Settings.TrackError = chkTrackLog.Checked

            My.Settings.Save()

            'If System.IO.File.Exists(path) Then
            '    System.IO.File.Delete(path)
            'End If
            'dsSetup.Tables("LisConnection").Rows.Add(txtLisServer.Text, txtLisDatabase.Text, txtLisUser.Text, txtLisPassword.Text, blLisConnected)
            'dsSetup.Tables("Setup").Rows.Add(chkAutoStart.Checked, chkWriteLog.Checked, txtWaitConnect.Text)

            'If chkAutoStart.Checked Then
            '    setting.RunAtStartup(Application.ProductName, Application.ExecutablePath)
            'Else
            '    setting.RemoveValue(Application.ProductName)
            'End If

            'dsSetup.WriteXml(path)
            'dsSetup.Clear()

        Catch ex As Exception

        End Try
    End Sub

    Sub RetrieveSetup()

        'Dim path As String = GetSettingPath()

        'If System.IO.File.Exists(path) Then
        '    dsSetup = GetSettingDataset(path)
        '    dsTemp = dsSetup.Copy()

        '    For Each row As DataRow In dsSetup.Tables("LisConnection").Rows
        '        txtLisServer.Text = row("Server")
        '        txtLisDatabase.Text = row("Database")
        '        txtLisUser.Text = row("User")
        '        txtLisPassword.Text = row("Password")
        '        If row("Connected") Then
        '            picLis.Image = My.Resources.green_light32
        '        Else
        '            picLis.Image = My.Resources.red_light32
        '        End If
        '    Next
        '    For Each row As DataRow In dsSetup.Tables("Setup").Rows
        '        chkAutoStart.Checked = row("AutoStart")
        '        chkWriteLog.Checked = row("WriteLog")
        '        '   txtWaitConnect.Text = row("WaitConnectSecond")

        '    Next
        'End If
         
        If My.Settings.Connected Then
            picLis.Image = My.Resources.green_light32
        Else
            picLis.Image = My.Resources.red_light32
        End If
        txtLisServer.Text = My.Settings.Server
        txtLisDatabase.Text = My.Settings.Database
        txtLisUser.Text = My.Settings.User
        txtLisPassword.Text = My.Settings.Password

        chkAutoStart.Checked = My.Settings.AutoStart
        chkWriteLog.Checked = My.Settings.WriteLog
        txtWaitConnect.Text = My.Settings.WaitConnectSecond
        chkTrackLog.Checked = My.Settings.TrackError
        'dsSetup.Clear()

    End Sub

    Private Sub btCancel_Click(sender As System.Object, e As System.EventArgs) Handles btCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btApply_Click(sender As System.Object, e As System.EventArgs) Handles btApply.Click
        dsTemp = dsSetup.Copy
        UpdateSetup()
    End Sub

    Private Sub btOk_Click(sender As System.Object, e As System.EventArgs) Handles btOk.Click
        UpdateSetup()
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub frmSetup_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        InitDataset()
        RetrieveSetup()

        If chkAutoStart.Checked Then
            txtWaitConnect.Enabled = True
        Else
            txtWaitConnect.Enabled = False
        End If
    End Sub

    Private Sub chkAutoStart_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoStart.CheckedChanged
        If chkAutoStart.Checked Then
            txtWaitConnect.Enabled = True
        Else
            txtWaitConnect.Enabled = False
        End If
    End Sub

    Private Sub txtWaitConnect_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtWaitConnect.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub
End Class