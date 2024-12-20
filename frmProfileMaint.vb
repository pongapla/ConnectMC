Imports Microsoft.Win32
Imports System.Security.Permissions

Public Class frmProfileMaint
    'Dim p As Configuration.SettingsProperty
    Private Const SERVER_NAME As String = "ServerName"
    Private Const DATABASE_NAME As String = "DatabaseName"
    Private Const USER_NAME As String = "UserName"

    Public profileName As String
    Public serverName As String
    Public databaseName As String
    Public userName As String

    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click

        Dim keyName As String = "SOFTWARE\LisInterface\" & txtProfileName.Text

        Dim f As New RegistryPermission( _
        RegistryPermissionAccess.Read Or RegistryPermissionAccess.Write Or RegistryPermissionAccess.Create, _
        "HKEY_LOCAL_MACHINE\" & keyName)

        ' Create the registry key object.
        Dim regKey As Object
        Try
            'Check if it exists.  If it doesn't it will throw an error
            regKey = My.Computer.Registry.LocalMachine.OpenSubKey(keyName, True).GetValue(SERVER_NAME)

        Catch ex As Exception
            regKey = Nothing
        End Try

        If regKey Is Nothing Then

            ' It doesn't exist here. Create the key.
            regKey = Registry.LocalMachine.CreateSubKey(keyName)
            ' Next, set the key name and value.
            regKey.SetValue(SERVER_NAME, txtServerName.Text)
            regKey.SetValue(DATABASE_NAME, txtDbName.Text)
            regKey.SetValue(USER_NAME, txtUserName.Text)
            profileName = keyName

            ' Happy message.
            Debug.Print("Registry key added.")
        Else
            My.Computer.Registry.LocalMachine.OpenSubKey(keyName, True).SetValue(SERVER_NAME, txtServerName.Text)
            My.Computer.Registry.LocalMachine.OpenSubKey(keyName, True).SetValue(DATABASE_NAME, txtDbName.Text)
            My.Computer.Registry.LocalMachine.OpenSubKey(keyName, True).SetValue(USER_NAME, txtUserName.Text)
        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Public Sub New(ByVal profileName As String, ByVal serverName As String, ByVal databaseName As String, ByVal userName As String)

        MyBase.New()
        InitializeComponent()

        txtProfileName.Text = profileName
        txtServerName.Text = serverName
        txtDbName.Text = databaseName
        txtUserName.Text = userName

    End Sub

End Class