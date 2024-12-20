Imports Microsoft.Win32
Imports System.Security.Permissions
Imports System.Data.SqlClient

Public Class frmLogin

    Private Const SERVER_NAME As String = "ServerName"
    Private Const DATABASE_NAME As String = "DatabaseName"
    Private Const USER_NAME As String = "UserName"
    Public conString As String

    Private Sub btnOk_Click(sender As System.Object, e As System.EventArgs) Handles btnOk.Click
        login()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        UpdateProfile()
        DevLookUpEditProfile.ItemIndex() = 0
    End Sub

    Private Sub UpdateDefaultProfile(ByVal profile As String)

    End Sub

    Function GetProfile() As DataTable

        Dim dt As New DataTable
        dt.Columns.Add("ProfileName")
        dt.Columns.Add("ServerName")
        dt.Columns.Add("DatabaseName")
        dt.Columns.Add("UserName")
        Dim key As String = ""

        Dim keyName As String = "SOFTWARE\LisInterface"
        Dim subkeyNames() As String
        Dim valueNames() As String
        Dim serverName As String = ""
        Dim databaseName As String = ""
        Dim userName As String = ""

        subkeyNames = Registry.LocalMachine.OpenSubKey(keyName).GetSubKeyNames

        For Each key In subkeyNames
            'Debug.Print(key)
            valueNames = Registry.LocalMachine.OpenSubKey(keyName & "\" & key).GetValueNames
            For Each value In valueNames
                'Debug.Print(value & " - " & Registry.LocalMachine.OpenSubKey(keyName & "\" & key).GetValue(value)) 'Gets all keys in the specified subKey
                Select Case value
                    Case "ServerName"
                        serverName = Registry.LocalMachine.OpenSubKey(keyName & "\" & key).GetValue(value)
                    Case "DatabaseName"
                        databaseName = Registry.LocalMachine.OpenSubKey(keyName & "\" & key).GetValue(value)
                    Case "UserName"
                        userName = Registry.LocalMachine.OpenSubKey(keyName & "\" & key).GetValue(value)
                    Case Else
                        End
                End Select
            Next

            Dim row As DataRow = dt.NewRow
            row(0) = key
            row(1) = serverName
            row(2) = databaseName
            row(3) = userName
            dt.Rows.Add(row)
        Next

        Return dt

    End Function

    Private Sub UpdateProfile()
        Dim dtProfile As DataTable = GetProfile()
        'Dim profileName As String = DevLookUpEditProfile.Properties.GetDisplayText(DevLookUpEditProfile.EditValue)

        DevLookUpEditProfile.Properties.DataSource = dtProfile
        DevLookUpEditProfile.Properties.DisplayMember = "ProfileName"
        DevLookUpEditProfile.Properties.ValueMember = "ProfileName"

        'Dim dvProfile As New DataView(dtProfile)

        'dvProfile.RowFilter = "analyzer_ref_cd = '" & profileName & "'"
        'If dvProfile.Count = 1 Then

        '    txtServerName.Text = dvProfile.Item(0)(0)
        '    txtDbName.Text = dvProfile.Item(0)(1)
        '    txtUserName.Text = dvProfile.Item(0)(2)
        'End If

    End Sub

    Private Sub btnAddProfile_Click(sender As System.Object, e As System.EventArgs) Handles btnAddProfile.Click

        Dim f As New frmProfileMaint("", "", "", "")
        If f.ShowDialog = Windows.Forms.DialogResult.OK Then

            txtServerName.Text = f.txtServerName.Text
            txtDbName.Text = f.txtDbName.Text
            txtUserName.Text = f.txtUserName.Text

        End If

        UpdateProfile()
        DevLookUpEditProfile.ItemIndex = DevLookUpEditProfile.Properties.GetDataSourceRowIndex("ProfileName", f.txtProfileName.Text)

    End Sub

    Private Sub btnEditProfile_Click(sender As System.Object, e As System.EventArgs) Handles btnEditProfile.Click

        Dim profileName As String = DevLookUpEditProfile.Properties.GetDataSourceValue("ProfileName", DevLookUpEditProfile.ItemIndex)
        Dim serverName As String = DevLookUpEditProfile.Properties.GetDataSourceValue("ServerName", DevLookUpEditProfile.ItemIndex)
        Dim databaseName As String = DevLookUpEditProfile.Properties.GetDataSourceValue("DatabaseName", DevLookUpEditProfile.ItemIndex)
        Dim userName As String = DevLookUpEditProfile.Properties.GetDataSourceValue("UserName", DevLookUpEditProfile.ItemIndex)

        Dim f As New frmProfileMaint(profileName, serverName, databaseName, userName)
        f.profileName = profileName
        f.serverName = serverName
        f.databaseName = databaseName
        f.userName = userName
        f.txtProfileName.Enabled = False

        If f.ShowDialog = Windows.Forms.DialogResult.OK Then

            txtServerName.Text = f.txtServerName.Text
            txtDbName.Text = f.txtDbName.Text
            txtUserName.Text = f.txtUserName.Text

        End If

        UpdateProfile()
    End Sub

    Private Sub btnDeleteProfile_Click(sender As System.Object, e As System.EventArgs) Handles btnDeleteProfile.Click
        Dim profileName As String = DevLookUpEditProfile.Properties.GetDisplayText(DevLookUpEditProfile.EditValue)
        Dim keyName As String = "SOFTWARE\LisInterface\"

        Dim f As New RegistryPermission( _
        RegistryPermissionAccess.Read Or RegistryPermissionAccess.Write Or RegistryPermissionAccess.Create, _
        "HKEY_LOCAL_MACHINE\" & keyName)

        ' It doesn't exist here. Create the key.
        My.Computer.Registry.LocalMachine.OpenSubKey(keyName, True).DeleteSubKey(profileName)
        UpdateProfile()

        txtServerName.Text = ""
        txtDbName.Text = ""
        txtUserName.Text = ""
        txtPassword.Text = ""

    End Sub

    Private Sub DevLookUpEditProfile_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles DevLookUpEditProfile.EditValueChanged
        txtServerName.Text = DevLookUpEditProfile.Properties.GetDataSourceValue("ServerName", DevLookUpEditProfile.ItemIndex)
        txtDbName.Text = DevLookUpEditProfile.Properties.GetDataSourceValue("DatabaseName", DevLookUpEditProfile.ItemIndex)
        txtUserName.Text = DevLookUpEditProfile.Properties.GetDataSourceValue("UserName", DevLookUpEditProfile.ItemIndex)
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Application.Exit()
    End Sub


    Private Sub txtPassword_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            login()
        End If
    End Sub

    Private Sub txtServerName_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtServerName.KeyDown
        If e.KeyCode = Keys.Enter Then
            login()
        End If
    End Sub

    Private Sub txtDbName_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtDbName.KeyDown
        If e.KeyCode = Keys.Enter Then
            login()
        End If
    End Sub

    Private Sub txtUserName_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtUserName.KeyDown
        If e.KeyCode = Keys.Enter Then
            login()
        End If
    End Sub

    Private Sub login()

        conString = "Data Source=" & txtServerName.Text.Trim & ";Initial Catalog=" & txtDbName.Text.Trim & ";User ID=" & txtUserName.Text.Trim & ";Password=" & txtPassword.Text.Trim & ";"

        Try
            Using connection As New SqlClient.SqlConnection(conString)
                connection.Open()
                If connection.State = ConnectionState.Open Then
                    'MsgBox("OPEN")
                    Me.DialogResult = Windows.Forms.DialogResult.OK
                Else
                    MsgBox("Cannot connect to database !")
                End If

            End Using
        Catch sqex As SqlClient.SqlException
            MessageBox.Show("Sql Exception:  " & sqex.Message)

        Catch ex As DataException
            MessageBox.Show(ex.Message)
        End Try

    End Sub


End Class