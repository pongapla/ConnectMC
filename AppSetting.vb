Imports System.Xml
Imports System.Configuration
Imports System.Data.SqlClient
Imports Microsoft.Win32
Imports System.IO

Public Class AppSetting
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Dim FilePath As String
    Dim dsSetup As New DataSet
    Dim dsTemp As New DataSet

    Dim blHisConnected As Boolean
    Dim blLisConnected As Boolean

    Public Shared Function GetTestLisConnection(ByVal server As String, ByVal database As String, ByVal user As String, ByVal password As String) As Boolean
        Dim ret As Boolean
        Dim strSqlCon As String
        Dim objSqlCon As System.Data.SqlClient.SqlConnection = Nothing

        Try
            strSqlCon = String.Format("Data Source={0};User ID={1};Password={2};Initial Catalog={3};", server, user, password, database)
            objSqlCon = New System.Data.SqlClient.SqlConnection(strSqlCon)
            objSqlCon.Open()
            ret = True
        Catch ex As Exception
            ret = False
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetTestLisConnection", ex.Message))
        Finally
            objSqlCon.Close()
            objSqlCon.Dispose()
        End Try

        Return ret
    End Function

    Public Shared Function TestLisConnection() As Boolean
        Dim ret As Boolean
        Dim objSqlCon As System.Data.SqlClient.SqlConnection = Nothing

        Try
            objSqlCon = New System.Data.SqlClient.SqlConnection(GetSqlConnectionString())
            objSqlCon.Open()
            ret = True
        Catch ex As Exception
            ret = False
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "TestLisConnection", ex.Message))
        Finally
            objSqlCon.Close()
            objSqlCon.Dispose()
        End Try

        Return ret
    End Function

    Public Shared Function GetSettingDataset() As DataSet
        Dim ds As New DataSet

        Try
            'If System.IO.File.Exists(filePath) Then
            '    ds.ReadXml(filePath)
            'End If

            ds.Tables.Add("LisConnection")
            ds.Tables("LisConnection").Columns.Add("Server")
            ds.Tables("LisConnection").Columns.Add("Database")
            ds.Tables("LisConnection").Columns.Add("User")
            ds.Tables("LisConnection").Columns.Add("Password")
            ds.Tables("LisConnection").Columns.Add("Connected")

            ds.Tables.Add("Setup")
            ds.Tables("Setup").Columns.Add("AutoStart")
            ds.Tables("Setup").Columns.Add("WriteLog")
            ds.Tables("Setup").Columns.Add("WaitConnectSecond")
            ds.Tables("Setup").Columns.Add("TrackError")

            ds.Tables("LisConnection").Rows.Add(My.Settings.Server, My.Settings.Database, My.Settings.User, My.Settings.Password, My.Settings.Connected)
            ds.Tables("Setup").Rows.Add(My.Settings.AutoStart, My.Settings.WriteLog, My.Settings.WaitConnectSecond, My.Settings.TrackError)




        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetSettingDataset", ex.Message))
        End Try

        Return ds
        ' ValueSoft.CoreControlsLib.CoreCommon.
    End Function

     


    Public Shared Function GetSettingPath()
        Dim ret As String = ""
        Try
            ret = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()) & "\Setup.txt"
        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetSettingPath", ex.Message))
        End Try

        Return ret
    End Function

    Public Shared Function GetSqlConnectionString() As String

        'Dim path As String = GetSettingPath()
        Dim dsSqlCon As New DataSet
        Dim strSqlCon As String = ""

        Try
            'If System.IO.File.Exists(path) Then
            '    dsSqlCon = GetSettingDataset(path)

            '    For Each row As DataRow In dsSqlCon.Tables("LisConnection").Rows
            '        strSqlCon = String.Format("Data Source={0};User ID={1};Password={2};Initial Catalog={3};", row("Server"), row("User"), row("Password"), row("Database"))
            '    Next

            'End If

            dsSqlCon = GetSettingDataset()

            For Each row As DataRow In dsSqlCon.Tables("LisConnection").Rows
                strSqlCon = String.Format("Data Source={0};User ID={1};Password={2};Initial Catalog={3};", row("Server"), row("User"), row("Password"), row("Database"))
            Next
        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetSqlConnectionString", ex.Message))
        End Try

        Return strSqlCon

    End Function

  
  
    Public Shared Function GetAutoStart() As String
        'Dim path As String = GetSettingPath()
        'Dim dsSqlCon As New DataSet
        Dim ret As String = "False"

        Try
            'If System.IO.File.Exists(path) Then
            '    dsSqlCon = GetSettingDataset(path)

            '    For Each row As DataRow In dsSqlCon.Tables("Setup").Rows
            '        ret = row("AutoStart")
            '    Next

            'End If
            If My.Settings.AutoStart Then
                ret = "True"
            Else
                ret = "False"
            End If
 
        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetAutoStart", ex.Message))
        End Try

        Return ret
    End Function

    Public Shared Function GetWriteLog() As String
        'Dim path As String = GetSettingPath()
        'Dim dsSqlCon As New DataSet
        Dim ret As String = "False"

        Try
            'If System.IO.File.Exists(path) Then
            '    dsSqlCon = GetSettingDataset(path)

            '    For Each row As DataRow In dsSqlCon.Tables("Setup").Rows
            '        ret = row("WriteLog")
            '    Next

            'End If
            If My.Settings.WriteLog Then
                ret = "True"
            Else
                ret = "False"
            End If


        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "GetWriteLog", ex.Message))
        End Try

        Return ret
    End Function

    Public Shared Sub UpdateAppSettings(ByVal keyName As String, ByVal keyValue As String)
        Dim xmlDoc As New XmlDocument
        xmlDoc.Load(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)

        Try
            For Each element As XmlElement In xmlDoc.DocumentElement
                If element.Name = "appSettings" Then
                    For Each node As XmlNode In element.ChildNodes
                        If node.Attributes(0).Value = keyName Then
                            node.Attributes(1).Value = keyValue
                        End If
                    Next
                End If
            Next
            xmlDoc.Save(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile)

        Catch ex As Exception
            WriteErrorLog("", "Error", String.Format("{0} : {1}", "UpdateAppSettings", ex.Message))
        End Try
    End Sub

    Public Sub RunAtStartup(ByVal ApplicationName As String, ByVal ApplicationPath As String)
        Dim regKey As Microsoft.Win32.RegistryKey = Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
        With regKey
            .OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            .SetValue(ApplicationName, ApplicationPath)
        End With
        regKey.Close()
    End Sub

    Public Sub RemoveValue(ByVal ApplicationName As String)
        Dim regKey As Microsoft.Win32.RegistryKey = Registry.CurrentUser.CreateSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run")
        With regKey
            .OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Run", True)
            .DeleteValue(ApplicationName, False)
        End With
        regKey.Close()
    End Sub


    Public Shared Function GetVersion() As String
        ' Get the file version for the notepad. 
        ' Use either of the following two commands.
        FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Return myFileVersionInfo.FileVersion
    End Function

    'RunAtStartup(Application.ProductName,Application.ExecutablePath)
    'RemoveValue(Application.ProductName)

    Public Shared Function chkDBNull(ByVal strString As Object) As String
        If IsDBNull(strString) Then
            Return ""
        Else
            Return strString.ToString()
        End If
    End Function

    Public Shared Function WriteErrorLog(ByVal comPort As String, ByVal Role As String, ByVal strMessage As String) As Boolean
        Try

            Dim StartPath As String
            StartPath = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)

            If Not Directory.Exists(StartPath & "\Logs") Then
                Directory.CreateDirectory(StartPath & "\Logs")
            End If

            Dim FileName As String
            If Not comPort.Equals(String.Empty) Then
                FileName = StartPath & "\Logs\" & comPort & "_" & Now.Date.ToString("yyyyMMdd") & ".txt"
            Else
                FileName = StartPath & "\Logs\Error_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
            End If

            If System.IO.File.Exists(FileName) = False Then
                Dim oWrite As StreamWriter
                oWrite = System.IO.File.CreateText(FileName)
                oWrite.Close()
            End If

            Dim writer As New StreamWriter(FileName, True)
            With writer
                .WriteLine(Role & " :")
                .WriteLine(Now.Hour & ":" & Now.Minute & ":" & Now.Second & ">>")
                .WriteLine(strMessage)
                .WriteLine("--------")
                .WriteLine("")
            End With

            writer.Close()
            writer.Dispose()

        Catch ex As Exception
            MessageBox.Show(String.Format("{0} : {1}", "WriteErrorLog", ex.Message))
        End Try

    End Function
    Public Shared Function WriteErrorCal(ByVal comPort As String, ByVal Role As String, ByVal strMessage As String) As Boolean
        Try

            Dim StartPath As String
            StartPath = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)

            If Not Directory.Exists(StartPath & "\Logs") Then
                Directory.CreateDirectory(StartPath & "\Logs")
            End If

            Dim FileName As String
            If Not comPort.Equals(String.Empty) Then
                FileName = StartPath & "\Logs\" & "NotCal_" & comPort & "_" & Now.Date.ToString("yyyyMMdd") & ".txt"
            Else
                FileName = StartPath & "\Logs\Error_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
            End If

            If System.IO.File.Exists(FileName) = False Then
                Dim oWrite As StreamWriter
                oWrite = System.IO.File.CreateText(FileName)
                oWrite.Close()
            End If

            Dim writer As New StreamWriter(FileName, True)
            With writer 
                .WriteLine(strMessage) 
            End With

            writer.Close()
            writer.Dispose()

        Catch ex As Exception
            MessageBox.Show(String.Format("{0} : {1}", "WriteErrorLog", ex.Message))
        End Try

    End Function

    Public Shared Function WriteXml(ByVal dt As DataTable) As Boolean
        Dim StartPath As String

        Try

            StartPath = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)

            If Not Directory.Exists(StartPath & "\XML") Then
                Directory.CreateDirectory(StartPath & "\XML")
            End If

            Dim FileName As String
            FileName = StartPath & "\XML\" & Now.Date.ToString("yyyyMMdd") & ".txt"

        Catch ex As Exception
            WriteErrorLog("", "Error", "ComportMonitor.WriteXml: " & ex.Message)
        End Try
    End Function

End Class
