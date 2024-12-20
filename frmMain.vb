Imports System
Imports System.IO
Imports System.IO.Ports
Imports System.Threading
Imports System.Threading.Thread
Imports Microsoft.Win32
Imports System.Data.SqlClient
Imports DeviceControlApp.ComPortMonitor
Imports DeviceControlApp.AnalyzerInterface
Imports Schedule
Imports System.ComponentModel
Imports System.Configuration
Imports System.Configuration.ConfigurationErrorsException
Imports ValueSoft.Common
Imports ValueSoft.CoreControlsLib
Imports ValueSoft.DALManage
'System.Configuration.ConfigurationErrorsException' occurred in System.Configuration.dll


Public Class frmMain

   
    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Dim WithEvents Srv As ServerClass

    Dim HaveInterface As Boolean = False
    'Dim _dbdataset As DBDataset.DB
    Dim clsMon As ComPortMonitor
    Public _objCon As SqlConnection

    Dim SpaceCount As Byte = 0
    Dim LookUpTable As String = "0123456789ABCDEF"
    Dim RXArray(2047) As Char ' Text buffer. Must be global to be accessible from more threads.
    Dim RXCnt As Integer      ' Length of text buffer. Must be global too.

    Dim FilePath As String
    ' Make a new System.IO.Ports.SerialPort instance, which is able to fire events.
    Dim WithEvents COMPort As New SerialPort

    Dim dsDevice As New DataSet
    Dim sqlServerCon As String
    Dim autoStart As Boolean

    Private Sub frmMain_Closing(ByVal sender As Object, ByVal e As ComponentModel.CancelEventArgs) Handles MyBase.Closing
        StopMonitor()

        If lvDevice.Items.Count = 0 Then
            System.IO.File.Delete(FilePath)
            Exit Sub
        End If

        dsDevice = New DataSet
        dsDevice.Tables.Add("Device")

        dsDevice.Tables("Device").Columns.Add("Port")
        dsDevice.Tables("Device").Columns.Add("AnalyzerCd")
        dsDevice.Tables("Device").Columns.Add("AnalyzerDesc")
        dsDevice.Tables("Device").Columns.Add("AnalyzerSkey")
        dsDevice.Tables("Device").Columns.Add("SerialNo")
        dsDevice.Tables("Device").Columns.Add("AnalyzerModel")

        'dsDevice.Tables("Device").Rows.Clear()
        Dim row As DataRow

        For Each item As ListViewItem In lvDevice.Items
            row = dsDevice.Tables("Device").NewRow
            row.Item(0) = item.SubItems(1).Text
            row.Item(1) = item.SubItems(2).Text
            row.Item(2) = item.SubItems(3).Text
            row.Item(3) = item.SubItems(4).Text
            row.Item(4) = item.SubItems(5).Text
            row.Item(5) = item.SubItems(6).Text

            dsDevice.Tables("Device").Rows.Add(row)
        Next
        dsDevice.WriteXml(FilePath)
    End Sub

    Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        BackgroundWorker1.WorkerReportsProgress = True

        BackgroundWorker1.WorkerSupportsCancellation = True
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If MessageBox.Show("Do you want to Exit ?", "Important Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
            e.Cancel = True
        End If
    End Sub

     
    Private WithEvents aTimer As New System.Windows.Forms.Timer
    Private WithEvents aTimerStripStatus As New System.Windows.Forms.Timer
    Private TargetDT As DateTime
    Private CountDownFrom As TimeSpan = TimeSpan.FromMilliseconds(My.Settings.WaitConnectSecond * 1000)
    Private Sub timeTick(sender As Object, e As System.EventArgs) Handles aTimer.Tick

        Dim ts As TimeSpan = TargetDT.Subtract(DateTime.Now)
        If ts.TotalMilliseconds > 0 Then

            btnStart.Text = "Starting " & ts.ToString("mm\:ss")
            enabledStartAndStop("Starting")
        Else

            enabledStartAndStop("Started")
        End If
         
    End Sub
    Private Sub timeStripStatusTick(sender As Object, e As System.EventArgs) Handles aTimerStripStatus.Tick

        ToolStripStatusLabel1.Text = System.DateTime.Now

    End Sub

    Sub enabledStartAndStop(Status As String)
        Select Case Status
            Case "Starting"
                btnStart.ForeColor = Color.Gray
                btnStart.Enabled = False

                btnStop.Enabled = False
                btnStop.ForeColor = Color.Gray
            Case "Started"
                btnStart.Text = "Start Monitor"
                btnStart.ForeColor = Color.Black
                btnStart.Enabled = True
                btnStop.ForeColor = Color.Black
                btnStop.Enabled = True
            Case "Stoped"
                btnStart.Text = "Start Monitor"
                btnStart.ForeColor = Color.Gray
                btnStart.Enabled = True

                btnStop.Enabled = False
                btnStart.ForeColor = Color.Red
        End Select

    End Sub
    

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If My.Settings.CallUpgrade Then
                My.Settings.Upgrade()
                My.Settings.CallUpgrade = False
                My.Settings.Save()
            End If
        Catch ex As System.Configuration.ConfigurationException
            My.Computer.FileSystem.DeleteFile(CType(ex.InnerException, ConfigurationErrorsException).Filename)

            If (Process.GetProcessesByName(Process.GetCurrentProcess.ProcessName).Length < 3) Then
                System.Diagnostics.Process.Start(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
            End If

            Process.GetCurrentProcess().Kill()
        End Try

        ValueSoft.CoreControlsLib.VariableClass.ServerName = My.Settings.Server
        ValueSoft.CoreControlsLib.VariableClass.DatabaseName = My.Settings.Database
        ValueSoft.CoreControlsLib.VariableClass.UserID = My.Settings.User
        ValueSoft.CoreControlsLib.VariableClass.Password = My.Settings.Password 'ValueSoft.Common.Utility.Decrypt(My.Settings.Password)
        ValueSoft.CoreControlsLib.VariableClass.SQLUserID = My.Settings.User
        ValueSoft.CoreControlsLib.VariableClass.SQLPassword = My.Settings.Password 'ValueSoft.Common.Utility.Decrypt(My.Settings.Password)

        My.Settings.Reload()
        aTimerStripStatus.Interval = 1000 'The number of miliseconds in a second
        aTimerStripStatus.Enabled = True 'Start the timer
        If My.Settings.TrackError Then
            log.Info("Start")
        End If
        

        Me.KeyPreview = True
       

        If AppSetting.TestLisConnection = False Then
            SetFormEnable(False)
            Dim frmSetup As New frmSetup
            If frmSetup.ShowDialog = Windows.Forms.DialogResult.OK Then
                sqlServerCon = AppSetting.GetSqlConnectionString()
                If AppSetting.TestLisConnection Then
                    SetFormEnable(True)
                    MyGlobal.myConnectionString = AppSetting.GetSqlConnectionString()
                Else
                    WriteLog("ConnectionStatus", "Connection", " Connect Server Database Fail. (" & String.Format("Server : {0}, Database : {1}, User : {2}, AutoStart : {3}, WaitConnectSecond : {4},WriteLog : {5} ", My.Settings.Server, My.Settings.Database, My.Settings.User, My.Settings.AutoStart, My.Settings.WaitConnectSecond, My.Settings.WriteLog) & ")") 'Golf Mark

                End If
            End If
        Else
            SetFormEnable(True)
            MyGlobal.myConnectionString = AppSetting.GetSqlConnectionString()

              WriteLog("ConnectionStatus", "Connection", " Connect Server Database Success. (" & String.Format("Server : {0}, Database : {1}, User : {2}, AutoStart : {3}, WaitConnectSecond : {4},WriteLog : {5} ", My.Settings.Server, My.Settings.Database, My.Settings.User, My.Settings.AutoStart, My.Settings.WaitConnectSecond, My.Settings.WriteLog) & ")") 'Golf Mark

        End If

        Dim item As ListViewItem

        lvDevice.Items.Clear()

        Try
            FilePath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location()) & "\Interface.txt"

            If System.IO.File.Exists(FilePath) Then
                dsDevice.ReadXml(FilePath)
                For Each row As DataRow In dsDevice.Tables("Device").Rows
                    item = lvDevice.Items.Add("")
                    item.SubItems.Add(row.Item(0))
                    item.SubItems.Add(row.Item(1))
                    item.SubItems.Add(row.Item(2))
                    item.SubItems.Add(row.Item(3))
                    item.SubItems.Add(row.Item(4))
                    item.SubItems.Add(row.Item(5))
                    item.ImageIndex = 0
                Next

                If lvDevice.Items.Count > 0 Then
                    lvDevice.Items(0).Selected = True
                    'ListView1.TopItem.Selected = True  
                End If
            End If
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Error", "Form.Main: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Me.Text = Me.Text & " - v " & GetVersion()


        '1 Timer coments 2019-01-04
        'StartTimerAndStart(True)

        'Dim TimerRestart = New ScheduleTimer()
        'AddHandler TimerRestart.Elapsed, New ScheduledEventHandler(AddressOf timer_Elapsed)
        'TimerRestart.AddEvent(New ScheduledTime("Daily", "3:00"))
        'TimerRestart.AddEvent(New ScheduledTime("Daily", "3:01"))
        'TimerRestart.AddEvent(New ScheduledTime("Daily", "3:10"))
        'TimerRestart.Start()
        ' Timer 

        '2 Not Timer  2019-02-01
        StartTimerAndStart(False) 'Golf 2019-02-01 For พังโคน

        If My.Settings.TrackError Then
            log.Info("End")
        End If
    End Sub

    Dim arMon As New ArrayList

    Private Sub ChangeWaitState()

        For Each item As ListViewItem In lvDevice.Items
            item.ImageIndex = 1
        Next
        Application.DoEvents()
    End Sub

    Public Sub StopMonitor()
        ChangeWaitState()

        Dim mon As ComPortMonitor
        For i As Integer = arMon.Count - 1 To 0 Step -1
            mon = arMon.Item(i)
            mon.Closing()
            arMon.RemoveAt(i)
        Next

        For Each item As ListViewItem In lvDevice.Items
            item.ImageIndex = 0
            Application.DoEvents()
        Next

        StartTimerAndStart(False)

    End Sub



    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        If BackgroundWorker1.IsBusy <> True Then

            ' Start the asynchronous operation.

            BackgroundWorker1.RunWorkerAsync()

            StopMonitor()
            Thread.Sleep(200)
            Start()
            ' enabledStartAndStop("Started")
        End If
        'StopMonitor()
        'Thread.Sleep(200)
        'Start()
    End Sub


    Public Sub StartTimerAndStart(loaded As Boolean)
        autoStart = AppSetting.GetAutoStart


        If autoStart = True Then
            If loaded Then
                Start()
            End If

            If My.Settings.WaitConnectSecond > 0 Then

                If Not (loaded) Then

                    TargetDT = DateTime.Now.Add(CountDownFrom)
                    aTimer.Interval = 100 'ten times per second
                    aTimer.Start() 'start the timer
                End If

                Timer1.Start()
                Timer1.Interval = My.Settings.WaitConnectSecond * 1000
                Timer1.Enabled = True
            Else
                Timer1.Stop()
                Timer1.Enabled = False
            End If
        Else
            Timer1.Stop()
            Timer1.Enabled = False
        End If
    End Sub

    Public Sub Start()
        ChangeWaitState()

        For Each item As ListViewItem In lvDevice.Items

            clsMon = New ComPortMonitor(item.SubItems(1).Text, item.SubItems(6).Text, Int32.Parse(item.SubItems(4).Text))

            If clsMon.ErrorOccur Then
                item.ImageIndex = 0
                WriteLog(item.SubItems(1).Text, "Error", "ComportMonitor.Error Start(): " & clsMon.ErrorMessage) 'Golf Mark
                ' StartTimerAndStart()

            Else
                item.ImageIndex = 2
            End If
            arMon.Add(clsMon)
            Application.DoEvents()
        Next

    End Sub

    Private Sub SendData(ByVal Value As Object)
        Dim mon As ComPortMonitor
        For i As Integer = arMon.Count - 1 To 0 Step -1
            mon = arMon.Item(i)
            mon.SendData(Value)
        Next
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        StopMonitor()
        ' enabledStartAndStop("Stoped")
    End Sub

    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        If lvDevice.SelectedItems.Count <> 1 Then
            MsgBox("Please Select Device to Remove", MsgBoxStyle.Information, "Information")
        Else
            If MsgBox("Do you to remove device?", MsgBoxStyle.OkCancel, "Information") = MsgBoxResult.Ok Then
                lvDevice.SelectedItems.Item(0).Remove()
            End If
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Dim count As Int16 = lvDevice.Items.Count
        Dim comports(count) As String

        For i = 0 To count - 1
            comports(i) = lvDevice.Items(i).SubItems(1).Text
        Next

        Dim f As New frmAddDevice
        f.comPorts = comports

        If f.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim item As ListViewItem
            item = lvDevice.Items.Add("")
            item.SubItems.Add(f.cbComPort.Text)
            item.SubItems.Add(f.cbAnalyzerCd.Text)
            item.SubItems.Add(f.txtAnalyzerDesc.Text)
            item.SubItems.Add(f.lbAnalyzerSkey.Text)
            item.SubItems.Add(f.lbSerialNo.Text)
            item.SubItems.Add(f.lbAnalyzerModel.Text)
            item.ImageIndex = 0
        End If
    End Sub

    Private Sub tmRS232_tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmRS232.Tick
        SendData(Chr(6))
    End Sub

    Private Sub btnSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetup.Click
        ' FrmShowKeyTest 

        Dim frmSetup As New frmSetup
        If frmSetup.ShowDialog = Windows.Forms.DialogResult.OK Then
            sqlServerCon = AppSetting.GetSqlConnectionString()
            MyGlobal.myConnectionString = AppSetting.GetSqlConnectionString()

            'autoStart = AppSetting.GetAutoStart 
            'If autoStart = True Then
            '    StopAndStart()

            '    If My.Settings.WaitConnectSecond > 0 Then
            '        Timer1.Start()
            '        Timer1.Interval = My.Settings.WaitConnectSecond * 1000
            '        Timer1.Enabled = True
            '    Else
            '        Timer1.Enabled = False
            '    End If
            'Else
            '    Timer1.Enabled = False
            'End If
            CountDownFrom = TimeSpan.FromMilliseconds(My.Settings.WaitConnectSecond * 1000)
        End If

        If AppSetting.TestLisConnection = True Then
            SetFormEnable(True)
        Else
            SetFormEnable(False)
        End If

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

        Try

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

        Catch ex As Exception
            MsgBox("Error : " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return dt

    End Function

    Public Sub AppendText(text As String)
        Try

            If lbMessageService.Items.Count > 1000 Then
                lbMessageService.Items.Clear()
            End If

            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of String)(AddressOf AppendText), New Object() {text})
                Return
            End If
            lbMessageService.Items.Add(text)

            Dim visibleItems As Integer = lbMessageService.ClientSize.Height / lbMessageService.ItemHeight

            lbMessageService.TopIndex = Math.Max(lbMessageService.Items.Count - visibleItems + 1, 0)


        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Error", "Form.Main: " & ex.Message)
            log.Error(ex.Message)
        End Try

    End Sub

    Public Sub AppendResult(dtResult As DataTable)
        Try

            Dim formattedDate As String = DateTime.Now()
            If lvResult.Items.Count > 50 Then
                lvResult.Items.Clear()
            End If

            If Me.InvokeRequired Then
                Me.Invoke(New Action(Of DataTable)(AddressOf AppendResult), New Object() {dtResult})
                Return
            End If

            For Each row As DataRow In dtResult.Rows
                'btnLog.Rows.Add(0, formattedDate, row("result_date"), row("order_skey"), row("order_id").ToString, row("specimen_type_id").ToString, row("analyzer_skey").ToString, row("analyzer_cd").ToString, row("result_item_skey").ToString, row("analyzer_ref_cd"), row("alias_id"), row("result_value"), row("sending_status"))

                Dim item As ListViewItem
                item = lvResult.Items.Add("")
                item.SubItems.Add(formattedDate.ToString())
                item.SubItems.Add(row("result_date").ToString())
                item.SubItems.Add(row("order_id").ToString())
                item.SubItems.Add(row("specimen_type_id").ToString())
                item.SubItems.Add(row("analyzer_cd").ToString())
                item.SubItems.Add(row("analyzer_ref_cd").ToString())
                item.SubItems.Add(row("alias_id").ToString())
                item.SubItems.Add(row("result_value").ToString())

            Next

            'If dtResult.Rows.Count > 0 Then
            '    'btnResend.Enabled = True
            'End If

        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Error", "Form.Main: " & ex.Message)
            log.Error(ex.Message)
        End Try

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click
        Clear()
    End Sub

    Sub Clear()
        lvResult.Items.Clear()
        'btnResend.Enabled = False
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        lbMessageService.Items.Clear()
    End Sub

    Public Shared Function GetVersion() As String
        ' Get the file version for the notepad. 
        ' Use either of the following two commands.
        FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        Return myFileVersionInfo.FileVersion
    End Function

    Public Sub SetFormEnable(ByVal flag As Boolean)
        If flag = True Then
            btnStart.Enabled = True
            btnStop.Enabled = True
            btnAdd.Enabled = True
            btnRemove.Enabled = True
        Else
            btnStart.Enabled = False
            btnStop.Enabled = False
            btnAdd.Enabled = False
            btnRemove.Enabled = False
        End If
    End Sub

    Private Sub btnOpenLog_Click(sender As System.Object, e As System.EventArgs) Handles btnOpenLog.Click
        Process.Start("explorer.exe", Path.GetDirectoryName(Application.ExecutablePath) & "\Logs")
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try

            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\rxlyte2.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 4, "RxLyte")
            a.ExtractResult_RxLyte(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test RxLyte", ex.Message)
            log.Error(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\AS720.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 6, "AS720")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test AS720", ex.Message)
            log.Error(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Try
             

            Dim a As New AnalyzerInterface("", 45, "Architecti1000")
            'a.ReturnOrderDetailArchitecti1000("180316129", 0)
            a.ReturnOrderDetailArchitecti1000("180316130", 0)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test Imola ", ex.Message)
        End Try
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Try
            Dim file1 As FileInfo = New FileInfo("C:\Users\Tui\Desktop\Uriscan.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 10, "Uriscan")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test Uriscan ", ex.Message)
        End Try
    End Sub

    Private Sub btnCobas_Click(sender As Object, e As EventArgs) Handles btnCobas.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS411_1.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 0, "COBAS_e411")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test Uriscan ", ex.Message)
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\Analyzer Text\mythic2.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 23, "Mythic")
            a.ExtractResult(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test Mythic ", ex.Message)
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\daytona3.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 22, "Daytona")
            a.ExtractResult(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test Daytona ", ex.Message)
        End Try
    End Sub

    Private Sub btnLd500_Click(sender As Object, e As EventArgs) Handles btnLd500.Click
        Try

            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\LD500_QC1.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\LD500_2.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 26, "LD-500")
            a.ExtractResult_HbA1c(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test LD-500", ex.Message)
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'Golf Create Event 2016-04-04
        Try
            '  AppSetting.WriteErrorLog(COMPort, "DX300", String.Join("", returnArray.ToArray()))
            AppSetting.WriteErrorLog("", "DX300", "TESTGolf")
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\H-500_15.txt")

            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 29, "H-500") '170002532
            'Dim a As New AnalyzerInterface("", 17, "H-500")
            'Dim a As New AnalyzerInterface("", 16, "H-500")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
        End Try
        'Golf Create Event 2016-04-04
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        'Golf Create Event 2016-04-04
        Try
            '  AppSetting.WriteErrorLog(COMPort, "DX300", String.Join("", returnArray.ToArray()))
            AppSetting.WriteErrorLog("", "DX300", "TESTGolf")
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\H-500_15.txt")

            Dim script As String = file1.OpenText().ReadToEnd()

            ' Dim a As New AnalyzerInterface("", 29, "H-500")
            'Dim a As New AnalyzerInterface("", 17, "H-500")
            Dim a As New AnalyzerInterface("", 16, "H-500")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
        End Try
        'Golf Create Event 2016-04-04
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\DX300.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 30, "DX300")
            'Dim a As New AnalyzerInterface("", 30, "DX300")
            a.ExtractResult_Chem(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test DX300-001 ", ex.Message)
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300Receiver.txt")
        'Dim LabArr As New ArrayList 
        'Dim ResultStr As String = file1.OpenText().ReadToEnd()
        'Debug.WriteLine(ResultStr)

        ''Extract Result 
        'Dim LineArr() As String
        'Dim ItemStr As String
        'Dim ItemArr() As String
        'Dim LabNo As String = ""
        'Dim InterfaceStatus As String = ""
        'Dim j As Integer

        'LineArr = ResultStr.Split(Chr(10))
        'j = LineArr.Count

        'For i As Integer = 0 To j - 1
        '    ItemStr = LineArr(i)
        '    ItemArr = ItemStr.Split("|")
        '    If ItemArr.Length >= 3 Then
        '        If ItemArr(2).Length >= 9 Then
        '            InterfaceStatus = "Question" 
        '            LabNo = ItemArr(2)
        '            'LabNo = ItemArr(2)
        '            LabArr.Add(LabNo)
        '        End If
        '    End If
        '    'LabArr.Add(LabNo)
        'Next
        'Dim LabArr2 As ArrayList = LabArr
        'Dim returnArray As New ArrayList
        'Dim AnalyzerInterface As New DeviceControlApp.AnalyzerInterface("", 30, "DX300")
        'returnArray = AnalyzerInterface.ReturnOrderDetailDX300(LabArr) 
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\U120.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 16, "U120")
            'Dim a As New AnalyzerInterface("", 30, "DX300")
            a.ExtractResult_Micro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test U120 ", ex.Message)
        End Try
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\SUZUKA.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 32, "SUZUKA")
            'Dim a As New AnalyzerInterface("", 30, "DX300")
            a.ExtractResult_Chem(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test SUZUKA-001 ", ex.Message)
        End Try
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\LIAISON.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 35, "Liaison")
            'Dim a As New AnalyzerInterface("", 30, "DX300")
            a.ExtractResult_Chem(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Liaison ", ex.Message)
        End Try
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\Liaison5.txt")
            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dx300.txt")
            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 32, "Liaison")
            a.ExtractResult_Chem(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Liaison ", ex.Message)
        End Try


    End Sub


   

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\ISE6000.txt")
            Dim script As String = file1.OpenText().ReadToEnd()
            Dim a As New AnalyzerInterface("", 33, "ISE6000")
            a.ExtractResult_ISE6000(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test ISE6000", ex.Message)
        End Try
    End Sub

    Private Sub lvResult_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvResult.SelectedIndexChanged

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Try
            Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\HumaClot Pro\fileresult form_analyzer\re_170391704_PT-SI_0041.csv")

            Dim script As String = file1.OpenText().ReadToEnd()

            Dim a As New AnalyzerInterface("", 15, "humaclot_pro")
            a.ExtractResult_humaclot_pro(script)
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test humaclot_pro", ex.Message)
        End Try
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        Try
            Dim ResultStr As String
            ResultStr = ""
            If ResultStr.IndexOf(Chr(5)) <> -1 Then
                'Have ENQ
                'Check EOT
                'send ACK 
                If (ResultStr.IndexOf(Chr(4)) = -1) Or (ResultStr.Substring(ResultStr.Length - 1) = Chr(5)) Then
                    Dim ResultStr2 As String = ResultStr.Substring(ResultStr.Length - 1)
                End If
                Dim ResultStr3 As String = ResultStr.Substring(ResultStr.Length - 1)
            End If

            'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\modana2.txt")
            'Dim script As String = file1.OpenText().ReadToEnd()


            'Dim a As New AnalyzerInterface("", 18, "RXMODENA")
            'a.ExtractResult_Chem(script)


        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test RXMODENA ", ex.Message)
        End Try
    End Sub



    

    Dim count As Integer
    Dim EndCount As Integer
    Dim returnArray As New ArrayList
    Public Sub SendAnswer(line As String)
        COMPort.Write(line)
        count = count + 1
    End Sub


    Sub StopAndStart()
        StopMonitorStopAndStart()
        Thread.Sleep(200)
        StartStopAndStart()

        'If StopMonitorStopAndStart() Then
        '    WriteLog(COMPort.PortName, "Error", "ComportMonitor.Error StopAndStart(): connect") 'Golf Mark
        '    Thread.Sleep(200)
        '    StartStopAndStart()
        'Else
        '    WriteLog(COMPort.PortName, "Error", "ComportMonitor.Error StopAndStart(): Disconnect") 'Golf Mark
        'End If


    End Sub


    Private Sub timer_Elapsed(sender As Object, e As ScheduledEventArgs)


        StopAndStart()

        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\imola14.txt")
        'Dim script As String = file1.OpenText().ReadToEnd()

        'Dim a As New AnalyzerInterface("", 21, "Imola")
        'a.ExtractResult_Chem(script)

    End Sub

    Private Sub StopMonitorStopAndStart()
        ChangeWaitStateStopAndStart()

        Dim mon As ComPortMonitor

        For i As Integer = arMon.Count - 1 To 0 Step -1
            mon = arMon.Item(i)
            mon.Closing()
            arMon.RemoveAt(i)
        Next

        ChangeWaitStateStopAndStart0()

    End Sub

    Private Sub ChangeWaitStateStopAndStart()

        Dim listview2 As ListView = Nothing
        If lvDevice.InvokeRequired Then
            lvDevice.Invoke(New Action(Of ListView)(AddressOf ChangeWaitStateStopAndStart), lvDevice)

        Else
            For Each item As ListViewItem In lvDevice.Items
                item.ImageIndex = 1
            Next
        End If

        Application.DoEvents()
    End Sub

    Private Sub ChangeWaitStateStopAndStart0()

        Dim listview2 As ListView = Nothing
        If lvDevice.InvokeRequired Then
            lvDevice.Invoke(New Action(Of ListView)(AddressOf ChangeWaitStateStopAndStart0), lvDevice)

        Else
            For Each item As ListViewItem In lvDevice.Items
                item.ImageIndex = 0
            Next
        End If

        Application.DoEvents()
    End Sub

    Private Sub StartStopAndStart()
        ChangeWaitStateStopAndStart()

        If lvDevice.InvokeRequired Then
            lvDevice.Invoke(New Action(Of ListView)(AddressOf StartStopAndStart), lvDevice)
        Else
            For Each item As ListViewItem In lvDevice.Items

                clsMon = New ComPortMonitor(item.SubItems(1).Text, item.SubItems(6).Text, Int32.Parse(item.SubItems(4).Text))

                If clsMon.ErrorOccur Then
                    item.ImageIndex = 0
                    WriteLog(item.SubItems(1).Text, "Error", "ComportMonitor.Error StartStopAndStart(): " & clsMon.ErrorMessage) 'Golf Mark 
                    StartStopAndStart()
                    'If autoStart = True Then
                    '    Timer1.Start()
                    'End If
                Else
                    item.ImageIndex = 2
                End If
                arMon.Add(clsMon)
                Application.DoEvents()
            Next
        End If




    End Sub

    Private Delegate Function GetItems(lstview As ListView) As ListView.ListViewItemCollection

    Private Function getListViewItems(lstview As ListView) As ListView.ListViewItemCollection
        Dim temp As New ListView.ListViewItemCollection(New ListView())
        If Not lstview.InvokeRequired Then
            For Each item As ListViewItem In lstview.Items
                temp.Add(DirectCast(item.Clone(), ListViewItem))
            Next
            Return temp
        Else
            Return DirectCast(Me.Invoke(New GetItems(AddressOf getListViewItems), New Object() {lstview}), ListView.ListViewItemCollection)
        End If
    End Function

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim ResultStr As String = "R099999999961757610534019120717113AHDL75mg/dL2TGL79R099999999961757610534019120717113AHDL75mg/dL2TGL79mg/dL6CHOL198mg/dLCA"
        If ResultStr.IndexOf(Chr(2)) <> -1 And ResultStr.IndexOf(Chr(3)) <> -1 Then
            Dim aaaa As String = ResultStr.Reverse.ToString

            Dim a As Integer = ResultStr.Reverse().ToString.IndexOf(Chr(3))
            Dim aa As Integer = ResultStr.Reverse().ToString.IndexOf(Chr(2))
            ResultStr = ResultStr.Reverse().ToString.Substring(ResultStr.Reverse().ToString.IndexOf(Chr(3)), ResultStr.Reverse().ToString.IndexOf(Chr(2))).Reverse()

        End If
        'Check STX
        'Dim sss As String = "P"
        ''Dim sss As String = "1P"
        'Dim sss2 As String = "P"
        ''Dim ResultStr As String = "P23450103B"
        'Dim ResultStr As String = "I1700243374C"
        'Dim LineArr() As String
        'LineArr = ResultStr.Split(Chr(28))
        'Dim LabNo As String = LineArr(1)
        'Dim FIRSTPOLL As String = LineArr(2)
        'Dim REQUEST As String = LineArr(3)
        'If sss.Length > 1 And sss.IndexOf(Chr(2)) <> -1 Then
        '    sss2 = "3"
        'End If
        'Try 
        '    Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\modanaq1.txt")
        '    Dim script As String = file1.OpenText().ReadToEnd()
        '    Dim Str As String
        '    Dim StrArr As New ArrayList

        '    'Not Repeated for each known carrier
        '    'Str = Chr(2) & Chr(28) & "P" & Chr(28) & "92300" & Chr(28) & "1" & Chr(28) & "1" & Chr(28) & "0" & Chr(28)
        '    'Str &= GetCheckSumValue(Str.Remove(0, 1))
        '    'Str &= Chr(3)
        '    'StrArr.Add(Str)


        '    'Repeated for each known carrier
        '    Str = Chr(2) & Chr(28) & "P" & Chr(28) & "92300" & Chr(28) & "1" & Chr(28) & "1" & Chr(28) & "0" & Chr(28) & "A" & Chr(28)
        '    Str &= GetCheckSumValue(Str.Remove(0, 1))
        '    Str &= Chr(3)
        '    StrArr.Add(Str)

        '    For Each line As String In StrArr
        '        script = line
        '        ReceiverDimension(script)
        '    Next

        'Catch ex As Exception
        '    AppSetting.WriteErrorLog("", "Dimension", ex.Message)
        'End Try
    End Sub

    Public Function GetCheckSumValue(ByVal frame As String) As String
        Dim checksum As String = "00"

        Dim byteVal As Integer = 0
        Dim sumOfChars As Integer = 0
        Dim complete As Boolean = False

        Try

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
                checksum = Convert.ToString(sumOfChars Mod 256, 16).ToUpper()
            End If

        Catch ex As Exception
            'AppSetting.WriteErrorLog(comPort, "Error", "AnalyzerInterface.GetCheckSumValue: " & ex.Message)
            log.Error(ex.Message)
        End Try

        Return DirectCast(If(checksum.Length = 1, "0" & checksum, checksum), String)
    End Function
   

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        ' Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dimension1.txt")
        Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\dimension_c.txt")
        'dimension_c.txt()
        Dim script As String = file1.OpenText().ReadToEnd()
        'Dim Str As String
        'Dim StrArr As New ArrayList
        'Dim FSBlank As String = Chr(28) & "" & Chr(28) 
        'Str = Chr(2) & "R" & Chr(28) & "" & Chr(28) & "" & Chr(28) & "1596" & Chr(28) & "1" & Chr(28) & "" & Chr(28) & "0" & Chr(28) & "170001264" & Chr(28) & "1"
        'Str &= Chr(28) & "1" & Chr(28) & "5" & Chr(28) & "GLU" & Chr(28) & "86.00" & Chr(28) & "mg/dL" & Chr(28) & "" & Chr(28) & "BUN" & Chr(28) & "8" & Chr(28) & "mg/dL" & Chr(28) & "" & Chr(28)
        'Str &= GetCheckSumValue(Str.Remove(0, 1))
        'Str &= Chr(3)
        'script = Str

        Dim a As New AnalyzerInterface("", 20, "Dimension") ' iSLIM_SRS
        a.ExtractResult(script)
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        ' Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_3.txt")
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\C11_HBA1c.txt")
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_4.txt")
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_8.txt")
        Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_9.txt")

        ' Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_11.txt")
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_10.txt") 'rerun

        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_12.txt")
        'Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_13.txt") 'rerun

        ' Dim file1 As FileInfo = New FileInfo(" D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_14.txt") 'rerun


        Dim script As String = file1.OpenText().ReadToEnd()
        Dim Str As String
        Dim StrArr As New ArrayList

        Dim a As New AnalyzerInterface("", 19, "COBAS_C311")
        a.ExtractResult(script)
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        'Dim returnArray As New ArrayList
        Dim file1 As FileInfo = New FileInfo("D:\Project\AnalyzerInterface_Test_TextFile\COBAS_C311_q2.txt")
        Dim script As String = file1.OpenText().ReadToEnd()
        Dim Str As String
        Dim StrArr As New ArrayList
        Dim LineArr() As String
        Dim ItemStr As String
        Dim ItemArr() As String
        Dim LabNo As String

        Dim ResultStr As String = script

        Dim a As New AnalyzerInterface("", 19, "COBAS_C311")

        LineArr = ResultStr.Split(Chr(13))

        Dim j As Integer

        j = LineArr.Count
        Dim LabNoArrRes() As String
        For i As Integer = 0 To j - 1
            ItemStr = LineArr(i)
            ItemArr = ItemStr.Split("|")

            Select Case ItemArr(0)
                Case "H"
                    'Head Line 

                Case "Q"
                    If ItemArr.Length >= 3 Then
                        If ItemArr(2).Length >= 3 Then
                            ItemArr(2) = ItemArr(2) + "^R"
                            LabNo = ItemArr(2)


                            LabNoArrRes = ItemArr(2).Split("^")
                            LabNo = LabNoArrRes(2).Trim

                            Dim SpecimenIdString As String = ItemArr(2)
                            Dim SpecimenId As String
                            Dim SpecimenIdArr() As String
                            Dim labSeq As String
                            Dim RackIDNo As String

                            Dim PositionNo As String
                            Dim SampleType As String
                            Dim ContainerType As String
                            SpecimenIdArr = SpecimenIdString.Split("^")
                            If SpecimenIdArr.Length > 0 Then
                                SpecimenId = SpecimenIdArr(2).Trim
                                labSeq = SpecimenIdArr(3)
                                RackIDNo = SpecimenIdArr(4)
                                PositionNo = SpecimenIdArr(5)
                                SampleType = SpecimenIdArr(7)
                                ContainerType = SpecimenIdArr(8)
                            End If



                            'Example Result Value -1^0.533
                            'itemArrRes = ItemArr(3).Split("^")
                            'If itemArrRes.Count > 1 Then
                            '    ItemArr(3) = itemArrRes(1).Trim
                            'End If


                        End If
                        Dim Priority As String
                        Select Case ItemArr(12)
                            Case "O"
                                Priority = "R"
                            Case "A"
                                Priority = "S"
                            Case Else
                                Priority = "R"
                        End Select
                    End If
                    'Exit For
                Case "L"
                    'Last Line

                Case "C"
                    'Lab Detail Line
            End Select

        Next


        '170002749
        ' returnArray = ReturnOrderDetail("^^             170000767^0^50042^042^^S0^")





        Dim str1 As String

        str1 = Chr(2) & "1H|\^&|||host^1|||||cobasc311|TSDWN^REPLY|P|1"
        str1 &= Chr(13)
        str1 &= "P|1|"
        str1 &= Chr(13)
        str1 &= "O|1|            0600085542|0^50024^024^^S0^SC|^^^668^\^^^421^\^^^452^\^^^700^\^^^798^\^^^781^\^^^435^\^^^552^\^^^678^\^^^413^\^^^712^\^^^734^\^^^683^\^^^687^\^^^685^\^^^698^\^^^714^\^^^701^\"
        str1 &= Chr(23)
        Dim StrAll As String = str1
        Dim checksumvalue2 As String = GetCheckSumValue(str1.Remove(0, 1))
        str1 &= GetCheckSumValue(str1.Remove(0, 1))
        str1 &= Chr(13) & Chr(10)


        Dim str2 As String

        str2 = Chr(2) & "2^^^989^\^^^990^\^^^991^\^^^763^\^^^57^|R||||||A||||1|||||||20170530180611|||O"
        str2 &= Chr(13)
        str2 &= "L|1|N"
        str2 &= Chr(13) & Chr(3)
        'Dim checksumvalue As String = GetCheckSumValue(str2.Remove(0, 1))
        Dim StrAll2 As String = StrAll.Remove(0, 1) & str2
        Dim checksumvalue As String = GetCheckSumValue(str2)
        str2 &= GetCheckSumValue(str2.Remove(0, 1))
        str2 &= Chr(13) & Chr(10)
        StrArr.Add(str2)


        returnArray = ReturnOrderDetail("^^             170015432^0^50042^042^^S0^") '26Test





    End Sub

    Public Function ReturnOrderDetail(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList
        Dim analyzerModel As String = "COBAS_C311"

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


            If analyzerModel = "DX300" Then
                SpecimenId = SpecimenId.Substring(0, SpecimenId.IndexOf(Chr(94))).Trim 'Golf 


            ElseIf analyzerModel = "SUZUKA" Then
                SpecimenId = SpecimenId.Substring(0, SpecimenId.IndexOf(Chr(94))).Trim 'Golf 
            ElseIf analyzerModel = "COBAS_C311" Then
                Return ReturnOrderDetail_COBAS_C311(SpecimenId)

            End If


            Return StrArr
        Catch ex As Exception

        End Try
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

        Dim dtTestResult As DataTable
        Dim dtTestResultToUpdate As DataTable



        Dim SpecimenIdString As String = SpecimenId
        Dim SpecimenIdSend As String = ""
        If SpecimenId.Length > 0 Then

            SpecimenIdArr = SpecimenIdString.Split("^")
            SpecimenIdSend = SpecimenIdArr(2)
            SpecimenId = SpecimenIdArr(2).Trim
            labSeq = SpecimenIdArr(3)
            RackIDNo = SpecimenIdArr(4)
            PositionNo = SpecimenIdArr(5)
            SampleType = SpecimenIdArr(7)
            ContainerType = SpecimenIdArr(8)
        End If


        Dim a As New AnalyzerInterface("", 19, "COBAS_C311")

        dtTestResult = a.GetResultCode(SpecimenId, "COBAS_C311", "19", "2017-05-12", "1", "Y").Copy
        Dim dvTestResult = New DataView(dtTestResult)
        dtTestResultToUpdate = dtTestResult.Clone()

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
        Dim Priority As String = "R"
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

        Str &= Chr(13) & "L|1|N"
        Dim Str1 As String
        Dim Str2 As String
        If Str.Length > 230 Then
            Str1 = Str.Substring(0, 230)
            Str2 = Str.Substring(230, Str.Length - Str1.Length)
            Str1 &= Chr(23) '23	17	00010111	ETB	end of trans. block
            Str1 &= GetCheckSumValue(Str1.Remove(0, 1))
            Str1 &= Chr(13) & Chr(10)
            StrArr.Add(Str1)

            If Str2.Length <= 230 Then
                Str2 = Chr(2) & Str2
                Str2 &= Chr(13) & Chr(3)
                Str2 &= GetCheckSumValue(Str2.Remove(0, 1))
                Str2 &= Chr(13) & Chr(10)
                StrArr.Add(Str2)
            End If
        Else
            Str &= Chr(13) & Chr(3)
            Str &= GetCheckSumValue(Str.Remove(0, 1))
            Str &= Chr(13) & Chr(10)
            StrArr.Add(Str)
        End If



        Return StrArr
    End Function

    Private Sub lvDevice_Click(sender As Object, e As EventArgs) Handles lvDevice.Click

        For Each item As ListViewItem In lvDevice.SelectedItems
            Dim Analyzer_Code As String = item.SubItems(2).Text
        Next
    End Sub



    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Dim returnArray As New ArrayList
        Dim REQUEST As String
        Select Case REQUEST
            Case "0"
                Dim strLocal As String
                strLocal = Chr(2) & "W" & Chr(28)
                strLocal &= GetCheckSumValue(strLocal.Remove(0, 1))
                strLocal &= Chr(3)
                returnArray.Add(strLocal)
            Case Else
                'returnArray = ReturnOrderDetailDimension("170001264") 'GLU,'BUN
                returnArray = ReturnOrderDetailDimension("170024377")

        End Select
    End Sub

    Public Function ReturnOrderDetailDimension(ByVal SpecimenId As String, Optional ByRef Running As Integer = 1) As ArrayList

        Dim StrArr As New ArrayList
        Dim Str As String
        Dim MultiTest As String = Chr(157)
        Dim SepTest As String = "^^^"
        Dim EndStr As String = Chr(13)
        Dim i As Integer
        Dim dtPatient As New DataTable
        Dim SpecimenIdReturn As String = SpecimenId 'Golf 
        Dim analyzerModel As String = "Dimension"
        Dim dtTestResult As New DataTable
        Dim dtTestResultToUpdate As New DataTable
        Dim analyzerSkey As String = "20"
        Dim analyzerDate As DateTime = Now

        Try
            dtPatient = GetPatientInformation(SpecimenId).Copy
            dtTestResult = GetResultCode(SpecimenId, analyzerModel, analyzerSkey, analyzerDate, "1").Copy
            Dim dvTestResult = New DataView(dtTestResult)
            dtTestResultToUpdate = dtTestResult.Clone()



            Select Case analyzerModel
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
                        Dim AGAP_Flag As Boolean = False

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
                            If row2.Item("analyzer_ref_cd") = "AGAP" Then
                                AGAP_Flag = True
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


            End Select


        Catch ex As ValueSoft.CoreControlsLib.CustomException

        Catch ex As Exception

        End Try

        Return StrArr

    End Function

    Function GetResultCode(ByVal SpecimenId As String, ByVal analyzer As String, ByVal analyzerSkey As Int32, ByVal analyzerDate As DateTime, ByVal flagExist As String) As DataTable
        Dim parm As IDbDataParameter()
        Dim SqlProvider As ValueSoft.DALManage.SqlDataProvider
        Dim conString As String = MyGlobal.myConnectionString
        SqlProvider = New ValueSoft.DALManage.SqlDataProvider(New SqlConnection(conString).ConnectionString)
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
        Catch ex As ValueSoft.CoreControlsLib.CustomException

        Catch ex As Exception

        Finally

            SqlProvider.Dispose()
        End Try


        Return dt

    End Function

    Public Function GetPatientInformation(ByVal SpecimenId As String) As DataTable
        Dim parm As IDbDataParameter()
        Dim conString As String = MyGlobal.myConnectionString
        Dim patientSqlProvider As ValueSoft.DALManage.SqlDataProvider = New ValueSoft.DALManage.SqlDataProvider(New SqlConnection(conString).ConnectionString)
        Dim dt As New DataTable

        Try
            parm = patientSqlProvider.GetParameterArray(1)
            parm(0) = patientSqlProvider.GetParameter("lab_no", DbType.String, SpecimenId, ParameterDirection.Input)
            patientSqlProvider.ExecuteDataTableSP(dt, "sp_lis_interface_get_patient_name", parm)
        Catch ex As ValueSoft.CoreControlsLib.CustomException

        Catch ex As Exception

        Finally
            patientSqlProvider.Dispose()
        End Try

        Return dt
    End Function

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As BackgroundWorker = CType(sender, BackgroundWorker)

        Dim i As Integer 
        For i = 1 To 10 
            If (worker.CancellationPending = True) Then 
                e.Cancel = True 
                Exit For 
            Else 
                System.Threading.Thread.Sleep(500) 
                worker.ReportProgress(i * 10) 
            End If 
        Next
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Me.lblResult.Text = (e.ProgressPercentage.ToString() + "%") 
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        If e.Cancelled = True Then 
            Me.lblResult.Text = "Canceled!" 
        ElseIf e.Error IsNot Nothing Then 
            Me.lblResult.Text = "Error: " & e.Error.Message 
        Else 
            Me.lblResult.Text = "Done!" 
        End If

        Me.btnStart.Enabled = True

        ' MessageBox.Show("Working Finished.")
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Try
                '  AppSetting.WriteErrorLog(COMPort, "DX300", String.Join("", returnArray.ToArray()))
                AppSetting.WriteErrorLog("", "DX300", "TESTGolf")
                Dim file1 As FileInfo = New FileInfo(strFileName)

                Dim script As String = file1.OpenText().ReadToEnd()

                ' Dim a As New AnalyzerInterface("", 29, "H-500")
                'Dim a As New AnalyzerInterface("", 17, "H-500")
                Dim a As New AnalyzerInterface("", 16, "H-500")
                a.ExtractResult_Micro(script)
            Catch ex As Exception
                AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
            End Try
        End If




    End Sub
    Private Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Dim dt As DataTable
        Dim analyzer_skey As Integer
        If e.Control And e.Alt And e.Shift And e.KeyCode = Keys.F1 Then

            Dim FrmShowKeyTest As New FrmShowKeyTest
            FrmShowKeyTest.ShowDialog()


        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.H Then
            'H-500
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try
                    AppSetting.WriteErrorLog("", "DX300", "TESTGolf")
                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 16, "H-500")
                    dt = b.GetANALYZER_skey("H-500")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "H-500")
                    a.ExtractResult_Micro(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.M Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 19, "RXMODENA")
                    dt = b.GetANALYZER_skey("RXMODENA")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "RXMODENA")
                    a.ExtractResult_Chem(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.A Then
            'AS720
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try 
                    Dim file1 As FileInfo = New FileInfo(strFileName) 
                    Dim script As String = file1.OpenText().ReadToEnd()

                    Dim b As New AnalyzerInterface("", 6, "AS720")
                    dt = b.GetANALYZER_skey("AS720")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim a As New AnalyzerInterface("", analyzer_skey, "AS720")
                    a.ExtractResult_Micro(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test AS720", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.R Then
            'AS720
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try
                    Dim file1 As FileInfo = New FileInfo(strFileName)
                    Dim script As String = file1.OpenText().ReadToEnd()

                    Dim b As New AnalyzerInterface("", 6, "RxLyte")
                    dt = b.GetANALYZER_skey("RxLyte")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim a As New AnalyzerInterface("", analyzer_skey, "RxLyte")
                    a.ExtractResult_RxLyte(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test AS720", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If

            
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.I Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "Imola")
                    dt = b.GetANALYZER_skey("Imola")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "Imola")
                    a.ExtractResult_Chem(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.L Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "Liaison")
                    dt = b.GetANALYZER_skey("Liaison")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "Liaison")
                    a.ExtractResult_Chem(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.X Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "XS")
                    dt = b.GetANALYZER_skey("XS")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "XS")
                    a.ExtractResult_Chem(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "XS", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If

        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.C Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "Mythix22")
                    dt = b.GetANALYZER_skey("Mythix22")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "Mythix22")
                    a.ExtractResult(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Mythix22", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.U Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "U120")
                    dt = b.GetANALYZER_skey("H-500")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "U120")
                    a.ExtractResult(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "U120", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.F2 Then
            'Cal
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)
                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim LineArr() As String
                    Dim j As Integer
                    LineArr = script.Split(Environment.NewLine)
                    j = LineArr.Count
                    If j <= 0 Then
                        Return
                    End If
                    Dim a As New AnalyzerInterface("", 16, "H-500")
                    Dim Specimen_type_id As String

                    For i = 0 To j - 1
                        Application.DoEvents()
                        LineArr(i) = LineArr(i).Replace(Chr(13), String.Empty)
                        If IsNumeric(LineArr(i)) Then
                            Specimen_type_id = LineArr(i).Trim
                            WriteLog("calculator formula", "cal", " Calculating (" & String.Format("Specimen_type_id : {0}  ", Specimen_type_id) & ")") 'Golf Mark 
                            a.UpdateFormula(Specimen_type_id)
                        End If

                    Next

                    Dim StartPath As String
                    StartPath = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)

                    If Not Directory.Exists(StartPath & "\Logs\cal") Then
                        Directory.CreateDirectory(StartPath & "\Logs\cal")
                    End If


                    If File.Exists(strFileName) Then

                        Dim FileNew As String

                        FileNew = StartPath & "\Logs\cal\" & file1.Name

                        file1.CopyTo(FileNew)
                    End If

                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.T Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "Architecti1000")
                    dt = b.GetANALYZER_skey("Architecti1000")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "Architecti1000")
                    a.ExtractResult_Architecti1000(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.P Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "C4000")
                    dt = b.GetANALYZER_skey("C4000")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "C4000")
                    a.ExtractResult_C4000(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test C4000", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        ElseIf e.Control And e.Alt And e.Shift And e.KeyCode = Keys.F Then
            'Modena
            Dim fd As OpenFileDialog = New OpenFileDialog()
            Dim strFileName As String

            fd.Title = "Open File Dialog"
            fd.InitialDirectory = "C:\"
            fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
            fd.FilterIndex = 2
            fd.RestoreDirectory = True

            If fd.ShowDialog() = DialogResult.OK Then
                strFileName = fd.FileName
                Try

                    Dim file1 As FileInfo = New FileInfo(strFileName)

                    Dim b As New AnalyzerInterface("", 21, "PATHFAST")
                    dt = b.GetANALYZER_skey("PATHFAST")
                    If dt.Rows.Count <= 0 Then
                        Return
                    End If
                    analyzer_skey = dt.Rows(0).Item("analyzer_skey").ToString()

                    Dim script As String = file1.OpenText().ReadToEnd()
                    Dim a As New AnalyzerInterface("", analyzer_skey, "PATHFAST")
                    a.ExtractResult_PATHFAST(script)
                Catch ex As Exception
                    AppSetting.WriteErrorLog("", "Test PATHFAST", ex.Message)
                    log.Error(ex.Message)
                End Try
            End If
        End If


    End Sub


    

    Private Sub lbMessageService_KeyDown(sender As Object, e As KeyEventArgs) Handles lbMessageService.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.C Then
            Dim copy_buffer As New System.Text.StringBuilder
            For Each item As Object In lbMessageService.SelectedItems
                copy_buffer.AppendLine(item.ToString)
            Next
            If copy_buffer.Length > 0 Then
                Clipboard.SetText(copy_buffer.ToString)
            End If
        End If
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Dim mon As ComPortMonitor
        For g As Integer = arMon.Count - 1 To 0 Step -1 'Golf mark 
            mon = arMon.Item(g)
            If Not (mon.checkComport()) Then
                'WriteLog(COMPort.PortName, "Error", "ComportMonitor.Error StopMonitorStopAndStart(): Disconnect") 'Golf Mark
                ' Start()
                StopAndStart()
            Else
                'WriteLog(COMPort.PortName, "Error", "ComportMonitor.Error StopMonitorStopAndStart(): connect") 'Golf Mark
                Timer1.Stop()
            End If
        Next

        If arMon.Count <= 0 Then
            'WriteLog(COMPort.PortName, "Error", "ComportMonitor.Error StopMonitorStopAndStart(): Disconnect") 'Golf Mark
            '   Start()
            StopAndStart()
        End If

    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        Try
            Dim J As Integer = 5
            Dim watch As Stopwatch = Stopwatch.StartNew()
            'For i As Integer = 1 To 100000000
            '    ' Threading.Thread.Sleep(1)
            '    If watch.Elapsed.TotalMilliseconds > 5000 Then
            '        Exit For
            '        'Exit do

            '    End If
            'Next


            Do
                J = J + 1
                If watch.Elapsed.TotalMilliseconds > 5000 Then
                    Exit Do


                End If
            Loop Until (J = 100000000)


            watch.Stop()
            Dim ssss As Integer = watch.Elapsed.TotalMilliseconds
            AppSetting.WriteErrorLog("", "Receive", "ComPortMonitor.Receiver  out Loop: " & "DFLKSJLFJSDF" & " Time : " & watch.Elapsed.TotalMilliseconds & " Milliseconds")
            Console.WriteLine(watch.Elapsed.TotalMilliseconds)

            'For i As Integer = 1 To J
            '    i = i - 1
            'Next 

            log.Info("Info Message")
            log.Error("Error Message")
            log.Warn("Warning Message")
        Catch ex As Exception
            AppSetting.WriteErrorLog("", "Test ISE6000", ex.Message)
        End Try

    End Sub
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Dim J As Integer = 5
        For i As Integer = 1 To J
            i = i - 1
        Next
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
       
        Dim ex As String = "Exception of type 'System.OutOfMemoryException' was thrown."

        If ex.IndexOf("OutOfMemoryException") <> -1 Then
            AppSetting.WriteErrorCal("Com1", "", "170003212")
        End If
         
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        log.Info("Info Message") 
        log.Error("Error Message")
        log.Warn("Warning Message")
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        ReceiverMaglumi()
    End Sub

    Sub ReceiverMaglumi()
         
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim ResultStr As String
        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
            Try 
                Dim file1 As FileInfo = New FileInfo(strFileName)

                Dim script As String = file1.OpenText().ReadToEnd()
                ResultStr = script
                'Dim a As New AnalyzerInterface("", 16, "H-500")
                'a.ExtractResult_Micro(script)
            Catch ex As Exception
                AppSetting.WriteErrorLog("", "Test H-500-001", ex.Message)
            End Try
        End If

        If ResultStr.IndexOf(Chr(5)) <> -1 Then
            Dim LineArr() As String
            Dim ItemStr As String
            Dim ItemArr() As String
            Dim InterfaceStatus As String = ""
            Dim LabArr As New ArrayList
            Dim LabNo As String = ""
            Dim j As Integer

            LineArr = ResultStr.Split(Chr(10)) 'Chr(10)
            j = LineArr.Count

            For i As Integer = 0 To j - 1
                ItemStr = LineArr(i)
                ItemArr = ItemStr.Split("|")

                If ItemArr.Count <= 0 Then
                    Continue For
                End If

                If ItemArr(0) = "" Then
                    Continue For
                End If

                'If ItemArr(0).Length < 3 Then
                '    Continue For
                'End If

                Select Case ItemArr(0)
                    Case "H"
                        'Head Line 

                    Case "Q"
                        If ItemArr.Length >= 3 Then
                            If ItemArr(2).Length >= 3 Then
                                InterfaceStatus = "Question"

                                LabNo = ItemArr(2)

                                LabArr.Add(LabNo)
                            End If
                        End If
                        'Exit For
                    Case "L"
                        'Last Line
                    Case "P"
                        'Result State
                        InterfaceStatus = "Result"
                        Exit For
                    Case "O"
                        'LabNo Line
                        InterfaceStatus = "Result"
                        If ItemArr.Length >= 3 Then
                            If ItemArr(2).Length >= 3 Then 'If ItemArr(2).Length >= 9 Then --Golf 2017-03-06
                                LabNo = ItemArr(2)
                            End If
                        End If
                    Case "C"
                        'Lab Detail Line
                End Select
            Next
        End If



    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        Dim AnalyzerInterface As New DeviceControlApp.AnalyzerInterface("", 0, "")
        AnalyzerInterface.TESTCheckSum()
    End Sub

    Private Sub frmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            Me.Hide()
            NotifyIcon1.Text = "iSLIM ANALYZERs"

            'NotifyIcon1.Visible = True
            'NotifyIcon1.Icon = SystemIcons.Application
            'NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            'NotifyIcon1.BalloonTipTitle = "iSLIM ANALYZERs"
            'NotifyIcon1.BalloonTipText = "iSLIM ANALYZERs"
            'NotifyIcon1.ShowBalloonTip(50000)
            'ShowInTaskbar = False
        End If
    End Sub 

    Private Sub DisplayMainScreenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DisplayMainScreenToolStripMenuItem.Click
        Me.Show()
        Me.WindowState = FormWindowState.Maximized
        NotifyIcon1.Visible = False
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub NotifyIcon1_DoubleClick1(sender As Object, e As EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Maximized
        NotifyIcon1.Visible = False
    End Sub
End Class