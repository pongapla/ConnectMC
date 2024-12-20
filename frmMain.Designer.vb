<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.lvDevice = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader6 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader7 = CType(New System.Windows.Forms.ColumnHeader(),System.Windows.Forms.ColumnHeader)
        Me.imgList = New System.Windows.Forms.ImageList(Me.components)
        Me.tmRS232 = New System.Windows.Forms.Timer(Me.components)
        Me.lbMessageService = New System.Windows.Forms.ListBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.btnOpenLog = New System.Windows.Forms.Button()
        Me.btnSetup = New System.Windows.Forms.Button()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.statusMain = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.btnCobas = New System.Windows.Forms.Button()
        Me.lvResult = New System.Windows.Forms.ListView()
        Me.ColumnHeader8 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader9 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader10 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader11 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader12 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader13 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader14 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader15 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader16 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.btnLd500 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.Button12 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Button15 = New System.Windows.Forms.Button()
        Me.Button16 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.Button18 = New System.Windows.Forms.Button()
        Me.Button19 = New System.Windows.Forms.Button()
        Me.Button20 = New System.Windows.Forms.Button()
        Me.Button21 = New System.Windows.Forms.Button()
        Me.Button22 = New System.Windows.Forms.Button()
        Me.Button23 = New System.Windows.Forms.Button()
        Me.Button24 = New System.Windows.Forms.Button()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.lblResult = New System.Windows.Forms.Label()
        Me.Button25 = New System.Windows.Forms.Button()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Button26 = New System.Windows.Forms.Button()
        Me.Button27 = New System.Windows.Forms.Button()
        Me.Button28 = New System.Windows.Forms.Button()
        Me.Button29 = New System.Windows.Forms.Button()
        Me.Button30 = New System.Windows.Forms.Button()
        Me.Button31 = New System.Windows.Forms.Button()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.DisplayMainScreenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.statusMain.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvDevice
        '
        Me.lvDevice.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lvDevice.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5, Me.ColumnHeader6, Me.ColumnHeader7})
        Me.lvDevice.FullRowSelect = True
        Me.lvDevice.GridLines = True
        Me.lvDevice.Location = New System.Drawing.Point(12, 82)
        Me.lvDevice.Name = "lvDevice"
        Me.lvDevice.Size = New System.Drawing.Size(330, 392)
        Me.lvDevice.SmallImageList = Me.imgList
        Me.lvDevice.TabIndex = 2
        Me.lvDevice.UseCompatibleStateImageBehavior = False
        Me.lvDevice.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "#"
        Me.ColumnHeader1.Width = 25
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Port"
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Analyzer Code"
        Me.ColumnHeader3.Width = 100
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Analyzer Description"
        Me.ColumnHeader4.Width = 130
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Analyzer Skey"
        Me.ColumnHeader5.Width = 0
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Serial No"
        Me.ColumnHeader6.Width = 0
        '
        'ColumnHeader7
        '
        Me.ColumnHeader7.Text = "Analyzer Model"
        Me.ColumnHeader7.Width = 0
        '
        'imgList
        '
        Me.imgList.ImageStream = CType(resources.GetObject("imgList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.imgList.TransparentColor = System.Drawing.Color.Transparent
        Me.imgList.Images.SetKeyName(0, "Red.png")
        Me.imgList.Images.SetKeyName(1, "yellow.png")
        Me.imgList.Images.SetKeyName(2, "green.png")
        '
        'tmRS232
        '
        Me.tmRS232.Interval = 500
        '
        'lbMessageService
        '
        Me.lbMessageService.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbMessageService.FormattingEnabled = True
        Me.lbMessageService.ItemHeight = 16
        Me.lbMessageService.Location = New System.Drawing.Point(12, 480)
        Me.lbMessageService.Name = "lbMessageService"
        Me.lbMessageService.ScrollAlwaysVisible = True
        Me.lbMessageService.Size = New System.Drawing.Size(1451, 84)
        Me.lbMessageService.TabIndex = 12
        '
        'btnClear
        '
        Me.btnClear.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClear.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnClear.Location = New System.Drawing.Point(1185, 586)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(135, 44)
        Me.btnClear.TabIndex = 15
        Me.btnClear.Text = "Clear Results"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Button2.Location = New System.Drawing.Point(1326, 586)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(135, 44)
        Me.Button2.TabIndex = 16
        Me.Button2.Text = "Clear Logs"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'btnOpenLog
        '
        Me.btnOpenLog.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnOpenLog.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnOpenLog.Image = Global.DeviceControlApp.My.Resources.Resources.green_folder_32
        Me.btnOpenLog.Location = New System.Drawing.Point(463, 586)
        Me.btnOpenLog.Name = "btnOpenLog"
        Me.btnOpenLog.Size = New System.Drawing.Size(110, 44)
        Me.btnOpenLog.TabIndex = 17
        Me.btnOpenLog.Text = "Open Log"
        Me.btnOpenLog.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnOpenLog.UseVisualStyleBackColor = True
        '
        'btnSetup
        '
        Me.btnSetup.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSetup.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnSetup.Image = Global.DeviceControlApp.My.Resources.Resources.green_setting_32
        Me.btnSetup.Location = New System.Drawing.Point(354, 586)
        Me.btnSetup.Name = "btnSetup"
        Me.btnSetup.Size = New System.Drawing.Size(103, 44)
        Me.btnSetup.TabIndex = 9
        Me.btnSetup.Text = "Setup"
        Me.btnSetup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnSetup.UseVisualStyleBackColor = True
        '
        'btnRemove
        '
        Me.btnRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnRemove.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnRemove.Image = Global.DeviceControlApp.My.Resources.Resources.Delete1
        Me.btnRemove.Location = New System.Drawing.Point(180, 586)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(162, 44)
        Me.btnRemove.TabIndex = 8
        Me.btnRemove.Text = "Remove Device"
        Me.btnRemove.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnAdd.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnAdd.Image = Global.DeviceControlApp.My.Resources.Resources.Add1
        Me.btnAdd.Location = New System.Drawing.Point(12, 586)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(162, 44)
        Me.btnAdd.TabIndex = 7
        Me.btnAdd.Text = "Add Device"
        Me.btnAdd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnStop.Image = Global.DeviceControlApp.My.Resources.Resources._Exit
        Me.btnStop.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.btnStop.Location = New System.Drawing.Point(180, 12)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(162, 64)
        Me.btnStop.TabIndex = 6
        Me.btnStop.Text = "Stop Monitor"
        Me.btnStop.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.btnStart.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnStart.Image = Global.DeviceControlApp.My.Resources.Resources.Loading
        Me.btnStart.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.btnStart.Location = New System.Drawing.Point(12, 12)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(162, 64)
        Me.btnStart.TabIndex = 5
        Me.btnStart.Text = "Start Monitor"
        Me.btnStart.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.Location = New System.Drawing.Point(838, 574)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(57, 25)
        Me.Button1.TabIndex = 18
        Me.Button1.Text = "Rxlyte"
        Me.Button1.UseVisualStyleBackColor = True
        Me.Button1.Visible = False
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Location = New System.Drawing.Point(910, 586)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(62, 25)
        Me.Button3.TabIndex = 19
        Me.Button3.Text = "AS720"
        Me.Button3.UseVisualStyleBackColor = True
        Me.Button3.Visible = False
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Location = New System.Drawing.Point(593, 605)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(93, 25)
        Me.Button4.TabIndex = 20
        Me.Button4.Text = "Test Architecti1000"
        Me.Button4.UseVisualStyleBackColor = True
        Me.Button4.Visible = False
        '
        'Button5
        '
        Me.Button5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button5.Location = New System.Drawing.Point(945, 610)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(93, 25)
        Me.Button5.TabIndex = 21
        Me.Button5.Text = "Test Uriscan"
        Me.Button5.UseVisualStyleBackColor = True
        Me.Button5.Visible = False
        '
        'statusMain
        '
        Me.statusMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel1})
        Me.statusMain.Location = New System.Drawing.Point(0, 637)
        Me.statusMain.Name = "statusMain"
        Me.statusMain.Size = New System.Drawing.Size(1475, 22)
        Me.statusMain.TabIndex = 22
        Me.statusMain.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel1
        '
        Me.ToolStripStatusLabel1.Name = "ToolStripStatusLabel1"
        Me.ToolStripStatusLabel1.Size = New System.Drawing.Size(0, 17)
        '
        'btnCobas
        '
        Me.btnCobas.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCobas.Location = New System.Drawing.Point(976, 586)
        Me.btnCobas.Name = "btnCobas"
        Me.btnCobas.Size = New System.Drawing.Size(62, 25)
        Me.btnCobas.TabIndex = 23
        Me.btnCobas.Text = "COBAS"
        Me.btnCobas.UseVisualStyleBackColor = True
        Me.btnCobas.Visible = False
        '
        'lvResult
        '
        Me.lvResult.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lvResult.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader8, Me.ColumnHeader9, Me.ColumnHeader10, Me.ColumnHeader11, Me.ColumnHeader12, Me.ColumnHeader13, Me.ColumnHeader14, Me.ColumnHeader15, Me.ColumnHeader16})
        Me.lvResult.FullRowSelect = True
        Me.lvResult.GridLines = True
        Me.lvResult.Location = New System.Drawing.Point(354, 12)
        Me.lvResult.Name = "lvResult"
        Me.lvResult.Size = New System.Drawing.Size(1107, 462)
        Me.lvResult.TabIndex = 24
        Me.lvResult.UseCompatibleStateImageBehavior = False
        Me.lvResult.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader8
        '
        Me.ColumnHeader8.Text = "#"
        Me.ColumnHeader8.Width = 5
        '
        'ColumnHeader9
        '
        Me.ColumnHeader9.Text = "Received Date"
        Me.ColumnHeader9.Width = 150
        '
        'ColumnHeader10
        '
        Me.ColumnHeader10.Text = "Analyzed Date"
        Me.ColumnHeader10.Width = 150
        '
        'ColumnHeader11
        '
        Me.ColumnHeader11.Text = "Lab No."
        Me.ColumnHeader11.Width = 120
        '
        'ColumnHeader12
        '
        Me.ColumnHeader12.Text = "Specimen ID"
        Me.ColumnHeader12.Width = 110
        '
        'ColumnHeader13
        '
        Me.ColumnHeader13.Text = "Analyzer Code"
        Me.ColumnHeader13.Width = 110
        '
        'ColumnHeader14
        '
        Me.ColumnHeader14.Text = "Ref Code"
        Me.ColumnHeader14.Width = 100
        '
        'ColumnHeader15
        '
        Me.ColumnHeader15.Text = "Result Alias"
        Me.ColumnHeader15.Width = 100
        '
        'ColumnHeader16
        '
        Me.ColumnHeader16.Text = "Result Value"
        Me.ColumnHeader16.Width = 120
        '
        'Button6
        '
        Me.Button6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button6.Location = New System.Drawing.Point(1044, 570)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(62, 25)
        Me.Button6.TabIndex = 25
        Me.Button6.Text = "Mythic"
        Me.Button6.UseVisualStyleBackColor = True
        Me.Button6.Visible = False
        '
        'Button7
        '
        Me.Button7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button7.Location = New System.Drawing.Point(1044, 610)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(62, 25)
        Me.Button7.TabIndex = 26
        Me.Button7.Text = "Daytona"
        Me.Button7.UseVisualStyleBackColor = True
        Me.Button7.Visible = False
        '
        'btnLd500
        '
        Me.btnLd500.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLd500.Location = New System.Drawing.Point(1112, 569)
        Me.btnLd500.Name = "btnLd500"
        Me.btnLd500.Size = New System.Drawing.Size(59, 26)
        Me.btnLd500.TabIndex = 27
        Me.btnLd500.Text = "LD-500"
        Me.btnLd500.UseVisualStyleBackColor = True
        Me.btnLd500.Visible = False
        '
        'Button8
        '
        Me.Button8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button8.Location = New System.Drawing.Point(813, 610)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(107, 25)
        Me.Button8.TabIndex = 28
        Me.Button8.Text = "Test H-500 _13"
        Me.Button8.UseVisualStyleBackColor = True
        Me.Button8.Visible = False
        '
        'Button9
        '
        Me.Button9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button9.Location = New System.Drawing.Point(692, 605)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(93, 25)
        Me.Button9.TabIndex = 29
        Me.Button9.Text = "Test dx300"
        Me.Button9.UseVisualStyleBackColor = True
        Me.Button9.Visible = False
        '
        'Button10
        '
        Me.Button10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button10.Location = New System.Drawing.Point(666, 514)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(128, 25)
        Me.Button10.TabIndex = 30
        Me.Button10.Text = "Test dx300 return"
        Me.Button10.UseVisualStyleBackColor = True
        Me.Button10.Visible = False
        '
        'Button11
        '
        Me.Button11.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button11.Location = New System.Drawing.Point(711, 569)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(107, 25)
        Me.Button11.TabIndex = 31
        Me.Button11.Text = "Test U120"
        Me.Button11.UseVisualStyleBackColor = True
        Me.Button11.Visible = False
        '
        'Button12
        '
        Me.Button12.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button12.Location = New System.Drawing.Point(782, 605)
        Me.Button12.Name = "Button12"
        Me.Button12.Size = New System.Drawing.Size(93, 25)
        Me.Button12.TabIndex = 32
        Me.Button12.Text = "Test SUZUKA"
        Me.Button12.UseVisualStyleBackColor = True
        Me.Button12.Visible = False
        '
        'Button13
        '
        Me.Button13.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button13.Location = New System.Drawing.Point(838, 499)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(151, 25)
        Me.Button13.TabIndex = 33
        Me.Button13.Text = "Test LIAISON to LIS"
        Me.Button13.UseVisualStyleBackColor = True
        Me.Button13.Visible = False
        '
        'Button14
        '
        Me.Button14.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button14.Location = New System.Drawing.Point(1185, 514)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(93, 25)
        Me.Button14.TabIndex = 34
        Me.Button14.Text = "Test Liaison"
        Me.Button14.UseVisualStyleBackColor = True
        Me.Button14.Visible = False
        '
        'Button15
        '
        Me.Button15.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button15.Location = New System.Drawing.Point(593, 574)
        Me.Button15.Name = "Button15"
        Me.Button15.Size = New System.Drawing.Size(107, 25)
        Me.Button15.TabIndex = 35
        Me.Button15.Text = "Test H-500 _11"
        Me.Button15.UseVisualStyleBackColor = True
        Me.Button15.Visible = False
        '
        'Button16
        '
        Me.Button16.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button16.Location = New System.Drawing.Point(804, 579)
        Me.Button16.Name = "Button16"
        Me.Button16.Size = New System.Drawing.Size(139, 25)
        Me.Button16.TabIndex = 36
        Me.Button16.Text = "Test Hang stopTime"
        Me.Button16.UseVisualStyleBackColor = True
        Me.Button16.Visible = False
        '
        'Button17
        '
        Me.Button17.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button17.Location = New System.Drawing.Point(1016, 530)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(77, 25)
        Me.Button17.TabIndex = 37
        Me.Button17.Text = "ISE6000"
        Me.Button17.UseVisualStyleBackColor = True
        Me.Button17.Visible = False
        '
        'Button18
        '
        Me.Button18.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button18.Location = New System.Drawing.Point(945, 555)
        Me.Button18.Name = "Button18"
        Me.Button18.Size = New System.Drawing.Size(107, 25)
        Me.Button18.TabIndex = 38
        Me.Button18.Text = "HumaClot Pro"
        Me.Button18.UseVisualStyleBackColor = True
        Me.Button18.Visible = False
        '
        'Button19
        '
        Me.Button19.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button19.Location = New System.Drawing.Point(1016, 499)
        Me.Button19.Name = "Button19"
        Me.Button19.Size = New System.Drawing.Size(125, 25)
        Me.Button19.TabIndex = 39
        Me.Button19.Text = "Test RXMODENA"
        Me.Button19.UseVisualStyleBackColor = True
        Me.Button19.Visible = False
        '
        'Button20
        '
        Me.Button20.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button20.Location = New System.Drawing.Point(417, 555)
        Me.Button20.Name = "Button20"
        Me.Button20.Size = New System.Drawing.Size(125, 25)
        Me.Button20.TabIndex = 40
        Me.Button20.Text = "Dimension"
        Me.Button20.UseVisualStyleBackColor = True
        Me.Button20.Visible = False
        '
        'Button21
        '
        Me.Button21.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button21.Location = New System.Drawing.Point(286, 555)
        Me.Button21.Name = "Button21"
        Me.Button21.Size = New System.Drawing.Size(125, 25)
        Me.Button21.TabIndex = 41
        Me.Button21.Text = "DimensionResult"
        Me.Button21.UseVisualStyleBackColor = True
        Me.Button21.Visible = False
        '
        'Button22
        '
        Me.Button22.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button22.Location = New System.Drawing.Point(23, 555)
        Me.Button22.Name = "Button22"
        Me.Button22.Size = New System.Drawing.Size(125, 25)
        Me.Button22.TabIndex = 43
        Me.Button22.Text = "COBAS_C311Result"
        Me.Button22.UseVisualStyleBackColor = True
        Me.Button22.Visible = False
        '
        'Button23
        '
        Me.Button23.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button23.Location = New System.Drawing.Point(154, 555)
        Me.Button23.Name = "Button23"
        Me.Button23.Size = New System.Drawing.Size(125, 25)
        Me.Button23.TabIndex = 42
        Me.Button23.Text = "COBAS_C311"
        Me.Button23.UseVisualStyleBackColor = True
        Me.Button23.Visible = False
        '
        'Button24
        '
        Me.Button24.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button24.Location = New System.Drawing.Point(548, 555)
        Me.Button24.Name = "Button24"
        Me.Button24.Size = New System.Drawing.Size(125, 25)
        Me.Button24.TabIndex = 44
        Me.Button24.Text = "Dimension Test Order"
        Me.Button24.UseVisualStyleBackColor = True
        Me.Button24.Visible = False
        '
        'BackgroundWorker1
        '
        '
        'lblResult
        '
        Me.lblResult.AutoSize = True
        Me.lblResult.Location = New System.Drawing.Point(1112, 614)
        Me.lblResult.Name = "lblResult"
        Me.lblResult.Size = New System.Drawing.Size(0, 16)
        Me.lblResult.TabIndex = 45
        '
        'Button25
        '
        Me.Button25.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button25.Location = New System.Drawing.Point(1305, 560)
        Me.Button25.Name = "Button25"
        Me.Button25.Size = New System.Drawing.Size(107, 25)
        Me.Button25.TabIndex = 46
        Me.Button25.Text = "H-500 browse"
        Me.Button25.UseVisualStyleBackColor = True
        Me.Button25.Visible = False
        '
        'Timer1
        '
        '
        'Button26
        '
        Me.Button26.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button26.Location = New System.Drawing.Point(945, 579)
        Me.Button26.Name = "Button26"
        Me.Button26.Size = New System.Drawing.Size(139, 25)
        Me.Button26.TabIndex = 47
        Me.Button26.Text = "Test Hang"
        Me.Button26.UseVisualStyleBackColor = True
        Me.Button26.Visible = False
        '
        'Button27
        '
        Me.Button27.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button27.Location = New System.Drawing.Point(1044, 601)
        Me.Button27.Name = "Button27"
        Me.Button27.Size = New System.Drawing.Size(139, 25)
        Me.Button27.TabIndex = 48
        Me.Button27.Text = "Cal"
        Me.Button27.UseVisualStyleBackColor = True
        Me.Button27.Visible = False
        '
        'Button28
        '
        Me.Button28.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button28.Location = New System.Drawing.Point(736, 538)
        Me.Button28.Name = "Button28"
        Me.Button28.Size = New System.Drawing.Size(139, 25)
        Me.Button28.TabIndex = 49
        Me.Button28.Text = "Test Log"
        Me.Button28.UseVisualStyleBackColor = True
        Me.Button28.Visible = False
        '
        'Button29
        '
        Me.Button29.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button29.Location = New System.Drawing.Point(625, 560)
        Me.Button29.Name = "Button29"
        Me.Button29.Size = New System.Drawing.Size(125, 25)
        Me.Button29.TabIndex = 50
        Me.Button29.Text = "Maglumi_Inquiry"
        Me.Button29.UseVisualStyleBackColor = True
        Me.Button29.Visible = False
        '
        'Button30
        '
        Me.Button30.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button30.Location = New System.Drawing.Point(1154, 560)
        Me.Button30.Name = "Button30"
        Me.Button30.Size = New System.Drawing.Size(93, 25)
        Me.Button30.TabIndex = 51
        Me.Button30.Text = "Test Liaison"
        Me.Button30.UseVisualStyleBackColor = True
        Me.Button30.Visible = False
        '
        'Button31
        '
        Me.Button31.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button31.Location = New System.Drawing.Point(463, 499)
        Me.Button31.Name = "Button31"
        Me.Button31.Size = New System.Drawing.Size(139, 25)
        Me.Button31.TabIndex = 52
        Me.Button31.Text = "Test Log"
        Me.Button31.UseVisualStyleBackColor = True
        Me.Button31.Visible = False
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DisplayMainScreenToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(245, 198)
        '
        'DisplayMainScreenToolStripMenuItem
        '
        Me.DisplayMainScreenToolStripMenuItem.Image = Global.DeviceControlApp.My.Resources.Resources.Loading
        Me.DisplayMainScreenToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.DisplayMainScreenToolStripMenuItem.Name = "DisplayMainScreenToolStripMenuItem"
        Me.DisplayMainScreenToolStripMenuItem.Size = New System.Drawing.Size(244, 86)
        Me.DisplayMainScreenToolStripMenuItem.Text = "Display Main Screen"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Image = CType(resources.GetObject("ExitToolStripMenuItem.Image"), System.Drawing.Image)
        Me.ExitToolStripMenuItem.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(244, 86)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "NotifyIcon1"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(1475, 659)
        Me.Controls.Add(Me.Button31)
        Me.Controls.Add(Me.Button30)
        Me.Controls.Add(Me.Button29)
        Me.Controls.Add(Me.Button28)
        Me.Controls.Add(Me.Button27)
        Me.Controls.Add(Me.Button26)
        Me.Controls.Add(Me.Button25)
        Me.Controls.Add(Me.lblResult)
        Me.Controls.Add(Me.Button24)
        Me.Controls.Add(Me.Button22)
        Me.Controls.Add(Me.Button23)
        Me.Controls.Add(Me.Button21)
        Me.Controls.Add(Me.Button20)
        Me.Controls.Add(Me.Button19)
        Me.Controls.Add(Me.Button18)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.Button16)
        Me.Controls.Add(Me.Button15)
        Me.Controls.Add(Me.Button14)
        Me.Controls.Add(Me.Button13)
        Me.Controls.Add(Me.Button12)
        Me.Controls.Add(Me.Button11)
        Me.Controls.Add(Me.Button10)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.btnLd500)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.lvResult)
        Me.Controls.Add(Me.btnCobas)
        Me.Controls.Add(Me.statusMain)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnOpenLog)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.lbMessageService)
        Me.Controls.Add(Me.btnSetup)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.lvDevice)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "iSLIM ANALYZERs"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.statusMain.ResumeLayout(False)
        Me.statusMain.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout

End Sub
    Friend WithEvents lvDevice As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents imgList As System.Windows.Forms.ImageList
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnSetup As System.Windows.Forms.Button
    Friend WithEvents tmRS232 As System.Windows.Forms.Timer
    Friend WithEvents lbMessageService As System.Windows.Forms.ListBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents btnOpenLog As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents statusMain As System.Windows.Forms.StatusStrip
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader7 As System.Windows.Forms.ColumnHeader
    Friend WithEvents btnCobas As System.Windows.Forms.Button
    Friend WithEvents lvResult As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader8 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader9 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader10 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader11 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader12 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader13 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader14 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader15 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader16 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents btnLd500 As System.Windows.Forms.Button
    Friend WithEvents Button8 As System.Windows.Forms.Button
    Friend WithEvents Button9 As System.Windows.Forms.Button
    Friend WithEvents Button10 As System.Windows.Forms.Button
    Friend WithEvents Button11 As System.Windows.Forms.Button
    Friend WithEvents Button12 As System.Windows.Forms.Button
    Friend WithEvents Button13 As System.Windows.Forms.Button
    Friend WithEvents Button14 As System.Windows.Forms.Button
    Friend WithEvents Button15 As System.Windows.Forms.Button
    Friend WithEvents Button16 As System.Windows.Forms.Button
    Friend WithEvents Button17 As System.Windows.Forms.Button
    Friend WithEvents Button18 As System.Windows.Forms.Button
    Friend WithEvents Button19 As System.Windows.Forms.Button
    Friend WithEvents Button20 As System.Windows.Forms.Button
    Friend WithEvents Button21 As System.Windows.Forms.Button
    Friend WithEvents Button22 As System.Windows.Forms.Button
    Friend WithEvents Button23 As System.Windows.Forms.Button
    Friend WithEvents Button24 As System.Windows.Forms.Button
    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents lblResult As System.Windows.Forms.Label
    Friend WithEvents Button25 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Button26 As System.Windows.Forms.Button
    Friend WithEvents Button27 As System.Windows.Forms.Button
    Friend WithEvents Button28 As System.Windows.Forms.Button
    Friend WithEvents ToolStripStatusLabel1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Button29 As System.Windows.Forms.Button
    Friend WithEvents Button30 As System.Windows.Forms.Button
    Friend WithEvents Button31 As System.Windows.Forms.Button
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents DisplayMainScreenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
End Class
