<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSetup
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSetup))
        Me.tabSetup = New System.Windows.Forms.TabPage()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtWaitConnect = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.chkWriteLog = New System.Windows.Forms.CheckBox()
        Me.chkAutoStart = New System.Windows.Forms.CheckBox()
        Me.tabHis = New System.Windows.Forms.TabPage()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.picLis = New System.Windows.Forms.PictureBox()
        Me.txtLisPassword = New System.Windows.Forms.TextBox()
        Me.txtLisUser = New System.Windows.Forms.TextBox()
        Me.txtLisDatabase = New System.Windows.Forms.TextBox()
        Me.txtLisServer = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.btOk = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.btApply = New System.Windows.Forms.Button()
        Me.chkTrackLog = New System.Windows.Forms.CheckBox()
        Me.tabSetup.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.tabHis.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.picLis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabControl1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabSetup
        '
        Me.tabSetup.Controls.Add(Me.GroupBox4)
        Me.tabSetup.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.tabSetup.Location = New System.Drawing.Point(4, 22)
        Me.tabSetup.Name = "tabSetup"
        Me.tabSetup.Padding = New System.Windows.Forms.Padding(3)
        Me.tabSetup.Size = New System.Drawing.Size(756, 366)
        Me.tabSetup.TabIndex = 2
        Me.tabSetup.Text = "SETUP"
        Me.tabSetup.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.chkTrackLog)
        Me.GroupBox4.Controls.Add(Me.txtWaitConnect)
        Me.GroupBox4.Controls.Add(Me.Label1)
        Me.GroupBox4.Controls.Add(Me.chkWriteLog)
        Me.GroupBox4.Controls.Add(Me.chkAutoStart)
        Me.GroupBox4.Location = New System.Drawing.Point(190, 33)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(382, 180)
        Me.GroupBox4.TabIndex = 10
        Me.GroupBox4.TabStop = False
        '
        'txtWaitConnect
        '
        Me.txtWaitConnect.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtWaitConnect.Location = New System.Drawing.Point(187, 107)
        Me.txtWaitConnect.Name = "txtWaitConnect"
        Me.txtWaitConnect.Size = New System.Drawing.Size(100, 23)
        Me.txtWaitConnect.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(26, 110)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(162, 17)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Wait Connect (Second) :"
        '
        'chkWriteLog
        '
        Me.chkWriteLog.AutoSize = True
        Me.chkWriteLog.Location = New System.Drawing.Point(29, 75)
        Me.chkWriteLog.Name = "chkWriteLog"
        Me.chkWriteLog.Size = New System.Drawing.Size(88, 21)
        Me.chkWriteLog.TabIndex = 9
        Me.chkWriteLog.Text = "Write Log"
        Me.chkWriteLog.UseVisualStyleBackColor = True
        '
        'chkAutoStart
        '
        Me.chkAutoStart.AutoSize = True
        Me.chkAutoStart.Location = New System.Drawing.Point(29, 33)
        Me.chkAutoStart.Name = "chkAutoStart"
        Me.chkAutoStart.Size = New System.Drawing.Size(184, 21)
        Me.chkAutoStart.TabIndex = 8
        Me.chkAutoStart.Text = "Run when Windows Start"
        Me.chkAutoStart.UseVisualStyleBackColor = True
        '
        'tabHis
        '
        Me.tabHis.Controls.Add(Me.GroupBox2)
        Me.tabHis.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.tabHis.Location = New System.Drawing.Point(4, 22)
        Me.tabHis.Name = "tabHis"
        Me.tabHis.Padding = New System.Windows.Forms.Padding(3)
        Me.tabHis.Size = New System.Drawing.Size(756, 366)
        Me.tabHis.TabIndex = 0
        Me.tabHis.Text = "CONNECTION"
        Me.tabHis.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.picLis)
        Me.GroupBox2.Controls.Add(Me.txtLisPassword)
        Me.GroupBox2.Controls.Add(Me.txtLisUser)
        Me.GroupBox2.Controls.Add(Me.txtLisDatabase)
        Me.GroupBox2.Controls.Add(Me.txtLisServer)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.Label4)
        Me.GroupBox2.Controls.Add(Me.Label9)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Location = New System.Drawing.Point(208, 21)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(326, 313)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "LIS CONNECTION"
        '
        'picLis
        '
        Me.picLis.Image = Global.DeviceControlApp.My.Resources.Resources.red_light32
        Me.picLis.Location = New System.Drawing.Point(157, 246)
        Me.picLis.Name = "picLis"
        Me.picLis.Size = New System.Drawing.Size(34, 33)
        Me.picLis.TabIndex = 9
        Me.picLis.TabStop = False
        '
        'txtLisPassword
        '
        Me.txtLisPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtLisPassword.Location = New System.Drawing.Point(121, 197)
        Me.txtLisPassword.Name = "txtLisPassword"
        Me.txtLisPassword.Size = New System.Drawing.Size(172, 23)
        Me.txtLisPassword.TabIndex = 7
        Me.txtLisPassword.UseSystemPasswordChar = True
        '
        'txtLisUser
        '
        Me.txtLisUser.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtLisUser.Location = New System.Drawing.Point(121, 144)
        Me.txtLisUser.Name = "txtLisUser"
        Me.txtLisUser.Size = New System.Drawing.Size(172, 23)
        Me.txtLisUser.TabIndex = 6
        '
        'txtLisDatabase
        '
        Me.txtLisDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtLisDatabase.Location = New System.Drawing.Point(121, 94)
        Me.txtLisDatabase.Name = "txtLisDatabase"
        Me.txtLisDatabase.Size = New System.Drawing.Size(172, 23)
        Me.txtLisDatabase.TabIndex = 5
        '
        'txtLisServer
        '
        Me.txtLisServer.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.txtLisServer.Location = New System.Drawing.Point(121, 39)
        Me.txtLisServer.Name = "txtLisServer"
        Me.txtLisServer.Size = New System.Drawing.Size(172, 23)
        Me.txtLisServer.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.Location = New System.Drawing.Point(34, 200)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(73, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Password:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.Location = New System.Drawing.Point(34, 147)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(42, 17)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "User:"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label9.Location = New System.Drawing.Point(34, 97)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(73, 17)
        Me.Label9.TabIndex = 1
        Me.Label9.Text = "Database:"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label10.Location = New System.Drawing.Point(34, 42)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(54, 17)
        Me.Label10.TabIndex = 0
        Me.Label10.Text = "Server:"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabHis)
        Me.TabControl1.Controls.Add(Me.tabSetup)
        Me.TabControl1.Location = New System.Drawing.Point(12, 12)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(764, 392)
        Me.TabControl1.TabIndex = 0
        '
        'btOk
        '
        Me.btOk.Location = New System.Drawing.Point(206, 410)
        Me.btOk.Name = "btOk"
        Me.btOk.Size = New System.Drawing.Size(105, 36)
        Me.btOk.TabIndex = 1
        Me.btOk.Text = "OK"
        Me.btOk.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Location = New System.Drawing.Point(334, 410)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(105, 36)
        Me.btCancel.TabIndex = 2
        Me.btCancel.Text = "Cancel"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'btApply
        '
        Me.btApply.Location = New System.Drawing.Point(464, 410)
        Me.btApply.Name = "btApply"
        Me.btApply.Size = New System.Drawing.Size(105, 36)
        Me.btApply.TabIndex = 3
        Me.btApply.Text = "Apply"
        Me.btApply.UseVisualStyleBackColor = True
        '
        'chkTrackLog
        '
        Me.chkTrackLog.AutoSize = True
        Me.chkTrackLog.Location = New System.Drawing.Point(29, 143)
        Me.chkTrackLog.Name = "chkTrackLog"
        Me.chkTrackLog.Size = New System.Drawing.Size(125, 21)
        Me.chkTrackLog.TabIndex = 12
        Me.chkTrackLog.Text = "Track event log"
        Me.chkTrackLog.UseVisualStyleBackColor = True
        '
        'frmSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(788, 458)
        Me.Controls.Add(Me.btApply)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOk)
        Me.Controls.Add(Me.TabControl1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSetup"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "iSLIM ANALYZERs SETUP"
        Me.tabSetup.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.tabHis.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.picLis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabControl1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tabSetup As System.Windows.Forms.TabPage
    Friend WithEvents tabHis As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtLisPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtLisUser As System.Windows.Forms.TextBox
    Friend WithEvents txtLisDatabase As System.Windows.Forms.TextBox
    Friend WithEvents txtLisServer As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents btOk As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents btApply As System.Windows.Forms.Button
    Friend WithEvents picLis As System.Windows.Forms.PictureBox
    Friend WithEvents chkAutoStart As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents chkWriteLog As System.Windows.Forms.CheckBox
    Friend WithEvents txtWaitConnect As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents chkTrackLog As System.Windows.Forms.CheckBox
End Class
