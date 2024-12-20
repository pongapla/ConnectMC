<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLogin
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.txtUserName = New System.Windows.Forms.TextBox()
        Me.txtDbName = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtServerName = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DevLookUpEditProfile = New DevExpress.XtraEditors.LookUpEdit()
        Me.btnDeleteProfile = New System.Windows.Forms.Button()
        Me.gbProfile = New System.Windows.Forms.GroupBox()
        Me.btnEditProfile = New System.Windows.Forms.Button()
        Me.btnAddProfile = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gbProfileDetail = New System.Windows.Forms.GroupBox()
        Me.btnOk = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        CType(Me.DevLookUpEditProfile.Properties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbProfile.SuspendLayout()
        Me.gbProfileDetail.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(166, 140)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(148, 20)
        Me.txtPassword.TabIndex = 17
        Me.txtPassword.UseSystemPasswordChar = True
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(166, 107)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(148, 20)
        Me.txtUserName.TabIndex = 16
        '
        'txtDbName
        '
        Me.txtDbName.Location = New System.Drawing.Point(166, 72)
        Me.txtDbName.Name = "txtDbName"
        Me.txtDbName.Size = New System.Drawing.Size(148, 20)
        Me.txtDbName.TabIndex = 15
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(14, 29)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(41, 13)
        Me.Label5.TabIndex = 18
        Me.Label5.Text = "Profiles"
        '
        'txtServerName
        '
        Me.txtServerName.Location = New System.Drawing.Point(166, 39)
        Me.txtServerName.Name = "txtServerName"
        Me.txtServerName.Size = New System.Drawing.Size(148, 20)
        Me.txtServerName.TabIndex = 14
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(62, 143)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(53, 13)
        Me.Label4.TabIndex = 13
        Me.Label4.Text = "Password"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(62, 110)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(60, 13)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "User Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(62, 75)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(84, 13)
        Me.Label2.TabIndex = 11
        Me.Label2.Text = "Database Name"
        '
        'DevLookUpEditProfile
        '
        Me.DevLookUpEditProfile.Location = New System.Drawing.Point(59, 28)
        Me.DevLookUpEditProfile.Name = "DevLookUpEditProfile"
        Me.DevLookUpEditProfile.Properties.Buttons.AddRange(New DevExpress.XtraEditors.Controls.EditorButton() {New DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)})
        Me.DevLookUpEditProfile.Properties.Columns.AddRange(New DevExpress.XtraEditors.Controls.LookUpColumnInfo() {New DevExpress.XtraEditors.Controls.LookUpColumnInfo("ProfileName", 120, "Profile Name"), New DevExpress.XtraEditors.Controls.LookUpColumnInfo("ServerName", 160, "Server Name"), New DevExpress.XtraEditors.Controls.LookUpColumnInfo("DatabaseName", 100, "Database Name"), New DevExpress.XtraEditors.Controls.LookUpColumnInfo("UserName", 120, "User Name")})
        Me.DevLookUpEditProfile.Size = New System.Drawing.Size(145, 20)
        Me.DevLookUpEditProfile.TabIndex = 22
        '
        'btnDeleteProfile
        '
        Me.btnDeleteProfile.Location = New System.Drawing.Point(314, 26)
        Me.btnDeleteProfile.Name = "btnDeleteProfile"
        Me.btnDeleteProfile.Size = New System.Drawing.Size(58, 23)
        Me.btnDeleteProfile.TabIndex = 21
        Me.btnDeleteProfile.Text = "Delete"
        Me.btnDeleteProfile.UseVisualStyleBackColor = True
        '
        'gbProfile
        '
        Me.gbProfile.Controls.Add(Me.DevLookUpEditProfile)
        Me.gbProfile.Controls.Add(Me.btnDeleteProfile)
        Me.gbProfile.Controls.Add(Me.btnEditProfile)
        Me.gbProfile.Controls.Add(Me.btnAddProfile)
        Me.gbProfile.Controls.Add(Me.Label5)
        Me.gbProfile.Location = New System.Drawing.Point(20, 8)
        Me.gbProfile.Name = "gbProfile"
        Me.gbProfile.Size = New System.Drawing.Size(382, 62)
        Me.gbProfile.TabIndex = 21
        Me.gbProfile.TabStop = False
        '
        'btnEditProfile
        '
        Me.btnEditProfile.Location = New System.Drawing.Point(262, 26)
        Me.btnEditProfile.Name = "btnEditProfile"
        Me.btnEditProfile.Size = New System.Drawing.Size(51, 23)
        Me.btnEditProfile.TabIndex = 20
        Me.btnEditProfile.Text = "Edit"
        Me.btnEditProfile.UseVisualStyleBackColor = True
        '
        'btnAddProfile
        '
        Me.btnAddProfile.Location = New System.Drawing.Point(210, 26)
        Me.btnAddProfile.Name = "btnAddProfile"
        Me.btnAddProfile.Size = New System.Drawing.Size(51, 23)
        Me.btnAddProfile.TabIndex = 19
        Me.btnAddProfile.Text = "Add"
        Me.btnAddProfile.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(62, 42)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(69, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Server Name"
        '
        'gbProfileDetail
        '
        Me.gbProfileDetail.Controls.Add(Me.txtPassword)
        Me.gbProfileDetail.Controls.Add(Me.txtUserName)
        Me.gbProfileDetail.Controls.Add(Me.txtDbName)
        Me.gbProfileDetail.Controls.Add(Me.txtServerName)
        Me.gbProfileDetail.Controls.Add(Me.Label4)
        Me.gbProfileDetail.Controls.Add(Me.Label3)
        Me.gbProfileDetail.Controls.Add(Me.Label2)
        Me.gbProfileDetail.Controls.Add(Me.Label1)
        Me.gbProfileDetail.Location = New System.Drawing.Point(19, 76)
        Me.gbProfileDetail.Name = "gbProfileDetail"
        Me.gbProfileDetail.Size = New System.Drawing.Size(383, 180)
        Me.gbProfileDetail.TabIndex = 20
        Me.gbProfileDetail.TabStop = False
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(84, 273)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(107, 35)
        Me.btnOk.TabIndex = 22
        Me.btnOk.Text = "OK"
        Me.btnOk.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(230, 273)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(107, 35)
        Me.btnCancel.TabIndex = 23
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(426, 320)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.gbProfile)
        Me.Controls.Add(Me.gbProfileDetail)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Login"
        CType(Me.DevLookUpEditProfile.Properties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbProfile.ResumeLayout(False)
        Me.gbProfile.PerformLayout()
        Me.gbProfileDetail.ResumeLayout(False)
        Me.gbProfileDetail.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents txtDbName As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtServerName As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents DevLookUpEditProfile As DevExpress.XtraEditors.LookUpEdit
    Friend WithEvents btnDeleteProfile As System.Windows.Forms.Button
    Friend WithEvents gbProfile As System.Windows.Forms.GroupBox
    Friend WithEvents btnEditProfile As System.Windows.Forms.Button
    Friend WithEvents btnAddProfile As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gbProfileDetail As System.Windows.Forms.GroupBox
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
End Class
