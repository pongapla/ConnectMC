<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddDevice
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAddDevice))
        Me.cbAnalyzerCd = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnRemove = New System.Windows.Forms.Button()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.lbAnalyzerModel = New System.Windows.Forms.Label()
        Me.lbSerialNo = New System.Windows.Forms.Label()
        Me.cbComPort = New System.Windows.Forms.ComboBox()
        Me.lbAnalyzerSkey = New System.Windows.Forms.Label()
        Me.txtAnalyzerDesc = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cbAnalyzerCd
        '
        Me.cbAnalyzerCd.FormattingEnabled = True
        Me.cbAnalyzerCd.Location = New System.Drawing.Point(108, 23)
        Me.cbAnalyzerCd.Name = "cbAnalyzerCd"
        Me.cbAnalyzerCd.Size = New System.Drawing.Size(134, 24)
        Me.cbAnalyzerCd.TabIndex = 8
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, 26)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 16)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Analyzer Code"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(248, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 16)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Com Port"
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(219, 100)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(91, 28)
        Me.btnRemove.TabIndex = 10
        Me.btnRemove.Text = "Cancel"
        Me.btnRemove.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(118, 100)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(91, 28)
        Me.btnAdd.TabIndex = 9
        Me.btnAdd.Text = "Add"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'lbAnalyzerModel
        '
        Me.lbAnalyzerModel.AutoSize = True
        Me.lbAnalyzerModel.Location = New System.Drawing.Point(92, 84)
        Me.lbAnalyzerModel.Name = "lbAnalyzerModel"
        Me.lbAnalyzerModel.Size = New System.Drawing.Size(101, 16)
        Me.lbAnalyzerModel.TabIndex = 11
        Me.lbAnalyzerModel.Text = "lbAnalyzerModel"
        Me.lbAnalyzerModel.Visible = False
        '
        'lbSerialNo
        '
        Me.lbSerialNo.AutoSize = True
        Me.lbSerialNo.Location = New System.Drawing.Point(196, 84)
        Me.lbSerialNo.Name = "lbSerialNo"
        Me.lbSerialNo.Size = New System.Drawing.Size(66, 16)
        Me.lbSerialNo.TabIndex = 12
        Me.lbSerialNo.Text = "lbSerialNo"
        Me.lbSerialNo.Visible = False
        '
        'cbComPort
        '
        Me.cbComPort.Enabled = False
        Me.cbComPort.FormattingEnabled = True
        Me.cbComPort.Location = New System.Drawing.Point(315, 23)
        Me.cbComPort.Name = "cbComPort"
        Me.cbComPort.Size = New System.Drawing.Size(80, 24)
        Me.cbComPort.TabIndex = 6
        '
        'lbAnalyzerSkey
        '
        Me.lbAnalyzerSkey.AutoSize = True
        Me.lbAnalyzerSkey.Location = New System.Drawing.Point(264, 84)
        Me.lbAnalyzerSkey.Name = "lbAnalyzerSkey"
        Me.lbAnalyzerSkey.Size = New System.Drawing.Size(94, 16)
        Me.lbAnalyzerSkey.TabIndex = 13
        Me.lbAnalyzerSkey.Text = "lbAnalyzerSkey"
        Me.lbAnalyzerSkey.Visible = False
        '
        'txtAnalyzerDesc
        '
        Me.txtAnalyzerDesc.BackColor = System.Drawing.SystemColors.Window
        Me.txtAnalyzerDesc.Location = New System.Drawing.Point(108, 57)
        Me.txtAnalyzerDesc.Multiline = True
        Me.txtAnalyzerDesc.Name = "txtAnalyzerDesc"
        Me.txtAnalyzerDesc.ReadOnly = True
        Me.txtAnalyzerDesc.Size = New System.Drawing.Size(287, 24)
        Me.txtAnalyzerDesc.TabIndex = 14
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 60)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "Analyzer Desc"
        '
        'frmAddDevice
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(411, 144)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtAnalyzerDesc)
        Me.Controls.Add(Me.lbAnalyzerSkey)
        Me.Controls.Add(Me.lbSerialNo)
        Me.Controls.Add(Me.lbAnalyzerModel)
        Me.Controls.Add(Me.btnRemove)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.cbAnalyzerCd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cbComPort)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "frmAddDevice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Device"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cbAnalyzerCd As System.Windows.Forms.ComboBox
    'Friend WithEvents dsMaster As DBDataset.MasterData
    'Friend WithEvents taAnalyzer As DBDataset.MasterDataTableAdapters.AnalyzerTableAdapter
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents lbAnalyzerModel As System.Windows.Forms.Label
    Friend WithEvents lbSerialNo As System.Windows.Forms.Label
    Friend WithEvents cbComPort As System.Windows.Forms.ComboBox
    Friend WithEvents lbAnalyzerSkey As System.Windows.Forms.Label
    Friend WithEvents txtAnalyzerDesc As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
End Class
