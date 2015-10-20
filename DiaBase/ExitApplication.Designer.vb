<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExitApplication
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ExitApplication))
        Me.ExitApplicationHeaderLABEL = New System.Windows.Forms.Label()
        Me.ExitApplicationInformationLABEL = New System.Windows.Forms.Label()
        Me.ExitApplicationSaveDatabaseCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.ExitApplicationBackupDatabaseCHRCKBOX = New System.Windows.Forms.CheckBox()
        Me.ExitApplicationConfirmBUTTON = New System.Windows.Forms.Button()
        Me.ExitApplicationCancelBUTTON = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ExitApplicationHeaderLABEL
        '
        Me.ExitApplicationHeaderLABEL.BackColor = System.Drawing.Color.Black
        Me.ExitApplicationHeaderLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExitApplicationHeaderLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExitApplicationHeaderLABEL.Location = New System.Drawing.Point(37, 32)
        Me.ExitApplicationHeaderLABEL.Name = "ExitApplicationHeaderLABEL"
        Me.ExitApplicationHeaderLABEL.Size = New System.Drawing.Size(325, 23)
        Me.ExitApplicationHeaderLABEL.TabIndex = 11
        Me.ExitApplicationHeaderLABEL.Text = "SAVE AND BACKUP BEFORE EXITING"
        Me.ExitApplicationHeaderLABEL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ExitApplicationInformationLABEL
        '
        Me.ExitApplicationInformationLABEL.BackColor = System.Drawing.Color.Black
        Me.ExitApplicationInformationLABEL.ForeColor = System.Drawing.Color.White
        Me.ExitApplicationInformationLABEL.Location = New System.Drawing.Point(37, 66)
        Me.ExitApplicationInformationLABEL.Name = "ExitApplicationInformationLABEL"
        Me.ExitApplicationInformationLABEL.Size = New System.Drawing.Size(325, 73)
        Me.ExitApplicationInformationLABEL.TabIndex = 10
        Me.ExitApplicationInformationLABEL.Text = resources.GetString("ExitApplicationInformationLABEL.Text")
        '
        'ExitApplicationSaveDatabaseCHECKBOX
        '
        Me.ExitApplicationSaveDatabaseCHECKBOX.AutoSize = True
        Me.ExitApplicationSaveDatabaseCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.ExitApplicationSaveDatabaseCHECKBOX.Checked = True
        Me.ExitApplicationSaveDatabaseCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ExitApplicationSaveDatabaseCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExitApplicationSaveDatabaseCHECKBOX.Location = New System.Drawing.Point(40, 150)
        Me.ExitApplicationSaveDatabaseCHECKBOX.Name = "ExitApplicationSaveDatabaseCHECKBOX"
        Me.ExitApplicationSaveDatabaseCHECKBOX.Size = New System.Drawing.Size(100, 17)
        Me.ExitApplicationSaveDatabaseCHECKBOX.TabIndex = 7
        Me.ExitApplicationSaveDatabaseCHECKBOX.Text = "Save Database"
        Me.ExitApplicationSaveDatabaseCHECKBOX.UseVisualStyleBackColor = False
        '
        'ExitApplicationBackupDatabaseCHRCKBOX
        '
        Me.ExitApplicationBackupDatabaseCHRCKBOX.AutoSize = True
        Me.ExitApplicationBackupDatabaseCHRCKBOX.BackColor = System.Drawing.Color.Black
        Me.ExitApplicationBackupDatabaseCHRCKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExitApplicationBackupDatabaseCHRCKBOX.Location = New System.Drawing.Point(40, 173)
        Me.ExitApplicationBackupDatabaseCHRCKBOX.Name = "ExitApplicationBackupDatabaseCHRCKBOX"
        Me.ExitApplicationBackupDatabaseCHRCKBOX.Size = New System.Drawing.Size(112, 17)
        Me.ExitApplicationBackupDatabaseCHRCKBOX.TabIndex = 6
        Me.ExitApplicationBackupDatabaseCHRCKBOX.Text = "Backup Database"
        Me.ExitApplicationBackupDatabaseCHRCKBOX.UseVisualStyleBackColor = False
        '
        'ExitApplicationConfirmBUTTON
        '
        Me.ExitApplicationConfirmBUTTON.BackColor = System.Drawing.Color.DimGray
        Me.ExitApplicationConfirmBUTTON.BackgroundImage = CType(resources.GetObject("ExitApplicationConfirmBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ExitApplicationConfirmBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ExitApplicationConfirmBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExitApplicationConfirmBUTTON.Location = New System.Drawing.Point(199, 170)
        Me.ExitApplicationConfirmBUTTON.Name = "ExitApplicationConfirmBUTTON"
        Me.ExitApplicationConfirmBUTTON.Size = New System.Drawing.Size(86, 23)
        Me.ExitApplicationConfirmBUTTON.TabIndex = 9
        Me.ExitApplicationConfirmBUTTON.Text = "Exit"
        Me.ExitApplicationConfirmBUTTON.UseVisualStyleBackColor = False
        '
        'ExitApplicationCancelBUTTON
        '
        Me.ExitApplicationCancelBUTTON.BackColor = System.Drawing.Color.DimGray
        Me.ExitApplicationCancelBUTTON.BackgroundImage = CType(resources.GetObject("ExitApplicationCancelBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ExitApplicationCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ExitApplicationCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExitApplicationCancelBUTTON.Location = New System.Drawing.Point(291, 170)
        Me.ExitApplicationCancelBUTTON.Name = "ExitApplicationCancelBUTTON"
        Me.ExitApplicationCancelBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.ExitApplicationCancelBUTTON.TabIndex = 8
        Me.ExitApplicationCancelBUTTON.Text = "Cancel"
        Me.ExitApplicationCancelBUTTON.UseVisualStyleBackColor = False
        '
        'ExitApplication
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(398, 223)
        Me.ControlBox = False
        Me.Controls.Add(Me.ExitApplicationHeaderLABEL)
        Me.Controls.Add(Me.ExitApplicationInformationLABEL)
        Me.Controls.Add(Me.ExitApplicationConfirmBUTTON)
        Me.Controls.Add(Me.ExitApplicationCancelBUTTON)
        Me.Controls.Add(Me.ExitApplicationSaveDatabaseCHECKBOX)
        Me.Controls.Add(Me.ExitApplicationBackupDatabaseCHRCKBOX)
        Me.MaximumSize = New System.Drawing.Size(414, 262)
        Me.MinimumSize = New System.Drawing.Size(414, 262)
        Me.Name = "ExitApplication"
        Me.Text = "Exit DiaBase"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ExitApplicationHeaderLABEL As System.Windows.Forms.Label
    Friend WithEvents ExitApplicationInformationLABEL As System.Windows.Forms.Label
    Friend WithEvents ExitApplicationConfirmBUTTON As System.Windows.Forms.Button
    Friend WithEvents ExitApplicationCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents ExitApplicationSaveDatabaseCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents ExitApplicationBackupDatabaseCHRCKBOX As System.Windows.Forms.CheckBox
End Class
