<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DatabaseManager
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DatabaseManager))
        Me.ManagerSaveFirstCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.ManagerOpenBUTTON = New System.Windows.Forms.Button()
        Me.ManagerCancelBUTTON = New System.Windows.Forms.Button()
        Me.ManagerHeaderLABEL = New System.Windows.Forms.Label()
        Me.ManagerDatabasesLISTBOX = New System.Windows.Forms.ListBox()
        Me.ManagerDeleteBUTTON = New System.Windows.Forms.Button()
        Me.ManagerRenameBUTTON = New System.Windows.Forms.Button()
        Me.ManagerMergeBUTTON = New System.Windows.Forms.Button()
        Me.ManagerCreateBUTTON = New System.Windows.Forms.Button()
        Me.ManagerSummaryBUTTON = New System.Windows.Forms.Button()
        Me.ManagerRefreshBUTTON = New System.Windows.Forms.Button()
        Me.KeepManagerOpenCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'ManagerSaveFirstCHECKBOX
        '
        Me.ManagerSaveFirstCHECKBOX.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerSaveFirstCHECKBOX.AutoSize = True
        Me.ManagerSaveFirstCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.ManagerSaveFirstCHECKBOX.Checked = True
        Me.ManagerSaveFirstCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.ManagerSaveFirstCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.ManagerSaveFirstCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ManagerSaveFirstCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerSaveFirstCHECKBOX.Location = New System.Drawing.Point(47, 281)
        Me.ManagerSaveFirstCHECKBOX.Name = "ManagerSaveFirstCHECKBOX"
        Me.ManagerSaveFirstCHECKBOX.Size = New System.Drawing.Size(198, 20)
        Me.ManagerSaveFirstCHECKBOX.TabIndex = 911
        Me.ManagerSaveFirstCHECKBOX.Text = "Save Current Database  First"
        Me.ManagerSaveFirstCHECKBOX.UseVisualStyleBackColor = False
        '
        'ManagerOpenBUTTON
        '
        Me.ManagerOpenBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerOpenBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerOpenBUTTON.BackgroundImage = CType(resources.GetObject("ManagerOpenBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerOpenBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerOpenBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerOpenBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerOpenBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerOpenBUTTON.Location = New System.Drawing.Point(321, 76)
        Me.ManagerOpenBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerOpenBUTTON.Name = "ManagerOpenBUTTON"
        Me.ManagerOpenBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerOpenBUTTON.TabIndex = 902
        Me.ManagerOpenBUTTON.Text = "Open"
        Me.ManagerOpenBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerCancelBUTTON
        '
        Me.ManagerCancelBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerCancelBUTTON.BackgroundImage = CType(resources.GetObject("ManagerCancelBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerCancelBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerCancelBUTTON.Location = New System.Drawing.Point(321, 304)
        Me.ManagerCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerCancelBUTTON.Name = "ManagerCancelBUTTON"
        Me.ManagerCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerCancelBUTTON.TabIndex = 901
        Me.ManagerCancelBUTTON.Text = "Close"
        Me.ManagerCancelBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerHeaderLABEL
        '
        Me.ManagerHeaderLABEL.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerHeaderLABEL.BackColor = System.Drawing.Color.Black
        Me.ManagerHeaderLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ManagerHeaderLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerHeaderLABEL.Location = New System.Drawing.Point(44, 33)
        Me.ManagerHeaderLABEL.Name = "ManagerHeaderLABEL"
        Me.ManagerHeaderLABEL.Size = New System.Drawing.Size(350, 23)
        Me.ManagerHeaderLABEL.TabIndex = 899
        Me.ManagerHeaderLABEL.Text = "DATABASE FILE MANAGER "
        Me.ManagerHeaderLABEL.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ManagerDatabasesLISTBOX
        '
        Me.ManagerDatabasesLISTBOX.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerDatabasesLISTBOX.BackColor = System.Drawing.Color.Black
        Me.ManagerDatabasesLISTBOX.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ManagerDatabasesLISTBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ManagerDatabasesLISTBOX.ForeColor = System.Drawing.Color.White
        Me.ManagerDatabasesLISTBOX.FormattingEnabled = True
        Me.ManagerDatabasesLISTBOX.ItemHeight = 16
        Me.ManagerDatabasesLISTBOX.Location = New System.Drawing.Point(48, 76)
        Me.ManagerDatabasesLISTBOX.Name = "ManagerDatabasesLISTBOX"
        Me.ManagerDatabasesLISTBOX.ScrollAlwaysVisible = True
        Me.ManagerDatabasesLISTBOX.Size = New System.Drawing.Size(238, 194)
        Me.ManagerDatabasesLISTBOX.TabIndex = 896
        '
        'ManagerDeleteBUTTON
        '
        Me.ManagerDeleteBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerDeleteBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerDeleteBUTTON.BackgroundImage = CType(resources.GetObject("ManagerDeleteBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerDeleteBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerDeleteBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerDeleteBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerDeleteBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerDeleteBUTTON.Location = New System.Drawing.Point(321, 164)
        Me.ManagerDeleteBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerDeleteBUTTON.Name = "ManagerDeleteBUTTON"
        Me.ManagerDeleteBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerDeleteBUTTON.TabIndex = 912
        Me.ManagerDeleteBUTTON.Text = "Delete"
        Me.ManagerDeleteBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerRenameBUTTON
        '
        Me.ManagerRenameBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerRenameBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerRenameBUTTON.BackgroundImage = CType(resources.GetObject("ManagerRenameBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerRenameBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerRenameBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerRenameBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerRenameBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerRenameBUTTON.Location = New System.Drawing.Point(321, 134)
        Me.ManagerRenameBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerRenameBUTTON.Name = "ManagerRenameBUTTON"
        Me.ManagerRenameBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerRenameBUTTON.TabIndex = 913
        Me.ManagerRenameBUTTON.Text = "Rename"
        Me.ManagerRenameBUTTON.UseCompatibleTextRendering = True
        Me.ManagerRenameBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerMergeBUTTON
        '
        Me.ManagerMergeBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerMergeBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerMergeBUTTON.BackgroundImage = CType(resources.GetObject("ManagerMergeBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerMergeBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerMergeBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerMergeBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerMergeBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerMergeBUTTON.Location = New System.Drawing.Point(321, 234)
        Me.ManagerMergeBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerMergeBUTTON.Name = "ManagerMergeBUTTON"
        Me.ManagerMergeBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerMergeBUTTON.TabIndex = 915
        Me.ManagerMergeBUTTON.Text = "Merge"
        Me.ManagerMergeBUTTON.UseVisualStyleBackColor = False
        Me.ManagerMergeBUTTON.Visible = False
        '
        'ManagerCreateBUTTON
        '
        Me.ManagerCreateBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerCreateBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerCreateBUTTON.BackgroundImage = CType(resources.GetObject("ManagerCreateBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerCreateBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerCreateBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerCreateBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerCreateBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerCreateBUTTON.Location = New System.Drawing.Point(321, 105)
        Me.ManagerCreateBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerCreateBUTTON.Name = "ManagerCreateBUTTON"
        Me.ManagerCreateBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerCreateBUTTON.TabIndex = 916
        Me.ManagerCreateBUTTON.Text = "Create"
        Me.ManagerCreateBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerSummaryBUTTON
        '
        Me.ManagerSummaryBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerSummaryBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerSummaryBUTTON.BackgroundImage = CType(resources.GetObject("ManagerSummaryBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerSummaryBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerSummaryBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerSummaryBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerSummaryBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerSummaryBUTTON.Location = New System.Drawing.Point(321, 204)
        Me.ManagerSummaryBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerSummaryBUTTON.Name = "ManagerSummaryBUTTON"
        Me.ManagerSummaryBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerSummaryBUTTON.TabIndex = 917
        Me.ManagerSummaryBUTTON.Text = "Info"
        Me.ManagerSummaryBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerRefreshBUTTON
        '
        Me.ManagerRefreshBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerRefreshBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerRefreshBUTTON.BackgroundImage = CType(resources.GetObject("ManagerRefreshBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerRefreshBUTTON.FlatAppearance.BorderSize = 2
        Me.ManagerRefreshBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerRefreshBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerRefreshBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerRefreshBUTTON.Location = New System.Drawing.Point(321, 274)
        Me.ManagerRefreshBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerRefreshBUTTON.Name = "ManagerRefreshBUTTON"
        Me.ManagerRefreshBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerRefreshBUTTON.TabIndex = 918
        Me.ManagerRefreshBUTTON.Text = "Refresh"
        Me.ManagerRefreshBUTTON.UseVisualStyleBackColor = False
        '
        'KeepManagerOpenCHECKBOX
        '
        Me.KeepManagerOpenCHECKBOX.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.KeepManagerOpenCHECKBOX.AutoSize = True
        Me.KeepManagerOpenCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.KeepManagerOpenCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.KeepManagerOpenCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeepManagerOpenCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.KeepManagerOpenCHECKBOX.Location = New System.Drawing.Point(47, 307)
        Me.KeepManagerOpenCHECKBOX.Name = "KeepManagerOpenCHECKBOX"
        Me.KeepManagerOpenCHECKBOX.Size = New System.Drawing.Size(149, 20)
        Me.KeepManagerOpenCHECKBOX.TabIndex = 919
        Me.KeepManagerOpenCHECKBOX.Text = "Auto Close Manager"
        Me.KeepManagerOpenCHECKBOX.UseVisualStyleBackColor = False
        '
        'DatabaseManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.BigSettings
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(440, 362)
        Me.ControlBox = False
        Me.Controls.Add(Me.KeepManagerOpenCHECKBOX)
        Me.Controls.Add(Me.ManagerRefreshBUTTON)
        Me.Controls.Add(Me.ManagerSummaryBUTTON)
        Me.Controls.Add(Me.ManagerCreateBUTTON)
        Me.Controls.Add(Me.ManagerMergeBUTTON)
        Me.Controls.Add(Me.ManagerRenameBUTTON)
        Me.Controls.Add(Me.ManagerDeleteBUTTON)
        Me.Controls.Add(Me.ManagerSaveFirstCHECKBOX)
        Me.Controls.Add(Me.ManagerOpenBUTTON)
        Me.Controls.Add(Me.ManagerCancelBUTTON)
        Me.Controls.Add(Me.ManagerHeaderLABEL)
        Me.Controls.Add(Me.ManagerDatabasesLISTBOX)
        Me.MaximumSize = New System.Drawing.Size(665, 665)
        Me.MinimumSize = New System.Drawing.Size(400, 400)
        Me.Name = "DatabaseManager"
        Me.Text = "Database Manager"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ManagerSaveFirstCHECKBOX As System.Windows.Forms.CheckBox
    Friend WithEvents ManagerOpenBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerCancelBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerHeaderLABEL As System.Windows.Forms.Label
    Friend WithEvents ManagerDatabasesLISTBOX As System.Windows.Forms.ListBox
    Friend WithEvents ManagerDeleteBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerRenameBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerMergeBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerCreateBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerSummaryBUTTON As System.Windows.Forms.Button
    Friend WithEvents ManagerRefreshBUTTON As System.Windows.Forms.Button
    Friend WithEvents KeepManagerOpenCHECKBOX As System.Windows.Forms.CheckBox
End Class
