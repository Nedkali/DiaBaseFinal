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
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.KeepManagerOpenCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.ManagerRefreshBUTTON = New System.Windows.Forms.Button()
        Me.ManagerSummaryBUTTON = New System.Windows.Forms.Button()
        Me.ManagerCreateBUTTON = New System.Windows.Forms.Button()
        Me.ManagerRenameBUTTON = New System.Windows.Forms.Button()
        Me.ManagerDeleteBUTTON = New System.Windows.Forms.Button()
        Me.ManagerSaveFirstCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.ManagerOpenBUTTON = New System.Windows.Forms.Button()
        Me.ManagerCancelBUTTON = New System.Windows.Forms.Button()
        Me.DatabaseManagerSavedDatabasesLISTBOX = New System.Windows.Forms.ListBox()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox1.BackgroundImage = Global.DiaBase.My.Resources.Resources.Header
        Me.PictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox1.Location = New System.Drawing.Point(-2, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(568, 33)
        Me.PictureBox1.TabIndex = 969
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PictureBox2.BackgroundImage = Global.DiaBase.My.Resources.Resources.Footer
        Me.PictureBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.PictureBox2.Location = New System.Drawing.Point(-2, 258)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(568, 29)
        Me.PictureBox2.TabIndex = 970
        Me.PictureBox2.TabStop = False
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.BackColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(439, 39)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(2, 147)
        Me.Label2.TabIndex = 1040
        '
        'Label31
        '
        Me.Label31.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label31.BackColor = System.Drawing.Color.BurlyWood
        Me.Label31.Location = New System.Drawing.Point(32, 39)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(2, 147)
        Me.Label31.TabIndex = 1039
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.BackColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(32, 184)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(409, 2)
        Me.Label1.TabIndex = 1038
        '
        'Label29
        '
        Me.Label29.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label29.BackColor = System.Drawing.Color.BurlyWood
        Me.Label29.Location = New System.Drawing.Point(32, 38)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(409, 2)
        Me.Label29.TabIndex = 1037
        '
        'KeepManagerOpenCHECKBOX
        '
        Me.KeepManagerOpenCHECKBOX.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.KeepManagerOpenCHECKBOX.AutoSize = True
        Me.KeepManagerOpenCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.KeepManagerOpenCHECKBOX.Checked = True
        Me.KeepManagerOpenCHECKBOX.CheckState = System.Windows.Forms.CheckState.Checked
        Me.KeepManagerOpenCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.KeepManagerOpenCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.KeepManagerOpenCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.KeepManagerOpenCHECKBOX.Location = New System.Drawing.Point(32, 223)
        Me.KeepManagerOpenCHECKBOX.Name = "KeepManagerOpenCHECKBOX"
        Me.KeepManagerOpenCHECKBOX.Size = New System.Drawing.Size(182, 20)
        Me.KeepManagerOpenCHECKBOX.TabIndex = 1036
        Me.KeepManagerOpenCHECKBOX.Text = "Keep Manager Open After"
        Me.KeepManagerOpenCHECKBOX.UseVisualStyleBackColor = False
        '
        'ManagerRefreshBUTTON
        '
        Me.ManagerRefreshBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerRefreshBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerRefreshBUTTON.BackgroundImage = CType(resources.GetObject("ManagerRefreshBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerRefreshBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerRefreshBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerRefreshBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerRefreshBUTTON.Location = New System.Drawing.Point(373, 223)
        Me.ManagerRefreshBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerRefreshBUTTON.Name = "ManagerRefreshBUTTON"
        Me.ManagerRefreshBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerRefreshBUTTON.TabIndex = 1035
        Me.ManagerRefreshBUTTON.Text = "Refresh"
        Me.ManagerRefreshBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerSummaryBUTTON
        '
        Me.ManagerSummaryBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerSummaryBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerSummaryBUTTON.BackgroundImage = CType(resources.GetObject("ManagerSummaryBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerSummaryBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerSummaryBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerSummaryBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerSummaryBUTTON.Location = New System.Drawing.Point(460, 159)
        Me.ManagerSummaryBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerSummaryBUTTON.Name = "ManagerSummaryBUTTON"
        Me.ManagerSummaryBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerSummaryBUTTON.TabIndex = 1034
        Me.ManagerSummaryBUTTON.Text = "Info"
        Me.ManagerSummaryBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerCreateBUTTON
        '
        Me.ManagerCreateBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerCreateBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerCreateBUTTON.BackgroundImage = CType(resources.GetObject("ManagerCreateBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerCreateBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerCreateBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerCreateBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerCreateBUTTON.Location = New System.Drawing.Point(460, 67)
        Me.ManagerCreateBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerCreateBUTTON.Name = "ManagerCreateBUTTON"
        Me.ManagerCreateBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerCreateBUTTON.TabIndex = 1033
        Me.ManagerCreateBUTTON.Text = "Create"
        Me.ManagerCreateBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerRenameBUTTON
        '
        Me.ManagerRenameBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerRenameBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerRenameBUTTON.BackgroundImage = CType(resources.GetObject("ManagerRenameBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerRenameBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerRenameBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerRenameBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerRenameBUTTON.Location = New System.Drawing.Point(460, 96)
        Me.ManagerRenameBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerRenameBUTTON.Name = "ManagerRenameBUTTON"
        Me.ManagerRenameBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerRenameBUTTON.TabIndex = 1032
        Me.ManagerRenameBUTTON.Text = "Rename"
        Me.ManagerRenameBUTTON.UseCompatibleTextRendering = True
        Me.ManagerRenameBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerDeleteBUTTON
        '
        Me.ManagerDeleteBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerDeleteBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerDeleteBUTTON.BackgroundImage = CType(resources.GetObject("ManagerDeleteBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerDeleteBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerDeleteBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerDeleteBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerDeleteBUTTON.Location = New System.Drawing.Point(460, 125)
        Me.ManagerDeleteBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerDeleteBUTTON.Name = "ManagerDeleteBUTTON"
        Me.ManagerDeleteBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerDeleteBUTTON.TabIndex = 1031
        Me.ManagerDeleteBUTTON.Text = "Delete"
        Me.ManagerDeleteBUTTON.UseVisualStyleBackColor = False
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
        Me.ManagerSaveFirstCHECKBOX.Location = New System.Drawing.Point(32, 197)
        Me.ManagerSaveFirstCHECKBOX.Name = "ManagerSaveFirstCHECKBOX"
        Me.ManagerSaveFirstCHECKBOX.Size = New System.Drawing.Size(195, 20)
        Me.ManagerSaveFirstCHECKBOX.TabIndex = 1030
        Me.ManagerSaveFirstCHECKBOX.Text = "Save Current Database First"
        Me.ManagerSaveFirstCHECKBOX.UseVisualStyleBackColor = False
        '
        'ManagerOpenBUTTON
        '
        Me.ManagerOpenBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerOpenBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerOpenBUTTON.BackgroundImage = CType(resources.GetObject("ManagerOpenBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerOpenBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerOpenBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerOpenBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerOpenBUTTON.Location = New System.Drawing.Point(460, 38)
        Me.ManagerOpenBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerOpenBUTTON.Name = "ManagerOpenBUTTON"
        Me.ManagerOpenBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerOpenBUTTON.TabIndex = 1029
        Me.ManagerOpenBUTTON.Text = "Open"
        Me.ManagerOpenBUTTON.UseVisualStyleBackColor = False
        '
        'ManagerCancelBUTTON
        '
        Me.ManagerCancelBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ManagerCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ManagerCancelBUTTON.BackgroundImage = CType(resources.GetObject("ManagerCancelBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.ManagerCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ManagerCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ManagerCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ManagerCancelBUTTON.Location = New System.Drawing.Point(460, 223)
        Me.ManagerCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ManagerCancelBUTTON.Name = "ManagerCancelBUTTON"
        Me.ManagerCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.ManagerCancelBUTTON.TabIndex = 1028
        Me.ManagerCancelBUTTON.Text = "Close"
        Me.ManagerCancelBUTTON.UseVisualStyleBackColor = False
        '
        'DatabaseManagerSavedDatabasesLISTBOX
        '
        Me.DatabaseManagerSavedDatabasesLISTBOX.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseManagerSavedDatabasesLISTBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseManagerSavedDatabasesLISTBOX.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DatabaseManagerSavedDatabasesLISTBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseManagerSavedDatabasesLISTBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseManagerSavedDatabasesLISTBOX.FormattingEnabled = True
        Me.DatabaseManagerSavedDatabasesLISTBOX.ItemHeight = 16
        Me.DatabaseManagerSavedDatabasesLISTBOX.Location = New System.Drawing.Point(32, 39)
        Me.DatabaseManagerSavedDatabasesLISTBOX.Name = "DatabaseManagerSavedDatabasesLISTBOX"
        Me.DatabaseManagerSavedDatabasesLISTBOX.ScrollAlwaysVisible = True
        Me.DatabaseManagerSavedDatabasesLISTBOX.Size = New System.Drawing.Size(409, 144)
        Me.DatabaseManagerSavedDatabasesLISTBOX.TabIndex = 1027
        '
        'DatabaseManager
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Black
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(564, 288)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.KeepManagerOpenCHECKBOX)
        Me.Controls.Add(Me.ManagerRefreshBUTTON)
        Me.Controls.Add(Me.ManagerSummaryBUTTON)
        Me.Controls.Add(Me.ManagerCreateBUTTON)
        Me.Controls.Add(Me.ManagerRenameBUTTON)
        Me.Controls.Add(Me.ManagerDeleteBUTTON)
        Me.Controls.Add(Me.ManagerSaveFirstCHECKBOX)
        Me.Controls.Add(Me.ManagerOpenBUTTON)
        Me.Controls.Add(Me.ManagerCancelBUTTON)
        Me.Controls.Add(Me.DatabaseManagerSavedDatabasesLISTBOX)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.PictureBox2)
        Me.MaximumSize = New System.Drawing.Size(750, 432)
        Me.MinimumSize = New System.Drawing.Size(580, 304)
        Me.Name = "DatabaseManager"
        Me.Text = "Database Manager"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents PictureBox2 As PictureBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label31 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Label29 As Label
    Friend WithEvents KeepManagerOpenCHECKBOX As CheckBox
    Friend WithEvents ManagerRefreshBUTTON As Button
    Friend WithEvents ManagerSummaryBUTTON As Button
    Friend WithEvents ManagerCreateBUTTON As Button
    Friend WithEvents ManagerRenameBUTTON As Button
    Friend WithEvents ManagerDeleteBUTTON As Button
    Friend WithEvents ManagerSaveFirstCHECKBOX As CheckBox
    Friend WithEvents ManagerOpenBUTTON As Button
    Friend WithEvents ManagerCancelBUTTON As Button
    Friend WithEvents DatabaseManagerSavedDatabasesLISTBOX As ListBox
End Class
