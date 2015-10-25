<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Export
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Export))
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.CancelButton1 = New System.Windows.Forms.Button()
        Me.ExportButton = New System.Windows.Forms.Button()
        Me.CreateFileBUTTON = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.NewDatabaseLABEL = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SavedDatabasesLISTBOX = New System.Windows.Forms.ListBox()
        Me.OpenDatabaseCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.DeleteItemsCHECKBOX = New System.Windows.Forms.CheckBox()
        Me.MoveItemsExportBUTTON = New System.Windows.Forms.Button()
        Me.MoveItemsCancelBUTTON = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.DatabaseFilenameTEXTBOX = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.BackColor = System.Drawing.SystemColors.ControlText
        Me.CheckBox1.ForeColor = System.Drawing.Color.BurlyWood
        Me.CheckBox1.Location = New System.Drawing.Point(1130, 110)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(146, 17)
        Me.CheckBox1.TabIndex = 916
        Me.CheckBox1.Text = "Delete Item/s after export"
        Me.CheckBox1.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(876, 341)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(115, 20)
        Me.Label2.TabIndex = 915
        Me.Label2.Text = "New File Name"
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.MenuText
        Me.TextBox1.Location = New System.Drawing.Point(875, 364)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(177, 20)
        Me.TextBox1.TabIndex = 914
        '
        'CancelButton1
        '
        Me.CancelButton1.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.CancelButton1.BackgroundImage = CType(resources.GetObject("CancelButton1.BackgroundImage"), System.Drawing.Image)
        Me.CancelButton1.FlatAppearance.BorderSize = 2
        Me.CancelButton1.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.CancelButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CancelButton1.ForeColor = System.Drawing.Color.BurlyWood
        Me.CancelButton1.Location = New System.Drawing.Point(1141, 264)
        Me.CancelButton1.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.CancelButton1.Name = "CancelButton1"
        Me.CancelButton1.Size = New System.Drawing.Size(146, 25)
        Me.CancelButton1.TabIndex = 913
        Me.CancelButton1.Text = "Cancel"
        Me.CancelButton1.UseVisualStyleBackColor = False
        '
        'ExportButton
        '
        Me.ExportButton.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ExportButton.BackgroundImage = CType(resources.GetObject("ExportButton.BackgroundImage"), System.Drawing.Image)
        Me.ExportButton.Enabled = False
        Me.ExportButton.FlatAppearance.BorderSize = 2
        Me.ExportButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ExportButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ExportButton.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExportButton.Location = New System.Drawing.Point(1141, 177)
        Me.ExportButton.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ExportButton.Name = "ExportButton"
        Me.ExportButton.Size = New System.Drawing.Size(146, 25)
        Me.ExportButton.TabIndex = 912
        Me.ExportButton.Text = "Expot"
        Me.ExportButton.UseVisualStyleBackColor = False
        '
        'CreateFileBUTTON
        '
        Me.CreateFileBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.CreateFileBUTTON.BackgroundImage = CType(resources.GetObject("CreateFileBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.CreateFileBUTTON.FlatAppearance.BorderSize = 2
        Me.CreateFileBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.CreateFileBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CreateFileBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.CreateFileBUTTON.Location = New System.Drawing.Point(1141, 361)
        Me.CreateFileBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.CreateFileBUTTON.Name = "CreateFileBUTTON"
        Me.CreateFileBUTTON.Size = New System.Drawing.Size(146, 25)
        Me.CreateFileBUTTON.TabIndex = 911
        Me.CreateFileBUTTON.Text = "Create"
        Me.CreateFileBUTTON.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(882, 73)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 20)
        Me.Label1.TabIndex = 910
        Me.Label1.Text = "Available Files"
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(875, 98)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(228, 225)
        Me.ListBox1.TabIndex = 909
        '
        'NewDatabaseLABEL
        '
        Me.NewDatabaseLABEL.BackColor = System.Drawing.Color.Black
        Me.NewDatabaseLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewDatabaseLABEL.ForeColor = System.Drawing.Color.SeaShell
        Me.NewDatabaseLABEL.Location = New System.Drawing.Point(51, 85)
        Me.NewDatabaseLABEL.Name = "NewDatabaseLABEL"
        Me.NewDatabaseLABEL.Size = New System.Drawing.Size(560, 82)
        Me.NewDatabaseLABEL.TabIndex = 923
        Me.NewDatabaseLABEL.Text = resources.GetString("NewDatabaseLABEL.Text")
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.BurlyWood
        Me.Label7.Location = New System.Drawing.Point(54, 177)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(3, 173)
        Me.Label7.TabIndex = 932
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.BurlyWood
        Me.Label6.Location = New System.Drawing.Point(339, 177)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(3, 173)
        Me.Label6.TabIndex = 931
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.BurlyWood
        Me.Label5.Location = New System.Drawing.Point(56, 347)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(285, 3)
        Me.Label5.TabIndex = 930
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.BurlyWood
        Me.Label3.Location = New System.Drawing.Point(56, 177)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(285, 3)
        Me.Label3.TabIndex = 929
        '
        'SavedDatabasesLISTBOX
        '
        Me.SavedDatabasesLISTBOX.BackColor = System.Drawing.Color.Black
        Me.SavedDatabasesLISTBOX.ForeColor = System.Drawing.Color.White
        Me.SavedDatabasesLISTBOX.FormattingEnabled = True
        Me.SavedDatabasesLISTBOX.Location = New System.Drawing.Point(54, 177)
        Me.SavedDatabasesLISTBOX.Name = "SavedDatabasesLISTBOX"
        Me.SavedDatabasesLISTBOX.Size = New System.Drawing.Size(288, 173)
        Me.SavedDatabasesLISTBOX.TabIndex = 928
        '
        'OpenDatabaseCHECKBOX
        '
        Me.OpenDatabaseCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.OpenDatabaseCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.OpenDatabaseCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.OpenDatabaseCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.OpenDatabaseCHECKBOX.Location = New System.Drawing.Point(359, 160)
        Me.OpenDatabaseCHECKBOX.Name = "OpenDatabaseCHECKBOX"
        Me.OpenDatabaseCHECKBOX.Size = New System.Drawing.Size(249, 60)
        Me.OpenDatabaseCHECKBOX.TabIndex = 927
        Me.OpenDatabaseCHECKBOX.Text = "Open Desination Database Once Items Are Exported"
        Me.OpenDatabaseCHECKBOX.UseVisualStyleBackColor = False
        '
        'DeleteItemsCHECKBOX
        '
        Me.DeleteItemsCHECKBOX.BackColor = System.Drawing.Color.Black
        Me.DeleteItemsCHECKBOX.FlatAppearance.BorderColor = System.Drawing.Color.BurlyWood
        Me.DeleteItemsCHECKBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeleteItemsCHECKBOX.ForeColor = System.Drawing.Color.BurlyWood
        Me.DeleteItemsCHECKBOX.Location = New System.Drawing.Point(359, 219)
        Me.DeleteItemsCHECKBOX.Name = "DeleteItemsCHECKBOX"
        Me.DeleteItemsCHECKBOX.Size = New System.Drawing.Size(249, 52)
        Me.DeleteItemsCHECKBOX.TabIndex = 926
        Me.DeleteItemsCHECKBOX.Text = "DO NOT Delete Exported Items From The Source Database"
        Me.DeleteItemsCHECKBOX.UseVisualStyleBackColor = False
        '
        'MoveItemsExportBUTTON
        '
        Me.MoveItemsExportBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.MoveItemsExportBUTTON.FlatAppearance.BorderSize = 2
        Me.MoveItemsExportBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.MoveItemsExportBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoveItemsExportBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.MoveItemsExportBUTTON.Location = New System.Drawing.Point(451, 358)
        Me.MoveItemsExportBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.MoveItemsExportBUTTON.Name = "MoveItemsExportBUTTON"
        Me.MoveItemsExportBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.MoveItemsExportBUTTON.TabIndex = 925
        Me.MoveItemsExportBUTTON.Text = "Export"
        Me.MoveItemsExportBUTTON.UseVisualStyleBackColor = False
        '
        'MoveItemsCancelBUTTON
        '
        Me.MoveItemsCancelBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.MoveItemsCancelBUTTON.FlatAppearance.BorderSize = 2
        Me.MoveItemsCancelBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.MoveItemsCancelBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.MoveItemsCancelBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.MoveItemsCancelBUTTON.Location = New System.Drawing.Point(534, 358)
        Me.MoveItemsCancelBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.MoveItemsCancelBUTTON.Name = "MoveItemsCancelBUTTON"
        Me.MoveItemsCancelBUTTON.Size = New System.Drawing.Size(73, 25)
        Me.MoveItemsCancelBUTTON.TabIndex = 924
        Me.MoveItemsCancelBUTTON.Text = "Cancel"
        Me.MoveItemsCancelBUTTON.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Black
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label4.Location = New System.Drawing.Point(104, 50)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(442, 23)
        Me.Label4.TabIndex = 922
        Me.Label4.Text = "ENTER DESTINATION DATABASE"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.BurlyWood
        Me.Label8.Location = New System.Drawing.Point(427, 363)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(3, 20)
        Me.Label8.TabIndex = 921
        '
        'Label26
        '
        Me.Label26.BackColor = System.Drawing.Color.BurlyWood
        Me.Label26.Location = New System.Drawing.Point(54, 362)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(3, 20)
        Me.Label26.TabIndex = 920
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.BurlyWood
        Me.Label9.Location = New System.Drawing.Point(54, 380)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(376, 3)
        Me.Label9.TabIndex = 919
        '
        'Label17
        '
        Me.Label17.BackColor = System.Drawing.Color.BurlyWood
        Me.Label17.Location = New System.Drawing.Point(54, 361)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(376, 3)
        Me.Label17.TabIndex = 918
        '
        'DatabaseFilenameTEXTBOX
        '
        Me.DatabaseFilenameTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseFilenameTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseFilenameTEXTBOX.Location = New System.Drawing.Point(55, 362)
        Me.DatabaseFilenameTEXTBOX.Name = "DatabaseFilenameTEXTBOX"
        Me.DatabaseFilenameTEXTBOX.Size = New System.Drawing.Size(374, 20)
        Me.DatabaseFilenameTEXTBOX.TabIndex = 917
        '
        'Export
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(650, 457)
        Me.Controls.Add(Me.NewDatabaseLABEL)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SavedDatabasesLISTBOX)
        Me.Controls.Add(Me.OpenDatabaseCHECKBOX)
        Me.Controls.Add(Me.DeleteItemsCHECKBOX)
        Me.Controls.Add(Me.MoveItemsExportBUTTON)
        Me.Controls.Add(Me.MoveItemsCancelBUTTON)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.DatabaseFilenameTEXTBOX)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CancelButton1)
        Me.Controls.Add(Me.ExportButton)
        Me.Controls.Add(Me.CreateFileBUTTON)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.MinimumSize = New System.Drawing.Size(493, 440)
        Me.Name = "Export"
        Me.Text = "Export"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label2 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents CancelButton1 As Button
    Friend WithEvents ExportButton As Button
    Friend WithEvents CreateFileBUTTON As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents NewDatabaseLABEL As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents SavedDatabasesLISTBOX As ListBox
    Friend WithEvents OpenDatabaseCHECKBOX As CheckBox
    Friend WithEvents DeleteItemsCHECKBOX As CheckBox
    Friend WithEvents MoveItemsExportBUTTON As Button
    Friend WithEvents MoveItemsCancelBUTTON As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label26 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label17 As Label
    Friend WithEvents DatabaseFilenameTEXTBOX As TextBox
End Class
