﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CreateFileBUTTON = New System.Windows.Forms.Button()
        Me.ExportButton = New System.Windows.Forms.Button()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(50, 70)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(228, 225)
        Me.ListBox1.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label1.Location = New System.Drawing.Point(57, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Available Files"
        '
        'CreateFileBUTTON
        '
        Me.CreateFileBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CreateFileBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.CreateFileBUTTON.BackgroundImage = CType(resources.GetObject("CreateFileBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.CreateFileBUTTON.FlatAppearance.BorderSize = 2
        Me.CreateFileBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.CreateFileBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CreateFileBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.CreateFileBUTTON.Location = New System.Drawing.Point(334, 331)
        Me.CreateFileBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.CreateFileBUTTON.Name = "CreateFileBUTTON"
        Me.CreateFileBUTTON.Size = New System.Drawing.Size(85, 25)
        Me.CreateFileBUTTON.TabIndex = 903
        Me.CreateFileBUTTON.Text = "Create"
        Me.CreateFileBUTTON.UseVisualStyleBackColor = False
        '
        'ExportButton
        '
        Me.ExportButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ExportButton.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ExportButton.BackgroundImage = CType(resources.GetObject("ExportButton.BackgroundImage"), System.Drawing.Image)
        Me.ExportButton.FlatAppearance.BorderSize = 2
        Me.ExportButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ExportButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ExportButton.ForeColor = System.Drawing.Color.BurlyWood
        Me.ExportButton.Location = New System.Drawing.Point(334, 85)
        Me.ExportButton.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ExportButton.Name = "ExportButton"
        Me.ExportButton.Size = New System.Drawing.Size(85, 25)
        Me.ExportButton.TabIndex = 904
        Me.ExportButton.Text = "Expot"
        Me.ExportButton.UseVisualStyleBackColor = False
        '
        'CancelButton
        '
        Me.CancelButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CancelButton.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.CancelButton.BackgroundImage = CType(resources.GetObject("CancelButton.BackgroundImage"), System.Drawing.Image)
        Me.CancelButton.FlatAppearance.BorderSize = 2
        Me.CancelButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.CancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.CancelButton.ForeColor = System.Drawing.Color.BurlyWood
        Me.CancelButton.Location = New System.Drawing.Point(334, 198)
        Me.CancelButton.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.CancelButton.Name = "CancelButton"
        Me.CancelButton.Size = New System.Drawing.Size(85, 25)
        Me.CancelButton.TabIndex = 905
        Me.CancelButton.Text = "Cancel"
        Me.CancelButton.UseVisualStyleBackColor = False
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox1.Location = New System.Drawing.Point(50, 336)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(129, 20)
        Me.TextBox1.TabIndex = 906
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Desktop
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.BurlyWood
        Me.Label2.Location = New System.Drawing.Point(51, 313)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(115, 20)
        Me.Label2.TabIndex = 907
        Me.Label2.Text = "New File Name"
        '
        'Export
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(496, 402)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.CancelButton)
        Me.Controls.Add(Me.ExportButton)
        Me.Controls.Add(Me.CreateFileBUTTON)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.Name = "Export"
        Me.Text = "Export"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListBox1 As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents CreateFileBUTTON As Button
    Friend WithEvents ExportButton As Button
    Friend WithEvents CancelButton As Button
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label2 As Label
End Class