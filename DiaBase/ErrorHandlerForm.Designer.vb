<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ErrorHandlerForm
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
        Me.ErrorTrapHeaderLABEL = New System.Windows.Forms.Label()
        Me.ErrorTrapYesBUTTON = New System.Windows.Forms.Button()
        Me.ErrorTrapNoBUTTON = New System.Windows.Forms.Button()
        Me.ErrorTrapMessageTEXTBOX = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'ErrorTrapHeaderLABEL
        '
        Me.ErrorTrapHeaderLABEL.BackColor = System.Drawing.Color.Black
        Me.ErrorTrapHeaderLABEL.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ErrorTrapHeaderLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.ErrorTrapHeaderLABEL.Location = New System.Drawing.Point(38, 36)
        Me.ErrorTrapHeaderLABEL.Name = "ErrorTrapHeaderLABEL"
        Me.ErrorTrapHeaderLABEL.Size = New System.Drawing.Size(352, 23)
        Me.ErrorTrapHeaderLABEL.TabIndex = 896
        Me.ErrorTrapHeaderLABEL.Text = "ErrorTrapHeadingLABEL - Set At Runtime"
        Me.ErrorTrapHeaderLABEL.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'ErrorTrapYesBUTTON
        '
        Me.ErrorTrapYesBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ErrorTrapYesBUTTON.FlatAppearance.BorderSize = 2
        Me.ErrorTrapYesBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ErrorTrapYesBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ErrorTrapYesBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ErrorTrapYesBUTTON.Location = New System.Drawing.Point(212, 236)
        Me.ErrorTrapYesBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ErrorTrapYesBUTTON.Name = "ErrorTrapYesBUTTON"
        Me.ErrorTrapYesBUTTON.Size = New System.Drawing.Size(87, 25)
        Me.ErrorTrapYesBUTTON.TabIndex = 895
        Me.ErrorTrapYesBUTTON.Text = "Yes Button"
        Me.ErrorTrapYesBUTTON.UseVisualStyleBackColor = False
        '
        'ErrorTrapNoBUTTON
        '
        Me.ErrorTrapNoBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.ErrorTrapNoBUTTON.FlatAppearance.BorderSize = 2
        Me.ErrorTrapNoBUTTON.FlatAppearance.MouseDownBackColor = System.Drawing.Color.Black
        Me.ErrorTrapNoBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ErrorTrapNoBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.ErrorTrapNoBUTTON.Location = New System.Drawing.Point(310, 236)
        Me.ErrorTrapNoBUTTON.Margin = New System.Windows.Forms.Padding(3, 1, 3, 1)
        Me.ErrorTrapNoBUTTON.Name = "ErrorTrapNoBUTTON"
        Me.ErrorTrapNoBUTTON.Size = New System.Drawing.Size(87, 25)
        Me.ErrorTrapNoBUTTON.TabIndex = 894
        Me.ErrorTrapNoBUTTON.Text = "No Button"
        Me.ErrorTrapNoBUTTON.UseVisualStyleBackColor = False
        '
        'ErrorTrapMessageTEXTBOX
        '
        Me.ErrorTrapMessageTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.ErrorTrapMessageTEXTBOX.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ErrorTrapMessageTEXTBOX.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ErrorTrapMessageTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.ErrorTrapMessageTEXTBOX.Location = New System.Drawing.Point(38, 68)
        Me.ErrorTrapMessageTEXTBOX.Multiline = True
        Me.ErrorTrapMessageTEXTBOX.Name = "ErrorTrapMessageTEXTBOX"
        Me.ErrorTrapMessageTEXTBOX.Size = New System.Drawing.Size(352, 152)
        Me.ErrorTrapMessageTEXTBOX.TabIndex = 897
        Me.ErrorTrapMessageTEXTBOX.Text = "ErrorTrap Message Textbox - Set At Runtime"
        '
        'ErrorHandlerForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.Setting
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(424, 294)
        Me.Controls.Add(Me.ErrorTrapMessageTEXTBOX)
        Me.Controls.Add(Me.ErrorTrapHeaderLABEL)
        Me.Controls.Add(Me.ErrorTrapYesBUTTON)
        Me.Controls.Add(Me.ErrorTrapNoBUTTON)
        Me.Name = "ErrorHandlerForm"
        Me.Text = "DiaBase Error Handler"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ErrorTrapHeaderLABEL As System.Windows.Forms.Label
    Friend WithEvents ErrorTrapYesBUTTON As System.Windows.Forms.Button
    Friend WithEvents ErrorTrapNoBUTTON As System.Windows.Forms.Button
    Friend WithEvents ErrorTrapMessageTEXTBOX As System.Windows.Forms.TextBox
End Class
