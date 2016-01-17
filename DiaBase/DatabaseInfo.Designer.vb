<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DatabaseInfo
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DatabaseInfo))
        Me.DatabaseInfoDATAGRIDVIEW = New System.Windows.Forms.DataGridView()
        Me.ItemBaseCOLUMN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalCOLUMN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.RatioCOLUMN = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Label121 = New System.Windows.Forms.Label()
        Me.Label120 = New System.Windows.Forms.Label()
        Me.Label122 = New System.Windows.Forms.Label()
        Me.Label123 = New System.Windows.Forms.Label()
        Me.DatabaseInfoSelectedTEXTBOX = New System.Windows.Forms.TextBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label142 = New System.Windows.Forms.Label()
        Me.Label147 = New System.Windows.Forms.Label()
        Me.Label148 = New System.Windows.Forms.Label()
        Me.DatabaseInfoTotalTEXTBOX = New System.Windows.Forms.TextBox()
        Me.DatabaseInfoSelectedLABEL = New System.Windows.Forms.Label()
        Me.DatabaseInfoTotalLABEL = New System.Windows.Forms.Label()
        Me.DatabaseInfoTABCONTROL = New System.Windows.Forms.TabControl()
        Me.DatabaseInfoItemBaseTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoItemQualityTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoRunesTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoMulesTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoAccountsTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoRealmsTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoDatabaseTABPAGE = New System.Windows.Forms.TabPage()
        Me.DatabaseInfoCloseBUTTON = New System.Windows.Forms.Button()
        Me.DatabaseInfoRefreshBUTTON = New System.Windows.Forms.Button()
        Me.DatabaseInfoOpenContainingBUTTON = New System.Windows.Forms.Button()
        CType(Me.DatabaseInfoDATAGRIDVIEW, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.DatabaseInfoTABCONTROL.SuspendLayout()
        Me.DatabaseInfoItemBaseTABPAGE.SuspendLayout()
        Me.SuspendLayout()
        '
        'DatabaseInfoDATAGRIDVIEW
        '
        Me.DatabaseInfoDATAGRIDVIEW.AllowUserToAddRows = False
        Me.DatabaseInfoDATAGRIDVIEW.AllowUserToDeleteRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.Black
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle1.Format = "N1"
        DataGridViewCellStyle1.NullValue = Nothing
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.Color.Black
        Me.DatabaseInfoDATAGRIDVIEW.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DatabaseInfoDATAGRIDVIEW.BackgroundColor = System.Drawing.Color.Black
        Me.DatabaseInfoDATAGRIDVIEW.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DatabaseInfoDATAGRIDVIEW.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Sunken
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.Color.Goldenrod
        DataGridViewCellStyle2.Format = "N1"
        DataGridViewCellStyle2.NullValue = Nothing
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.Color.Goldenrod
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DatabaseInfoDATAGRIDVIEW.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DatabaseInfoDATAGRIDVIEW.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DatabaseInfoDATAGRIDVIEW.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.ItemBaseCOLUMN, Me.TotalCOLUMN, Me.RatioCOLUMN})
        Me.DatabaseInfoDATAGRIDVIEW.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DatabaseInfoDATAGRIDVIEW.GridColor = System.Drawing.Color.Gray
        Me.DatabaseInfoDATAGRIDVIEW.Location = New System.Drawing.Point(3, 3)
        Me.DatabaseInfoDATAGRIDVIEW.Name = "DatabaseInfoDATAGRIDVIEW"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle3.Format = "N1"
        DataGridViewCellStyle3.NullValue = Nothing
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DatabaseInfoDATAGRIDVIEW.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.DatabaseInfoDATAGRIDVIEW.RowHeadersVisible = False
        DataGridViewCellStyle4.BackColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle4.Format = "N1"
        DataGridViewCellStyle4.NullValue = Nothing
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.BurlyWood
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black
        Me.DatabaseInfoDATAGRIDVIEW.RowsDefaultCellStyle = DataGridViewCellStyle4
        Me.DatabaseInfoDATAGRIDVIEW.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.DatabaseInfoDATAGRIDVIEW.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DatabaseInfoDATAGRIDVIEW.ShowCellErrors = False
        Me.DatabaseInfoDATAGRIDVIEW.ShowCellToolTips = False
        Me.DatabaseInfoDATAGRIDVIEW.ShowEditingIcon = False
        Me.DatabaseInfoDATAGRIDVIEW.ShowRowErrors = False
        Me.DatabaseInfoDATAGRIDVIEW.Size = New System.Drawing.Size(517, 252)
        Me.DatabaseInfoDATAGRIDVIEW.TabIndex = 0
        '
        'ItemBaseCOLUMN
        '
        Me.ItemBaseCOLUMN.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.ItemBaseCOLUMN.HeaderText = "Item Base"
        Me.ItemBaseCOLUMN.Name = "ItemBaseCOLUMN"
        '
        'TotalCOLUMN
        '
        Me.TotalCOLUMN.HeaderText = "Total"
        Me.TotalCOLUMN.Name = "TotalCOLUMN"
        Me.TotalCOLUMN.Width = 60
        '
        'RatioCOLUMN
        '
        Me.RatioCOLUMN.HeaderText = "Ratio"
        Me.RatioCOLUMN.Name = "RatioCOLUMN"
        Me.RatioCOLUMN.Width = 60
        '
        'Label121
        '
        Me.Label121.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label121.BackColor = System.Drawing.Color.BurlyWood
        Me.Label121.Location = New System.Drawing.Point(41, 55)
        Me.Label121.Name = "Label121"
        Me.Label121.Size = New System.Drawing.Size(389, 2)
        Me.Label121.TabIndex = 1915
        '
        'Label120
        '
        Me.Label120.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label120.BackColor = System.Drawing.Color.BurlyWood
        Me.Label120.Location = New System.Drawing.Point(41, 73)
        Me.Label120.Name = "Label120"
        Me.Label120.Size = New System.Drawing.Size(389, 2)
        Me.Label120.TabIndex = 1914
        '
        'Label122
        '
        Me.Label122.BackColor = System.Drawing.Color.BurlyWood
        Me.Label122.Location = New System.Drawing.Point(41, 55)
        Me.Label122.Name = "Label122"
        Me.Label122.Size = New System.Drawing.Size(3, 20)
        Me.Label122.TabIndex = 1913
        '
        'Label123
        '
        Me.Label123.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label123.BackColor = System.Drawing.Color.BurlyWood
        Me.Label123.Location = New System.Drawing.Point(428, 55)
        Me.Label123.Name = "Label123"
        Me.Label123.Size = New System.Drawing.Size(3, 20)
        Me.Label123.TabIndex = 1912
        '
        'DatabaseInfoSelectedTEXTBOX
        '
        Me.DatabaseInfoSelectedTEXTBOX.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoSelectedTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseInfoSelectedTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseInfoSelectedTEXTBOX.Location = New System.Drawing.Point(41, 55)
        Me.DatabaseInfoSelectedTEXTBOX.Name = "DatabaseInfoSelectedTEXTBOX"
        Me.DatabaseInfoSelectedTEXTBOX.Size = New System.Drawing.Size(389, 20)
        Me.DatabaseInfoSelectedTEXTBOX.TabIndex = 1911
        '
        'Label35
        '
        Me.Label35.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label35.BackColor = System.Drawing.Color.BurlyWood
        Me.Label35.Location = New System.Drawing.Point(459, 73)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(113, 2)
        Me.Label35.TabIndex = 1916
        '
        'Label142
        '
        Me.Label142.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label142.BackColor = System.Drawing.Color.BurlyWood
        Me.Label142.Location = New System.Drawing.Point(459, 55)
        Me.Label142.Name = "Label142"
        Me.Label142.Size = New System.Drawing.Size(113, 2)
        Me.Label142.TabIndex = 1917
        '
        'Label147
        '
        Me.Label147.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label147.BackColor = System.Drawing.Color.BurlyWood
        Me.Label147.Location = New System.Drawing.Point(569, 55)
        Me.Label147.Name = "Label147"
        Me.Label147.Size = New System.Drawing.Size(3, 20)
        Me.Label147.TabIndex = 1918
        '
        'Label148
        '
        Me.Label148.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label148.BackColor = System.Drawing.Color.BurlyWood
        Me.Label148.Location = New System.Drawing.Point(458, 55)
        Me.Label148.Name = "Label148"
        Me.Label148.Size = New System.Drawing.Size(3, 20)
        Me.Label148.TabIndex = 1919
        '
        'DatabaseInfoTotalTEXTBOX
        '
        Me.DatabaseInfoTotalTEXTBOX.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoTotalTEXTBOX.BackColor = System.Drawing.Color.Black
        Me.DatabaseInfoTotalTEXTBOX.ForeColor = System.Drawing.Color.White
        Me.DatabaseInfoTotalTEXTBOX.Location = New System.Drawing.Point(458, 55)
        Me.DatabaseInfoTotalTEXTBOX.Name = "DatabaseInfoTotalTEXTBOX"
        Me.DatabaseInfoTotalTEXTBOX.Size = New System.Drawing.Size(113, 20)
        Me.DatabaseInfoTotalTEXTBOX.TabIndex = 1920
        '
        'DatabaseInfoSelectedLABEL
        '
        Me.DatabaseInfoSelectedLABEL.AutoSize = True
        Me.DatabaseInfoSelectedLABEL.BackColor = System.Drawing.Color.Transparent
        Me.DatabaseInfoSelectedLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.DatabaseInfoSelectedLABEL.Location = New System.Drawing.Point(38, 41)
        Me.DatabaseInfoSelectedLABEL.Name = "DatabaseInfoSelectedLABEL"
        Me.DatabaseInfoSelectedLABEL.Size = New System.Drawing.Size(103, 13)
        Me.DatabaseInfoSelectedLABEL.TabIndex = 1921
        Me.DatabaseInfoSelectedLABEL.Text = "Database File Name"
        '
        'DatabaseInfoTotalLABEL
        '
        Me.DatabaseInfoTotalLABEL.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoTotalLABEL.AutoSize = True
        Me.DatabaseInfoTotalLABEL.BackColor = System.Drawing.Color.Transparent
        Me.DatabaseInfoTotalLABEL.ForeColor = System.Drawing.Color.BurlyWood
        Me.DatabaseInfoTotalLABEL.Location = New System.Drawing.Point(455, 41)
        Me.DatabaseInfoTotalLABEL.Name = "DatabaseInfoTotalLABEL"
        Me.DatabaseInfoTotalLABEL.Size = New System.Drawing.Size(59, 13)
        Me.DatabaseInfoTotalLABEL.TabIndex = 1922
        Me.DatabaseInfoTotalLABEL.Text = "Total Items"
        '
        'DatabaseInfoTABCONTROL
        '
        Me.DatabaseInfoTABCONTROL.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoItemBaseTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoItemQualityTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoRunesTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoMulesTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoAccountsTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoRealmsTABPAGE)
        Me.DatabaseInfoTABCONTROL.Controls.Add(Me.DatabaseInfoDatabaseTABPAGE)
        Me.DatabaseInfoTABCONTROL.Location = New System.Drawing.Point(41, 103)
        Me.DatabaseInfoTABCONTROL.Name = "DatabaseInfoTABCONTROL"
        Me.DatabaseInfoTABCONTROL.SelectedIndex = 0
        Me.DatabaseInfoTABCONTROL.Size = New System.Drawing.Size(531, 284)
        Me.DatabaseInfoTABCONTROL.TabIndex = 1923
        '
        'DatabaseInfoItemBaseTABPAGE
        '
        Me.DatabaseInfoItemBaseTABPAGE.Controls.Add(Me.DatabaseInfoDATAGRIDVIEW)
        Me.DatabaseInfoItemBaseTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoItemBaseTABPAGE.Name = "DatabaseInfoItemBaseTABPAGE"
        Me.DatabaseInfoItemBaseTABPAGE.Padding = New System.Windows.Forms.Padding(3)
        Me.DatabaseInfoItemBaseTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoItemBaseTABPAGE.TabIndex = 0
        Me.DatabaseInfoItemBaseTABPAGE.Text = "Item Base"
        Me.DatabaseInfoItemBaseTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoItemQualityTABPAGE
        '
        Me.DatabaseInfoItemQualityTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoItemQualityTABPAGE.Name = "DatabaseInfoItemQualityTABPAGE"
        Me.DatabaseInfoItemQualityTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoItemQualityTABPAGE.TabIndex = 6
        Me.DatabaseInfoItemQualityTABPAGE.Text = "Item Quality"
        Me.DatabaseInfoItemQualityTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoRunesTABPAGE
        '
        Me.DatabaseInfoRunesTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoRunesTABPAGE.Name = "DatabaseInfoRunesTABPAGE"
        Me.DatabaseInfoRunesTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoRunesTABPAGE.TabIndex = 5
        Me.DatabaseInfoRunesTABPAGE.Text = "Runes"
        Me.DatabaseInfoRunesTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoMulesTABPAGE
        '
        Me.DatabaseInfoMulesTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoMulesTABPAGE.Name = "DatabaseInfoMulesTABPAGE"
        Me.DatabaseInfoMulesTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoMulesTABPAGE.TabIndex = 2
        Me.DatabaseInfoMulesTABPAGE.Text = "Mules"
        Me.DatabaseInfoMulesTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoAccountsTABPAGE
        '
        Me.DatabaseInfoAccountsTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoAccountsTABPAGE.Name = "DatabaseInfoAccountsTABPAGE"
        Me.DatabaseInfoAccountsTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoAccountsTABPAGE.TabIndex = 3
        Me.DatabaseInfoAccountsTABPAGE.Text = "Accounts"
        Me.DatabaseInfoAccountsTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoRealmsTABPAGE
        '
        Me.DatabaseInfoRealmsTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoRealmsTABPAGE.Name = "DatabaseInfoRealmsTABPAGE"
        Me.DatabaseInfoRealmsTABPAGE.Padding = New System.Windows.Forms.Padding(3)
        Me.DatabaseInfoRealmsTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoRealmsTABPAGE.TabIndex = 1
        Me.DatabaseInfoRealmsTABPAGE.Text = "Realms"
        Me.DatabaseInfoRealmsTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoDatabaseTABPAGE
        '
        Me.DatabaseInfoDatabaseTABPAGE.Location = New System.Drawing.Point(4, 22)
        Me.DatabaseInfoDatabaseTABPAGE.Name = "DatabaseInfoDatabaseTABPAGE"
        Me.DatabaseInfoDatabaseTABPAGE.Size = New System.Drawing.Size(523, 258)
        Me.DatabaseInfoDatabaseTABPAGE.TabIndex = 4
        Me.DatabaseInfoDatabaseTABPAGE.Text = "Database"
        Me.DatabaseInfoDatabaseTABPAGE.UseVisualStyleBackColor = True
        '
        'DatabaseInfoCloseBUTTON
        '
        Me.DatabaseInfoCloseBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoCloseBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.DatabaseInfoCloseBUTTON.BackgroundImage = CType(resources.GetObject("DatabaseInfoCloseBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.DatabaseInfoCloseBUTTON.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.DatabaseInfoCloseBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.DatabaseInfoCloseBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.DatabaseInfoCloseBUTTON.Location = New System.Drawing.Point(497, 410)
        Me.DatabaseInfoCloseBUTTON.Name = "DatabaseInfoCloseBUTTON"
        Me.DatabaseInfoCloseBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.DatabaseInfoCloseBUTTON.TabIndex = 1924
        Me.DatabaseInfoCloseBUTTON.Text = "Close"
        Me.DatabaseInfoCloseBUTTON.UseVisualStyleBackColor = False
        '
        'DatabaseInfoRefreshBUTTON
        '
        Me.DatabaseInfoRefreshBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoRefreshBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.DatabaseInfoRefreshBUTTON.BackgroundImage = CType(resources.GetObject("DatabaseInfoRefreshBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.DatabaseInfoRefreshBUTTON.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.DatabaseInfoRefreshBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.DatabaseInfoRefreshBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.DatabaseInfoRefreshBUTTON.Location = New System.Drawing.Point(411, 410)
        Me.DatabaseInfoRefreshBUTTON.Name = "DatabaseInfoRefreshBUTTON"
        Me.DatabaseInfoRefreshBUTTON.Size = New System.Drawing.Size(75, 23)
        Me.DatabaseInfoRefreshBUTTON.TabIndex = 1925
        Me.DatabaseInfoRefreshBUTTON.Text = "Refresh"
        Me.DatabaseInfoRefreshBUTTON.UseVisualStyleBackColor = False
        '
        'DatabaseInfoOpenContainingBUTTON
        '
        Me.DatabaseInfoOpenContainingBUTTON.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.DatabaseInfoOpenContainingBUTTON.BackColor = System.Drawing.SystemColors.WindowFrame
        Me.DatabaseInfoOpenContainingBUTTON.BackgroundImage = CType(resources.GetObject("DatabaseInfoOpenContainingBUTTON.BackgroundImage"), System.Drawing.Image)
        Me.DatabaseInfoOpenContainingBUTTON.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.DatabaseInfoOpenContainingBUTTON.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.DatabaseInfoOpenContainingBUTTON.ForeColor = System.Drawing.Color.BurlyWood
        Me.DatabaseInfoOpenContainingBUTTON.Location = New System.Drawing.Point(41, 410)
        Me.DatabaseInfoOpenContainingBUTTON.Name = "DatabaseInfoOpenContainingBUTTON"
        Me.DatabaseInfoOpenContainingBUTTON.Size = New System.Drawing.Size(154, 23)
        Me.DatabaseInfoOpenContainingBUTTON.TabIndex = 1926
        Me.DatabaseInfoOpenContainingBUTTON.Text = "Open Database Directory"
        Me.DatabaseInfoOpenContainingBUTTON.UseVisualStyleBackColor = False
        '
        'DatabaseInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.DiaBase.My.Resources.Resources.BigSettings
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(612, 466)
        Me.ControlBox = False
        Me.Controls.Add(Me.DatabaseInfoOpenContainingBUTTON)
        Me.Controls.Add(Me.DatabaseInfoRefreshBUTTON)
        Me.Controls.Add(Me.DatabaseInfoCloseBUTTON)
        Me.Controls.Add(Me.DatabaseInfoTABCONTROL)
        Me.Controls.Add(Me.DatabaseInfoTotalLABEL)
        Me.Controls.Add(Me.DatabaseInfoSelectedLABEL)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Label142)
        Me.Controls.Add(Me.Label147)
        Me.Controls.Add(Me.Label148)
        Me.Controls.Add(Me.DatabaseInfoTotalTEXTBOX)
        Me.Controls.Add(Me.Label121)
        Me.Controls.Add(Me.Label120)
        Me.Controls.Add(Me.Label122)
        Me.Controls.Add(Me.Label123)
        Me.Controls.Add(Me.DatabaseInfoSelectedTEXTBOX)
        Me.MaximumSize = New System.Drawing.Size(870, 1036)
        Me.MinimumSize = New System.Drawing.Size(440, 320)
        Me.Name = "DatabaseInfo"
        Me.ShowIcon = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Database Information Breakdown"
        CType(Me.DatabaseInfoDATAGRIDVIEW, System.ComponentModel.ISupportInitialize).EndInit()
        Me.DatabaseInfoTABCONTROL.ResumeLayout(False)
        Me.DatabaseInfoItemBaseTABPAGE.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DatabaseInfoDATAGRIDVIEW As DataGridView
    Friend WithEvents Label121 As Label
    Friend WithEvents Label120 As Label
    Friend WithEvents Label122 As Label
    Friend WithEvents Label123 As Label
    Friend WithEvents DatabaseInfoSelectedTEXTBOX As TextBox
    Friend WithEvents Label35 As Label
    Friend WithEvents Label142 As Label
    Friend WithEvents Label147 As Label
    Friend WithEvents Label148 As Label
    Friend WithEvents DatabaseInfoTotalTEXTBOX As TextBox
    Friend WithEvents DatabaseInfoSelectedLABEL As Label
    Friend WithEvents DatabaseInfoTotalLABEL As Label
    Friend WithEvents DatabaseInfoTABCONTROL As TabControl
    Friend WithEvents DatabaseInfoItemBaseTABPAGE As TabPage
    Friend WithEvents DatabaseInfoRealmsTABPAGE As TabPage
    Friend WithEvents DatabaseInfoMulesTABPAGE As TabPage
    Friend WithEvents DatabaseInfoAccountsTABPAGE As TabPage
    Friend WithEvents DatabaseInfoDatabaseTABPAGE As TabPage
    Friend WithEvents DatabaseInfoCloseBUTTON As Button
    Friend WithEvents DatabaseInfoRefreshBUTTON As Button
    Friend WithEvents DatabaseInfoOpenContainingBUTTON As Button
    Friend WithEvents DatabaseInfoItemQualityTABPAGE As TabPage
    Friend WithEvents DatabaseInfoRunesTABPAGE As TabPage
    Friend WithEvents ItemBaseCOLUMN As DataGridViewTextBoxColumn
    Friend WithEvents TotalCOLUMN As DataGridViewTextBoxColumn
    Friend WithEvents RatioCOLUMN As DataGridViewTextBoxColumn
End Class
