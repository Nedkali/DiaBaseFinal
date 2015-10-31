'=========================================================================================================================================================================
'APPLICATION SETTINGS FORM - Handles all adjustmens and saves to all settings variables for the application.
'=========================================================================================================================================================================
Public Class Settings

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SETTINGS PATH CHECKER SUB ROUTINE          - Verifys the Etal and Default Database Paths in Settings Actually exist FOR BOTH THE PUBLIC AND NEDS VERSION NOW
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub CheckSettingsPaths()
        Me.EtalPathVerifyPICTUREBOX.Visible = False
        DatabaseVerifyPICTUREBOX.Visible = False
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(Me.EtalPathTEXTBOX.Text, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = True Or (My.Computer.FileSystem.DirectoryExists(String.Concat(Me.EtalPathTEXTBOX.Text, "\Scripts\AMS\MuleInventory"))) = True Then
            Me.EtalPathVerifyPICTUREBOX.Visible = True
            Me.EtalPathVerifyFailPICTUREBOX.Visible = False

        Else
            Me.EtalPathVerifyPICTUREBOX.Visible = False
            Me.EtalPathVerifyFailPICTUREBOX.Visible = True

        End If

        If My.Computer.FileSystem.FileExists(DefaultDatabaseTEXTBOX.Text) = True Then
            Me.DatabaseVerifyPICTUREBOX.Visible = True
            Me.DatabaseVerifyFailPICTUREBOX.Visible = False

        Else
            Me.DatabaseVerifyPICTUREBOX.Visible = False
            Me.DatabaseVerifyFailPICTUREBOX.Visible = True
        End If


    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SHOW SETTINGS FORM EVENT HANDLER   - Sets central location to display setting window
    '                                   - Apply DiabloII Font to Headers and Buttons
    '                                   - Updates Settings form with config values from global settings vars
    '                                   - Checks for path errors in current log settings and update "ticks" accordingly (may add crosses too for incorrect logs???)
    '                                   - Plays D2Dong Sound
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Settings_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        'Me.Left = Main.Left + 150 : Me.Top = Main.Top + 100         'Set Central window location

        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then
            'Apply DiabloII Fonts - Headers 14 point - Buttons 9 point
            SettingsEtalPathLABEL.Font = New Font(pfc.Families(0), 14, FontStyle.Regular)
            SettingsDatabasePathLABEL.Font = New Font(pfc.Families(0), 14, FontStyle.Regular)

            EtalPathBrowseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            DefaultDatabaseBrowseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SettingsSaveBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SettingsCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

        'Apply Settings Window Control Values and Strings From Global Settings Vars
        EtalPathTEXTBOX.Text = AppSettings.EtalPath                             'Etal Path
        DefaultDatabaseTEXTBOX.Text = AppSettings.DefaultDatabase               'Startup Database
        AutoLogingDelayNUMERICUPDOWN.Value = AppSettings.AutoLoggingDelay       'Autologger delay
        AutoBackupImportsCHECKBOX.Checked = AppSettings.BackupBeforeImports     'Backup before imports bool 
        HideAccountPassCHECKBOX.Checked = AppSettings.HideMulePass
        AutoBackupEditsCHECKBOX.Checked = AppSettings.BackupBeforeEdits         'Backup before item edits bool
        RemoveMuleDupeCHECKBOX.Checked = AppSettings.RemoveMuleDupes            'Remove mule dupe bool
        SoundMuteCHECKBOX.Checked = AppSettings.SoundMute                       'Sound Setting bool
        If AppSettings.DefaultRealm <> "" Then
            SearchRealmCBOX.SelectedItem = AppSettings.DefaultRealm
        Else SearchRealmCBOX.SelectedIndex = 0
        End If
        DefaultPasswordTEXTBOX.Text = AppSettings.DefaultPassword
        ResetDateTEXTBOX.Text = AppSettings.ResetDate
        CheckSettingsPaths()                                        'Verify current paths are correct and update ticks (should only throw to actual error message when exiting with invalid settings)
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'CANCEL BUTTON SETTINGS EVENT HANDLER       - Check current path settings are correct
    '                                           - Close settings window without saving only if current path details are correct
    '                                           -
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub CancelDefaultsBUTTON_Click(sender As Object, e As EventArgs) Handles SettingsCancelBUTTON.Click

        If EtalPathVerifyFailPICTUREBOX.Visible = True Then Main.ErrorHandler(201, 0, 0, 0) 'Check Etal Path Verification flag and Branch to Error Handler If Verify Failed
        If DatabaseVerifyFailPICTUREBOX.Visible = True Then Main.ErrorHandler(202, 0, 0, 0) 'Check Database Path Verification flag and Branch to Error Handler If Verify Failed

        If DatabaseVerifyPICTUREBOX.Visible = True And EtalPathVerifyPICTUREBOX.Visible = True Then             'Exit if both current paths pass verification
            If checkdate() = False Then MessageBox.Show("Date Error") : Return
            Me.Close()
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SAVE SETTINGS BUTTON EVENT HANDLER         - Check Paths have been verified
    '                                           - Update Global Configuration Variables
    '                                           - Save global Variabless to Settings.CFG file if verification has be passed (all pickboxes visible)
    '
    '  NOTE TO ME: May look good with red crosses on incorrect paths, would just need to look for tick instead as a bool or flag for verification result
    '
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SettingsSaveBUTTON_Click(sender As Object, e As EventArgs) Handles SettingsSaveBUTTON.Click

        If EtalPathVerifyFailPICTUREBOX.Visible = True Then Main.ErrorHandler(201, 0, 0, 0) 'Check Etal Path Verification flag and Branch to Error Handler If Verify Failed
        If DatabaseVerifyFailPICTUREBOX.Visible = True Then Main.ErrorHandler(202, 0, 0, 0) 'Check Database Path Verification flag and Branch to Error Handler If Verify Failed
        If checkdate() = False Then MessageBox.Show("Date Error") : Return
        'IF BOTH PATHS PASS VERIFICATION THEN CONTINUE WITH SAVE PROCEEDURE

        If DatabaseVerifyPICTUREBOX.Visible = True And EtalPathVerifyPICTUREBOX.Visible = True Then
            'Update Global Config Vars From Settings Controls
            AppSettings.EtalPath = EtalPathTEXTBOX.Text
            AppSettings.DefaultDatabase = DefaultDatabaseTEXTBOX.Text
            AppSettings.SoundMute = SoundMuteCHECKBOX.CheckState
            AppSettings.RemoveMuleDupes = RemoveMuleDupeCHECKBOX.CheckState
            AppSettings.HideMulePass = HideAccountPassCHECKBOX.CheckState
            AppSettings.BackupBeforeImports = AutoBackupImportsCHECKBOX.CheckState
            AppSettings.BackupBeforeEdits = AutoBackupEditsCHECKBOX.CheckState
            AppSettings.AutoLoggingDelay = AutoLogingDelayNUMERICUPDOWN.Value
            AppSettings.DefaultPassword = DefaultPasswordTEXTBOX.Text
            AppSettings.DefaultRealm = SearchRealmCBOX.Text
            AppSettings.ResetDate = ResetDateTEXTBOX.Text
            AppSettings.DisplayLineBreaks = Main.DisplayLineBreaksMainMenu.CheckState
            SaveSettingsFile()

            'Checks only one Check box is checked across all realm search checkboxes - like radio buttons but not as yucky looking
            If AppSettings.DefaultRealm = "USEast" Then Main.EastRealmCHECKBOX.Checked = True : Main.WestRealmCHECKBOX.Checked = False : Main.AsiaRealmCHECKBOX.Checked = False : Main.EuropeRealmCHECKBOX.Checked = False
            If AppSettings.DefaultRealm = "USWest" Then Main.EastRealmCHECKBOX.Checked = False : Main.WestRealmCHECKBOX.Checked = True : Main.AsiaRealmCHECKBOX.Checked = False : Main.EuropeRealmCHECKBOX.Checked = False
            If AppSettings.DefaultRealm = "Asia" Then Main.EastRealmCHECKBOX.Checked = False : Main.WestRealmCHECKBOX.Checked = False : Main.AsiaRealmCHECKBOX.Checked = True : Main.EuropeRealmCHECKBOX.Checked = False
            If AppSettings.DefaultRealm = "Europe" Then Main.EastRealmCHECKBOX.Checked = False : Main.WestRealmCHECKBOX.Checked = False : Main.AsiaRealmCHECKBOX.Checked = False : Main.EuropeRealmCHECKBOX.Checked = True
            Me.Close()

            'Update etal version info..
            If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "NED"
            If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "PUB"
            If AppSettings.EtalVersion = "NED" Then Main.Text = VersionAndRevision & " - Running Red Dragon Compataibility Mode"
            If AppSettings.EtalVersion = "PUB" Then Main.Text = VersionAndRevision & " - Running Black Empress Compatibility Mode"

        End If
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ETAL FOLDER PATH BROWSE BUTTON HANDLER     - Dialog to browse for the etal installation folder
    '                                           - Set Current Etal Path As Set Selected Path (if there is one)
    '                                           - Apply Dialog result to Settings Textbox
    '                                           - Run Path Check Routine After Input To Update Ticks
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub EtalPathBrowseBUTTON_Click(sender As Object, e As EventArgs) Handles EtalPathBrowseBUTTON.Click
        Main.SelectEtalPathDIALOG.Description = "Please Browse To Your Etal Installaton Folder And Select OK."
        If EtalPathTEXTBOX.Text <> Nothing Then Main.SelectEtalPathDIALOG.SelectedPath = EtalPathTEXTBOX.Text
        Main.SelectEtalPathDIALOG.ShowDialog()
        EtalPathTEXTBOX.Text = Main.SelectEtalPathDIALOG.SelectedPath
        CheckSettingsPaths()
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DEFAULT DATABASE BROWES BUTTON HANDLER     - Apply old Filename as the automatically selected filename in dialog if it is already verified otherwise apply Default.TXT
    '                                           - Set InstallPath\Databases as default directory
    '                                           - Set .TXT and All Files (*.*)as accepted file types for filtering
    '                                           - Apply Dialog Result to Settings Textbox
    '                                           - Run Path Check Routine After Input To Update Ticks
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DefaultDatabaseBrowseBUTTON_Click(sender As Object, e As EventArgs) Handles DefaultDatabaseBrowseBUTTON.Click
        If DatabaseVerifyPICTUREBOX.Visible = True Then
            Main.SelectDefaultDatabaseDIALOG.FileName = My.Computer.FileSystem.GetName(DefaultDatabaseTEXTBOX.Text)
        Else
            Main.SelectDefaultDatabaseDIALOG.FileName = AppSettings.DefaultDatabase
        End If
        Main.SelectDefaultDatabaseDIALOG.Title = "Browse For Database To Open On Startup."
        Main.SelectDefaultDatabaseDIALOG.InitialDirectory = AppSettings.InstallPath + "\Databases"
        Main.SelectDefaultDatabaseDIALOG.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
        Main.SelectDefaultDatabaseDIALOG.FilterIndex = 2
        Main.SelectDefaultDatabaseDIALOG.RestoreDirectory = True
        Main.SelectDefaultDatabaseDIALOG.CheckFileExists = True
        Main.SelectDefaultDatabaseDIALOG.AddExtension = True
        Main.SelectDefaultDatabaseDIALOG.ShowDialog()

        DefaultDatabaseTEXTBOX.Text = Main.SelectDefaultDatabaseDIALOG.FileName




        CheckSettingsPaths()
    End Sub


    Private Sub EtalPathTEXTBOX_TextChanged(sender As Object, e As EventArgs) Handles EtalPathTEXTBOX.TextChanged
        CheckSettingsPaths()

    End Sub

    Private Sub DefaultDatabaseTEXTBOX_TextChanged(sender As Object, e As EventArgs) Handles DefaultDatabaseTEXTBOX.TextChanged
        CheckSettingsPaths()
    End Sub


    Function checkdate()
        Dim temp = ResetDateTEXTBOX.Text.Split("/")
        If temp.Length <> 3 Then Return False ' need to add error handling here
        If temp(0) < 1 Or temp(0) > 31 Then Return False ' need to add error handling here
        If temp(1) < 1 Or temp(1) > 12 Then Return False ' need to add error handling here
        If temp(2) < 2015 Or temp(2) > 2020 Then Return False ' need to add error handling here
        Return True

    End Function





End Class