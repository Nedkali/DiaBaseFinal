'========================================================================================================================================================================
'EXIT APPLICATION FORM CONFIRMATION WINDOW - Is Displayed On The Close Application Event To Confirm Or Cancel - Includes Save And Backup Options Before Confirming Exit
'========================================================================================================================================================================
Public Class ExitApplication

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'EXIT APPLICATION WINDOW LOAD EVENT HANDLER     - Apply Diablo II Fonts To Headers And Buttons
    '                                               - Locate Window in Central Position
    '                                               - Plays Audio D2 Donk When UnMuted
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ExitApplication_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Apply Diablo Font To Header And Buttons
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then
            ExitApplicationHeaderLABEL.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            ExitApplicationCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ExitApplicationConfirmBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

        ExitApplicationSaveDatabaseCHECKBOX.Checked = AppSettings.SaveOnExit
        ExitApplicationBackupDatabaseCHECKBOX.Checked = AppSettings.BackupOnExit
        Me.Left = Main.Left + 200 : Me.Top = Main.Top + 150         'Set Central window location
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)


    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'EXIT APPLICATION CANCEL EXIT HANDLER   - Returns Dialog Result as NO to cancel the Close Appllication Event
    '                                       - Plays Audio D2 Donk When UnMuted
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ExitApplicationCancelBUTTON_Click(sender As Object, e As EventArgs) Handles ExitApplicationCancelBUTTON.Click
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'EXIT APPLICATION CONFIRM EXIT HANDLER  - Returns Dialog Result as YES to continue the Close Appllication Event
    '                                       - Plays Audio D2 Donk When UnMuted
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ExitApplicationConfirmBUTTON_Click(sender As Object, e As EventArgs) Handles ExitApplicationConfirmBUTTON.Click
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        DialogResult = Windows.Forms.DialogResult.Yes
    End Sub


End Class