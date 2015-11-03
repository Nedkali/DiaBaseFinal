'******************************************************************************
'**                   D I A B A S E    V E R S I O N    1 0                  **
'**                                                                          **
'**      W R I T T E N    B Y    A U S S I E H A C K    A N D    N E D       **
'**                                                                          **
'******************************************************************************

Imports System.IO
Imports System.Drawing.Text
Imports System.Globalization
Public Class Main
    'Trial - Setup Select All Function and support vars
    Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Private TriggerUpdate = GetType(ListBox.SelectedObjectCollection).GetField("stateDirty", 32 Or 4)
    Dim TriggerIndexChanged = GetType(ListBox).GetMethod("OnSelectedIndexChanged", 32 Or 4)

    '=====================================================================================================================================================================
    'TIMER ROUTINES
    '=====================================================================================================================================================================
    'DEFINES THE ImportTimer AS A GLOBAL SYSTEM TIMER (ONLY FOR AUTO IMPORTS)
    Public WithEvents ImportTimer As New System.Windows.Forms.Timer()
    Public WithEvents ButtonFlashTimer As New System.Windows.Forms.Timer()


    Sub StartTimer()
        Timercount = 0
        TimerSeconds = AppSettings.AutoLoggingDelay * 60
        ToolStripProgressBar1.Maximum = 100
        ImportTimer.Interval = 1000
        ImportTimer.Start()
        ImportLogRICHTEXTBOX.Text = Date.Today & " @ " & TimeOfDay & " - Session Start, AutoLogger Ready." & vbCrLf
    End Sub


    'Button flashing text effect for when autologger is paused
    Private Sub ButtonTIMER_Tick(sender As Object, e As EventArgs) Handles ButtonTIMER.Tick
        If AutologgingTimerBUTTON.Text <> "Pause AutoLogger" Then ButtonFlashCount = ButtonFlashCount + 1

        If ButtonFlashCount = 1 Then AutologgingTimerBUTTON.Text = Nothing
        If ButtonFlashCount = 10 Then AutologgingTimerBUTTON.Text = "Start AutoLogger"
        If ButtonFlashCount = 20 Then ButtonFlashCount = 0

    End Sub



    'IMPORT DELAY COUNTER - HANDLES IMPORT DLEAY AUTOLOGGING SYSTEM - BRANCES TO IMPORT ROUTINE WHEN DELAY IS UP
    Private Sub TimerEventProcessor(ByVal myObject As Object, ByVal MyEventArgs As EventArgs) Handles ImportTimer.Tick
        Timercount = Timercount + 1
        If Timercount > TimerSeconds Then
            Timercount = 0
            ImportTimer.Stop()
            ImportLogRICHTEXTBOX.AppendText(vbCrLf & "AutoLogger Running - " & Date.Today & " @ " & TimeOfDay & vbCrLf & vbCrLf & "Checking For New Log Files..." & vbCrLf)
            AutoLoggerRunning = True
            ImportLogFiles(False)
            AutoLoggerRunning = False
            ImportTimer.Start()
        End If

        Dim Timerprogress As Integer = Math.Round((Timercount / TimerSeconds) * 100)
        ToolStripProgressBar1.Value = Timerprogress
        ToolStripStatusLabel2.Text = "< " & Math.Ceiling((TimerSeconds - Timercount) / 60) & " Minuites"
    End Sub

    'TIMER BUTTON pause & restart routine
    Private Sub TimerStartPauseButton(sender As Object, e As EventArgs) Handles AutologgingTimerBUTTON.Click
        If AutoLoggerReady = False Then
            ButtonFlashTimer.Start()
            ImportTimer.Stop()
            'TimerRestart = True
            AutoLoggerReady = True
            AutologgingTimerBUTTON.Text = "Start AutoLogger"
            'ImportLogRICHTEXTBOX.AppendText(vbCrLf & vbCrLf & "AutoLogger Paused")
            Return
        End If

        If AutoLoggerReady = True Then
            AutoLoggerReady = False
            ImportTimer.Start()
            ButtonFlashTimer.Stop()
            AutologgingTimerBUTTON.Text = "Pause AutoLogger"
            'ImportLogRICHTEXTBOX.AppendText(vbCrLf & vbCrLf & "AutoLogger Ready")
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'LOAD EVENT HANDLER     - Installs app on first run by setting up required directories and files. Also checks files still exist on each subsequent resart
    '                       - Loads Setting from Settings.CFG file to Settings Global Vars
    '                       - Defines Font Family and Sets Up DiabloII Fancy Font For Headings and Buttons Across All Forms
    '                       - Display Default Database Name In Top Right Corner
    '                       - Loads the default Database into ItemDatabase Class Object and Populates the AllItemsLISTBOX
    '                       - Updates the Hide Dupes Main Form Checkbox From Global Settings Var
    '                       - 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    'MAIN FORM PUBLIC VARIABLES
    'Dim ShowSettingsWindowOnStartup As Boolean = False      'Used to trigger Settings Window On settings file is create Or when Path Verify Fails on startup verification


    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'CREATE ALL DIRECTORYS AND SUPPORT FILES THE APP REQUIRES TO FUNCTION
        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        If My.Computer.FileSystem.DirectoryExists(AppSettings.InstallPath + "\Extras") = False Then My.Computer.FileSystem.CreateDirectory(AppSettings.InstallPath + "\Extras") '                       'Verify Extras Directory
        If My.Computer.FileSystem.DirectoryExists(AppSettings.InstallPath + "\Databases") = False Then My.Computer.FileSystem.CreateDirectory(AppSettings.InstallPath + "\Databases") '                 'Verify Databases Directory
        If My.Computer.FileSystem.DirectoryExists(AppSettings.InstallPath + "\Databases\Backup") = False Then My.Computer.FileSystem.CreateDirectory(AppSettings.InstallPath + "\Databases\Backup") '   'Verify Backup Directory
        If My.Computer.FileSystem.DirectoryExists(AppSettings.InstallPath + "\Archive") = False Then My.Computer.FileSystem.CreateDirectory(AppSettings.InstallPath + "\Archive") '               'Verify Archive Directory

        'Verifys the Default.txt database file
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\Default.txt") = False Then
            Dim WriteFile As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(AppSettings.InstallPath + "\DataBases\Default.txt", False)
            WriteFile.Close()
        End If

        'Creates the Settings.cfg file if not found
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Settings.cfg") = False Then
            SaveSettingsFile()
        End If

        ReadSettingsFile()  'BRANCHES TO ReadSettingFile sub to load config into global config vars 

        'APPLY THE FANCY DIABLO2 FONT TO MAIN FORM HEADERS AND BUTTONS ONLY IF THE DiabloFont1.TTF FONT FILE IS PRESENT IN THE EXTRAS DIRECTORY
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then
            pfc.AddFontFile(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF")

            'Main Form Headers (16 point size)
            ItemAndMuleLABEL.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            AutoLoggingLABEL.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            ListsLABEL.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)
            StatisticsLABEL.Font = New Font(pfc.Families(0), 16, FontStyle.Regular)

            'List Tab Buttons (9 point size)
            ListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SearchListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            UserRefControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            TradesListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            SearchBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            AutologgingTimerBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

        'Set foucus on the main listbox
        AllItemsLISTBOX.Select() : ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Total Items"

        'Play the introduction laugh
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.BigDLaugh, AudioPlayMode.Background)

    End Sub


    '====================================================================================================================================================================
    'SHOWN MAIN FORM EVENT HANDLER  - Checks Paths and Branches to settings if path check fails
    '                               - Loads Default Database File (if there is one)
    '                               - If database is loaded also populates main listbox
    '                               - Focus On 
    '====================================================================================================================================================================
    Private Sub Main_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        'Verify The current Paths In Settings Are Correct - Auto Open Settings Window If Verify Fails - VERIFYS FOR BOTH PUBLIC AND NEDS ETAL VERSIONS NOW
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = False And My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\AMS\MuleInventory")) = False Then Settings.ShowDialog()

        'Verify the default database Path is validated
        If My.Computer.FileSystem.FileExists(AppSettings.DefaultDatabase) = False Then Settings.ShowDialog()

        'Branch to Load Default Database and populate main listbox once all setting proceedures are absolutly completed with all potential path errors handled
        If AppSettings.DefaultDatabase <> Nothing Then OpenDatabase(AppSettings.DefaultDatabase)

        'THIS DETERMINES WHICH ETAL IS CURRENTLY BEING USED (PUBLIC OR NEDS) AND SETS THE NEW AppSettings.EtalVersion var to either "NED" or "PUB" (not to be confused with ned at the pub :)
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "NED"
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "PUB"
        If AppSettings.EtalVersion = "NED" Then Me.Text = VersionAndRevision & " - Running Red Dragon Compataibility Mode"
        If AppSettings.EtalVersion = "PUB" Then Me.Text = VersionAndRevision & " - Running Black Empress Compatibility Mode"

        'Start The Import Timer, Focus The Search Field Textbox And Then Pass Control Back To The Main Form. App Startup is completed..
        StartTimer()
        SearchFieldCOMBOBOX.Select()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MAIN LISTBOX POPUP MENU - ON RIGHT MOUSE CLICK
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub AllItemsLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles AllItemsLISTBOX.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If AllItemsLISTBOX.SelectedIndex > -1 Then
                If Not String.IsNullOrEmpty(AllItemsLISTBOX.Text) Then
                    Me.ItemListboxCONTEXTMENUSTRIP.Show(Control.MousePosition)
                End If
            End If
        Else
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MAIN LISTBOX INDEX CHANGED HANDLER         - Branches to Display Item Statistics Routine
    '                                           - Tallys items or selected items depending on the number of items selected
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub AllItemsInDatabaseListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AllItemsLISTBOX.SelectedIndexChanged
        If AllItemsLISTBOX.Focused = False Then Return
        Dim a As Integer = AllItemsLISTBOX.SelectedIndex
        If a > -1 Then
            AllItemsLISTBOX.SelectedIndex = a
            DisplayItemStats(a)
            If AllItemsLISTBOX.SelectedItems.Count > 1 Then ItemTallyTEXTBOX.Text = AllItemsLISTBOX.SelectedItems.Count & " - Selected" : Return
        End If
        ItemTallyTEXTBOX.Text = ItemObjects.Count & " - Items"
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SEARCH LISTBOX INDEX CHANGED HANDLER       - Focuses Main Listbox Selected Item to Match the Serch Lists Currently Selected Item 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SearchLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchLISTBOX.SelectedIndexChanged
        'this needs more work to be have more functionality eg multi selected items
        Dim a As Integer = SearchLISTBOX.SelectedIndex
        If a > -1 Then
            AllItemsLISTBOX.SelectedItems.Clear() 'only selected item will be highlighted in both listings
            AllItemsLISTBOX.SelectedIndex = SearchReferenceList(a)
            DisplayItemStats(SearchReferenceList(a))
            If SearchLISTBOX.SelectedItems.Count > 1 Then ItemTallyTEXTBOX.Text = SearchLISTBOX.SelectedItems.Count & " - Selected" : Return
        End If
        ItemTallyTEXTBOX.Text = SearchReferenceList.Count & " - Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'USER LISTBOX INDEX CHANGED HANDLER         - Branches to Display Item Statistics Routine
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------




    'Sets the UserListbox.SelectedItem the same as the AllItemsInDatabaseListbox.SelectedItems to trigger the items stats to be displayed
    Private Sub UserLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles UserLISTBOX.SelectedIndexChanged

        'Sets User Items as TallyLabel value when only 1 item is selected else it sets TallyLabel value as SelectedItems.count
        If UserLISTBOX.SelectedItems.Count <= 1 Then
            ItemTallyTEXTBOX.Text = (UserLISTBOX.Items.Count & " - User")
        ElseIf UserLISTBOX.SelectedItems.Count > 1 Then
            ItemTallyTEXTBOX.Text = (UserLISTBOX.SelectedItems.Count & " - Selected")
        End If


        'enables or disables the open containing database function is user list contet menu 
        If UserLISTBOX.SelectedItems.Count = 1 Then
            If UserObjects(UserLISTBOX.SelectedIndex).DatabaseFilename <> DatabaseFileNameTEXTBOX.Text Then OpenContainingDatabaseToolStripMenuItem.Enabled = True Else OpenContainingDatabaseToolStripMenuItem.Enabled = False
        Else
            OpenContainingDatabaseToolStripMenuItem.Enabled = False
        End If

        'Branch to UserFunction module to run sub to display stats for selected user list item
        UserListFunctions.DisplaySelectedUserListItem()

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ALL ITEMS LISTBOX TAB BUTTON HANDLER - Selects Button And Displays the AllItemsLISTBOX
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ListControlTabBUTTON_Click(sender As Object, e As EventArgs) Handles ListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(0)
        ListControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        SearchListControlTabBUTTON.BackgroundImage = Nothing
        TradesListControlTabBUTTON.BackgroundImage = Nothing
        UserRefControlTabBUTTON.BackgroundImage = Nothing
        ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Items"
        DatabaseFileLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()

        'THIS CHECKS IF AN ITEM IS SELECTED IF NOT IT SELECTS THE FIRST IN THE LIST
        'IF AN ITEM IS ALREADY SELECTED THEN IT WILL TRIGGER THE STATS TO REFRESH FOR THAT ITEM - THIS KEEPS THE STATS DISPLAYED ALWAYS RELEVANT TO THE CURRENTLY SELECTED LIST... SORT OF THING
        If AllItemsLISTBOX.SelectedIndex = -1 And AllItemsLISTBOX.Items.Count > 0 Then AllItemsLISTBOX.SelectedIndex = 1 Else If AllItemsLISTBOX.SelectedIndex > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex)
        If AllItemsLISTBOX.Items.Count = 0 Then ClearItemStats()

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'USER LISTBOX TAB BUTTON HANDLER - Selects Button And Displays the UserLISTBOX
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub UserRefControlTabBUTTON_Click(sender As Object, e As EventArgs) Handles UserRefControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(3)
        SearchListControlTabBUTTON.BackgroundImage = Nothing
        ListControlTabBUTTON.BackgroundImage = Nothing
        TradesListControlTabBUTTON.BackgroundImage = Nothing
        UserRefControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        ItemTallyTEXTBOX.Text = (UserLISTBOX.Items.Count & " - User")
        DatabaseFileLABEL.Show()
        DatabaseFileNameTEXTBOX.Show()

        'THIS CHECKS IF AN ITEM IS SELECTED IF NOT IT SELECTS THE FIRST IN THE LIST
        'IF AN ITEM IS ALREADY SELECTED THEN IT WILL TRIGGER THE STATS TO REFRESH FOR THAT ITEM - THIS KEEPS THE STATS DISPLAYED ALWAYS RELEVANT TO THE CURRENTLY SELECTED LIST... SORT OF THING
        If UserLISTBOX.SelectedIndex = -1 And UserLISTBOX.Items.Count > 0 Then UserLISTBOX.SelectedIndex = 1 Else If UserLISTBOX.SelectedIndex > -1 Then UserListFunctions.DisplaySelectedUserListItem()
        If UserLISTBOX.Items.Count = 0 Then ClearItemStats()

    End Sub



    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SEARCH LISTBOX TAB BUTTON HANDLER - Selects Button And Displays the SearchLISTBOX
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SearchListControlTabBUTTON_Click(sender As Object, e As EventArgs) Handles SearchListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(1)
        SearchListControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        ListControlTabBUTTON.BackgroundImage = Nothing
        TradesListControlTabBUTTON.BackgroundImage = Nothing
        UserRefControlTabBUTTON.BackgroundImage = Nothing
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Matches"
        DatabaseFileLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()

        If SearchLISTBOX.SelectedIndex = -1 And SearchLISTBOX.Items.Count > 0 Then SearchLISTBOX.SelectedIndex = 1 Else If SearchLISTBOX.SelectedIndex > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex)
        If SearchLISTBOX.Items.Count = 0 Then ClearItemStats()

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ROUTINE CLEANS OUT THE ITEM STATS CONTROLS FOR WHEN AN ITEM IS NOT SELECTED OR WHEN NO ITEMS EXITS IN A LIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub ClearItemStats()
        Me.MuleRealmTEXTBOX.Clear()
        Me.MuleAccountTEXTBOX.Clear()
        Me.MuleNameTEXTBOX.Clear()
        Me.MulePasswordTEXTBOX.Clear()
        Me.CoreTypeTEXTBOX.Clear()
        Me.LadderTEXTBOX.Clear()
        Me.ItemStatsRICHTEXTBOX.Clear()
        Me.ItemNameRICHTEXTBOX.Clear()

        Me.ItemSkinPICTUREBOX.Image = Nothing
        Me.OpenDatabaseLABEL.Text = ""

    End Sub



    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'TRADE LIST LISTBOX TAB BUTTON HANDLER - Selects Button And Displays the TradeListLISTBOX
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub TradesListControlTabBUTTON_Click(sender As Object, e As EventArgs) Handles TradesListControlTabBUTTON.Click
        ListboxTABCONTROL.SelectTab(2)
        TradesListControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        SearchListControlTabBUTTON.BackgroundImage = Nothing
        ListControlTabBUTTON.BackgroundImage = Nothing
        UserRefControlTabBUTTON.BackgroundImage = Nothing
        DatabaseFileLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()
        Dim TradeItemCounter As Integer = 0
        For Each item In TradeListRICHTEXTBOX.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Entries"
    End Sub




    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DISPLAY SETTINGS FORM - Settings Window Handles All Global Config Functions for the Entire Application
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SettingsMainMenu_Click(sender As Object, e As EventArgs) Handles SettingsMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        Settings.ShowDialog()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        TimerStartPauseButton(sender, e)

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - DISPLAY DATABASE MANAGER FORM - Handles Open Create Delete Rename Database files all in one place for (hopefully an improvment)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub OpenDBManagerMainMenu_Click(sender As Object, e As EventArgs) Handles DatabaseManagerMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        DatabaseManager.ShowDialog()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        TimerStartPauseButton(sender, e)

    End Sub


    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - EXIT APP FUNCTION - Closes App On Demand    - Is Thrown only after the exit Event begins
    '                                                       - Shows The Exit Confirmation Windows to display a confirmation and save and or backup message
    '                                                       - Eiter Cancels The Exit Application Event OR Allows it to Continue Depending on the Yes / No Dialog Returned
    '                                                       - Save and Backup Options Branch From The Exit Application Code Before runtime Returns Here
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If AutoLoggerRunning = False Then ' Check The Autologger Is Inactive Before Displaying Confirmation Dialog

            ExitApplication.ShowDialog()
            If ExitApplication.DialogResult = Windows.Forms.DialogResult.No Then e.Cancel = True : Return
            If ExitApplication.DialogResult = Windows.Forms.DialogResult.Yes Then

                'Check The Automated Backup On Exit Checkbox And Branch To Backup Sub If Nessicary
                If ExitApplication.ExitApplicationBackupDatabaseCHRCKBOX.Checked = True Then CreateBackup(AppSettings.CurrentDatabase)

                'Check The Automated Save On Exit Checkbox - used to save overwrite whole database
                If ExitApplication.ExitApplicationSaveDatabaseCHECKBOX.Checked = True Then WriteToFile(0, AppSettings.CurrentDatabase, False)

                'Updates Realm Checkboxes and Display Line Breaks In Stats CheckStates To Settings File
                SaveSettingsFile()

            End If
        Else
            e.Cancel = True 'Automatically cancels Exit Event If Autologger Is Running (Avoids Potential Import Errors)
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR- CLOSE APPLICATION HANDLER - Shuts down main form
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ExitMainMenu_Click(sender As Object, e As EventArgs) Handles ExitApplicarionMENUITEM.Click
        Me.Close()
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - IMPORT NOW FUNCTION - Activated the autologger on demand as opposed to waiting for the delay to time out
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ImportNowMainMenu_Click_1(sender As Object, e As EventArgs) Handles ImportNowMENUITEM.Click

        AutoLoggerRunning = True
        Timercount = 0
        ImportLogRICHTEXTBOX.AppendText(vbCrLf & "AutoLogger Running - " & Date.Today & " @ " & TimeOfDay & vbCrLf & vbCrLf & "Checking For New Log Files..." & vbCrLf)
        ImportLogFiles(False)
        AutoLoggerRunning = False
        If AllItemsLISTBOX.Items.Count > 0 Then AllItemsLISTBOX.SelectedIndex = 0
        ListboxTABCONTROL.SelectTab(0)
        ItemTallyTEXTBOX.Text = ItemObjects.Count & " - Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - SAVE DATABASE FUNCTION - branches to save routine to save the current database 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SaveMainMenu_Click(sender As Object, e As EventArgs) Handles SaveDatabaseMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        WriteToFile(0, AppSettings.DefaultDatabase, False)
        TimerStartPauseButton(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CREATE BACKUP BUTTONPRESS HANDLER - Branches to backup routine to create backup of current file 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub BackupMainMenu_Click(sender As Object, e As EventArgs) Handles BackupDatabaseMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)

        'Setup User Input Form For Use With The Menu BarBackup Current Database Function
        UserInput.Text = "Database Backup System"
        UserInput.UserInputHeaderLABEL.Text = "BACKUP CURRENT DATABASE"

        UserInput.UserInputTEXTBOX.Text = Nothing
        UserInput.UserInputMessageLABEL.Text = "Please select BACKUP to create a backup file for the " & Chr(34) & Me.OpenDatabaseLABEL.Text & Chr(34) & " database file." + vbCrLf + vbCrLf + "Once created, select DATABASE then RESTORE BACKUP from the Menu Bar to rebuild the current database from backup."
        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "Backup"

        UserInput.DatabaseManagerBorder1LABEL.Visible = False
        UserInput.DatabaseManagerBorder2LABEL.Visible = False
        UserInput.DatabaseManagerBorder3LABEL.Visible = False
        UserInput.DatabaseManagerBorder4LABEL.Visible = False
        UserInput.UserInputTEXTBOX.Visible = False

        UserInput.UserInputTEXTBOX.SelectionStart = 0 : UserInput.UserInputTEXTBOX.SelectionLength = Len(UserInput.UserInputTEXTBOX.Text)
        UserInput.UserInputTEXTBOX.Select()

        'Confirms the backup proceedure and proceeds on confirmation
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Dim DialogResult = UserInput.ShowDialog
        If DialogResult = Windows.Forms.DialogResult.Yes Then
            CreateBackup(AppSettings.CurrentDatabase)
            If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)

        End If
        TimerStartPauseButton(sender, e)

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - RESTORE FROM BACKUP BUTTON PRESS HANDLER - Branches to backup routine to restore the current file from backup 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub RestoreBackupMainMenu_Click(sender As Object, e As EventArgs) Handles RestoreBackupMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        'Setup User Input Form For Use With The Menu Bar Restore Backup Function
        UserInput.Text = "Database Backup System"
        UserInput.UserInputHeaderLABEL.Text = "RESTORE DATABASE FROM BACKUP"

        UserInput.UserInputTEXTBOX.Text = Nothing
        UserInput.UserInputMessageLABEL.Text = "Please select RESTORE to rebuild the " & Chr(34) & Me.OpenDatabaseLABEL.Text & Chr(34) & " database from its associated backup file."
        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "Restore"

        UserInput.DatabaseManagerBorder1LABEL.Visible = False
        UserInput.DatabaseManagerBorder2LABEL.Visible = False
        UserInput.DatabaseManagerBorder3LABEL.Visible = False
        UserInput.DatabaseManagerBorder4LABEL.Visible = False
        UserInput.UserInputTEXTBOX.Visible = False

        UserInput.UserInputTEXTBOX.SelectionStart = 0 : UserInput.UserInputTEXTBOX.SelectionLength = Len(UserInput.UserInputTEXTBOX.Text)
        UserInput.UserInputTEXTBOX.Select()

        'Confirms the backup proceedure and proceeds on confirmation
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Dim DialogResult = UserInput.ShowDialog
        If DialogResult = Windows.Forms.DialogResult.Yes Then

            If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\Backup\" + Me.OpenDatabaseLABEL.Text + ".bak") = True Then

                'delete current database file
                My.Computer.FileSystem.DeleteFile(AppSettings.InstallPath + "\Databases\" + Me.OpenDatabaseLABEL.Text + ".txt")

                'Copy over replacement from backup directory
                My.Computer.FileSystem.CopyFile(AppSettings.InstallPath + "\Databases\Backup\" + Me.OpenDatabaseLABEL.Text + ".bak", AppSettings.InstallPath + "\Databases\" + Me.OpenDatabaseLABEL.Text + ".txt")

                OpenDatabase(AppSettings.InstallPath + "\Databases\" + Me.OpenDatabaseLABEL.Text + ".txt")

            Else
                ErrorHandler(901, 0, 0, 0) 'Throw to handler if no database file exists
            End If
            If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) 'confirmatory sound bite
        End If
        TimerStartPauseButton(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - EDIT ITEM FUNCTION - Displays edit item form and broances to EditItem.vb if 1 or more items are selected in the main listbox
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub EditExistingItemMainMenu_Click(sender As Object, e As EventArgs) Handles EditExistingItemMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        If AllItemsLISTBOX.SelectedIndex = -1 Then
            ErrorHandler(2, 0, 0, 0) : Return
        End If
        iEdit = AllItemsLISTBOX.SelectedIndex
        TimerStartPauseButton(sender, e)
        ItemEdit.ShowDialog()
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Sort ItemObjects alphabetically
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SortItemsMainMenu_Click(sender As Object, e As EventArgs) Handles SortItemsMainMenu.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        ItemTallyTEXTBOX.Text = ("Sorting A to Z)")
        ItemObjects.Sort(Function(x, y) x.ItemName.CompareTo(y.ItemName))
        PopulateAllItemsLISTBOX()
        TimerStartPauseButton(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - 'SENDS HIGHLIGHTED ITEMS TO THE TRADE LIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SendToTradeListToolStripMenuItem3_Click(sender As Object, e As EventArgs)

        If AutoLoggerRunning = True Then
            ImportLogRICHTEXTBOX.AppendText("Please wait Import in progress")
            Return
        End If
        ImportTimer.Stop()


        If AllItemsLISTBOX.SelectedIndices.Count > 0 Then
            Dim a As Integer = 0
            Dim count As Integer = 0
            'DupeCountProgressForm.Show() : DupeCountProgressForm.DupePROGRESSBAR.Value = 0 'show and reset progress bar
            For index = 0 To AllItemsLISTBOX.SelectedIndices.Count - 1

                'CALCULATE PROGRESS BAR
                'DupeCountProgressForm.DupePROGRESSBAR.Value = Int((count / AllItemsLISTBOX.Items.Count) * 100)
                count = count + 1
                a = AllItemsLISTBOX.SelectedIndices(index)

                Dim Temp = ItemObjects(a).ItemName
                If ItemObjects(a).ItemBase = "Rune" Or ItemObjects(a).ItemBase = "Gem" Or ItemObjects(a).ItemName.IndexOf("Token") > -1 Or ItemObjects(a).ItemName.IndexOf("Key of") > -1 Or ItemObjects(a).ItemName.IndexOf("Essence") > -1 Then
                    If ItemObjects(a).ItemName.IndexOf("Token") > -1 Then Temp = "Token"
                    TradeListRICHTEXTBOX.AppendText(Temp & vbCrLf & vbCrLf)

                Else
                    SendToTradeList(a)
                End If
            Next
            AllItemsLISTBOX.SelectedIndex = -1
        End If

        'DupeCountProgressForm.Close()

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB

        ListboxTABCONTROL.SelectTab(2)
        TradesListControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        SearchListControlTabBUTTON.BackgroundImage = Nothing
        ListControlTabBUTTON.BackgroundImage = Nothing
        UserRefControlTabBUTTON.BackgroundImage = Nothing

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In TradeListRICHTEXTBOX.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Entries"
        ImportTimer.Start()

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - SELLECT ALL ITEMS
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SelectAllMainMenu_Click(sender As Object, e As EventArgs) Handles SelectAllMENUITEM.Click
        SendMessage(AllItemsLISTBOX.Handle, &H185, New IntPtr(1), New IntPtr(-1))
        TriggerUpdate.SetValue(AllItemsLISTBOX.SelectedItems, True)
        TriggerIndexChanged.Invoke(AllItemsLISTBOX, New Object() {New EventArgs})

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - OPENS ADD NEW ITEM FORM
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub AddNewItemMainMenu_Click(sender As Object, e As EventArgs) Handles AddNewItemMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        ItemAdd.ShowDialog()
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - DELETE SELECTED ITEM/S
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DeleteItemMainMenu_Click(sender As Object, e As EventArgs) Handles DeleteItemMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)

        Dim FocusOnExit As Integer = AllItemsLISTBOX.SelectedIndex

        'check for backup on edits set to true if so then backup now
        If AppSettings.BackupBeforeEdits = True Then CreateBackup(AppSettings.CurrentDatabase)
        If AllItemsLISTBOX.SelectedIndices.Count > 0 Then
            UnDoCount.Add(AllItemsLISTBOX.SelectedIndices.Count - 1)
            For index = AllItemsLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                Dim a As Integer = AllItemsLISTBOX.SelectedIndices(index)
                AllItemsLISTBOX.Items.RemoveAt(a)
                UnDo.Add(ItemObjects(a)) : UnDoPos.Add(a) : UndoDeleteMENUITEM.Enabled = True
                ItemObjects.RemoveAt(a)

                'Removes Item From Search Listbox & Ref List
                Dim count As Integer = SearchReferenceList.Count - 1
                For Each item In SearchReferenceList

                    If SearchReferenceList(count) = a Then
                        SearchReferenceList.RemoveAt(count)
                        SearchLISTBOX.Items.RemoveAt(count)

                        For x = a To SearchReferenceList.Count - 1
                            SearchReferenceList(x) = SearchReferenceList(x) - 1
                        Next
                        Exit For
                    End If
                    count = count - 1
                Next
            Next
            ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Items"

            'SET THE DELETED OBJECT LOCATION IN THE LIST AS THE HIGHLIGHTED ITEM ON RETURN FROM DETETE
            If FocusOnExit >= (AllItemsLISTBOX.Items.Count) Then FocusOnExit = AllItemsLISTBOX.Items.Count - 1
            If AllItemsLISTBOX.Items.Count = 1 Then FocusOnExit = 0
            If AllItemsLISTBOX.Items.Count = 0 Then FocusOnExit = -1
            AllItemsLISTBOX.SelectedIndex = FocusOnExit

        End If

        TimerStartPauseButton(sender, e)

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEAR SEARCH LISTING
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearSearchListMainMenu_Click(sender As Object, e As EventArgs) Handles ClearSearchListMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        SearchLISTBOX.Items.Clear()
        SearchReferenceList.Clear()
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Items"
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEARS TRADELIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearTradeListMainMenu_Click(sender As Object, e As EventArgs) Handles ClearTradeListMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        TradeListRICHTEXTBOX.Clear()
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEARS USERLIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearUserListMainMenu_Click(sender As Object, e As EventArgs) Handles ClearUserListMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        UserLISTBOX.Items.Clear()
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Add blank line - Line Breaks to item stats display
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DisplayLineBreaksMainMenu_Click_1(sender As Object, e As EventArgs) Handles DisplayLineBreaksMENUITEM.Click

        'Update LineBreaks settings to false
        If DisplayLineBreaksMENUITEM.Checked = True Then
            DisplayLineBreaksMENUITEM.Checked = False

            'Refresh The Stats Display to apply the False Setting changes relevant to the list selected
            If ItemTallyTEXTBOX.Text.IndexOf("Items") > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex) '     If All Items List Is Selected
            If ItemTallyTEXTBOX.Text.IndexOf("Matches") > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex) '   If Search List Is Selected
            If ItemTallyTEXTBOX.Text.IndexOf("User") > -1 Then UserListFunctions.DisplaySelectedUserListItem() '      If User List Is Selected
            Return
        End If

        'Update LineBreaks settings to true
        If DisplayLineBreaksMENUITEM.Checked = False Then
            DisplayLineBreaksMENUITEM.Checked = True

            'Refresh The Stats Display to apply ther True setting changes relevant to the lists selected
            If ItemTallyTEXTBOX.Text.IndexOf("Items") > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex) '     If All Items List Is Selected
            If ItemTallyTEXTBOX.Text.IndexOf("Matches") > -1 Then DisplayItemStats(AllItemsLISTBOX.SelectedIndex) '   If Search List Is Selected
            If ItemTallyTEXTBOX.Text.IndexOf("User") > -1 Then UserListFunctions.DisplaySelectedUserListItem() '      If User List Is Selected
            Return
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Rebuild database from Archived logs
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub RebuildDBaseMainMenu_Click(sender As Object, e As EventArgs) Handles RebuildDefaultDatabaseMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)
        ImportLogRICHTEXTBOX.Text = "Checking for New Logs" & vbCrLf
        ImportLogFiles(True)
        PopulateAllItemsLISTBOX()
        TimerStartPauseButton(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Hide dupes in search results
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub HideDupesMainMenu_Click(sender As Object, e As EventArgs)
        If HideDupesMENUITEM.Checked = True Then
            HideDupesMENUITEM.Checked = False
            AppSettings.HideDupes = False
            Return
        End If
        If HideDupesMENUITEM.Checked = False Then
            HideDupesMENUITEM.Checked = True
            AppSettings.HideDupes = True
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Undo delete - works but restores one at a time atm
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub UndoDeleteMainMenu_Click(sender As Object, e As EventArgs) Handles UndoDeleteMENUITEM.Click
        Dim index = CInt(UnDoCount(UnDoCount.Count - 1))
        If UnDo.Count > 0 Then
            For index = CInt(UnDoCount(UnDoCount.Count - 1)) To 0 Step -1
                Dim a As Integer = UnDo.Count - 1
                Dim b As Integer = CInt(UnDoPos(a))
                ItemObjects.Insert(b, UnDo(a))
                UnDo.RemoveAt(a) : UnDoPos.RemoveAt(a)
                AllItemsLISTBOX.Items.Insert(b, ItemObjects(b).ItemName)
            Next

            UnDoCount.RemoveAt(UnDoCount.Count - 1)

        End If
        If UnDo.Count < 1 Then
            UndoDeleteMENUITEM.Enabled = False
        End If
        ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'All Item List MENU - rerouted to main menu option
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SortListItemsCMenu_Click(sender As Object, e As EventArgs) Handles SortListItemsCMenu.Click
        SortItemsMainMenu_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub DeleteItemsCMenu_Click(sender As Object, e As EventArgs) Handles DeleteItemsCMenu.Click
        DeleteItemMainMenu_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub EditItemItemsCMenu_Click(sender As Object, e As EventArgs) Handles EditItemItemsCMenu.Click
        EditExistingItemMainMenu_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub SendToTradeListItemsCMenu_Click(sender As Object, e As EventArgs) Handles SendToTradeListItemsCMenu.Click
        SendToTradeListToolStripMenuItem3_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub SelectAllItemsCMenu_Click(sender As Object, e As EventArgs) Handles SelectAllItemsCMenu.Click
        SelectAllMainMenu_Click(sender, e) 'links to main menu option
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SEARCH BUTTON - Branches to search routine if a realm checkbox is checked else aborts
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SearchBUTTON_Click(sender As Object, e As EventArgs) Handles SearchBUTTON.Click

        If EastRealmCHECKBOX.Checked = False And AsiaRealmCHECKBOX.Checked = False And WestRealmCHECKBOX.Checked = False And EuropeRealmCHECKBOX.Checked = False Then Return
        SearchRoutine()
    End Sub


    '=====================================================================================================================================================================
    'ERROR TRAP EVENT HANDLER - Handles all errors for application. If Possible all trapping handlers should terminate here.
    '=====================================================================================================================================================================
    Sub ErrorHandler(ErrorCode, SystemCode, ParameterOne, ParameterTwo)

        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.BaalLaugh, AudioPlayMode.Background)

        Select Case ErrorCode

            Case 1 'Handles autologger in progress
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.None
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "ALERT!!!"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = vbCrLf & vbCrLf & vbTab & "    Autlogging in Progress = Please wait"
                ErrorHandlerForm.ShowDialog()

            Case 2 'Handles NO ITEM SELECTED FOR EDITING
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.None
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "ALERT!!!"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = vbCrLf & vbCrLf & vbTab & "    NO ITEM SELECTED"
                ErrorHandlerForm.ShowDialog()


            Case 100 'FROM MAIN FORM LOAD HANDLER   - SETTINGS FILE NOT FOUND -
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.None
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SETTINGS FILE NOT FOUND"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "The application cannot loacate the Settings.CFG file." & vbCrLf & vbCrLf & "A new default Settings file will need to be created."
                ErrorHandlerForm.ShowDialog()

                If ErrorHandlerForm.DialogResult = Windows.Forms.DialogResult.No Then
                    SaveSettingsFile()
                    'ShowSettingsWindowOnStartup = True
                End If

            Case 101 'FROM MAIN FORM LOAD HANDLER   - READ SETTINGS FILE UNSPECIFIED SYSTEM ERROR -
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.None
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR LOADING SETTINGS"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "The application cannot load the Settings.CFG file and must terminate." & vbCrLf & vbCrLf & SystemCode.Message
                ErrorHandlerForm.ShowDialog()
                End

            Case 201 'FROM SETTINGS WINDOW SAVE AND CANCEL BUTTON HANDLERS      - ETAL PATH IN SETTINGS IS INCORRECT ERROR -
                MessageBox.Show("FROM SETTINGS WINDOW SAVE AND CANCEL BUTTON HANDLERS      - ETAL PATH IN SETTINGS IS INCORRECT ERROR -")
            'End

            Case 202 'FROM SETTINGS WINDOW SAVE AND CANCEL BUTTON HANDLERS      - DEFAULT DATABASE PATH IN SETTINGS IS INCORRECT ERROR -
                MessageBox.Show("FROM SETTINGS WINDOW SAVE AND CANCEL BUTTON HANDLERS      - DEFAULT DATABASE PATH IN SETTINGS IS INCORRECT ERROR -")
            'End

            Case 301 'FROM SETTINGS SAVE CONFIG FILE ROUTINE IN SAVE BUTTON HANDLER     - ERROR WHILE TRYING TO UPDATE SETTINGS FILE -
                MessageBox.Show("FROM SETTINGS SAVE CONFIG FILE ROUTINE IN SAVE BUTTON HANDLER     - ERROR WHILE TRYING TO UPDATE SETTINGS FILE -")
                End

            Case 401 'FROM DATABASE MANAGER FORM - OPEN DATABASE ROUTINE - SELECTED FILE DOES NOT EXIST
                MessageBox.Show("FROM DATABASE MANAGER FORM - OPEN DATABASE ROUTINE - SELECTED FILE DOES NOT EXIST -")
                End

            Case 402 'FROM DATABASE MANAGER FORM - OPEN DATABASE ROUTINE - SELECTED FILE IS NOT A VALID DIABASE FILE
                MessageBox.Show("FROM DATABASE MANAGER FORM - OPEN DATABASE ROUTINE - SELECTED FILE IS NOT A VALID DIABASE FILE -")
                End


            Case 501 'FROM OPEN DATABASE ROUTINE - Unexpected System Error
                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR OPENING DATABASE"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Text = "Full Report"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "The application has encountered a problem while opening the " & Chr(34) & ParameterTwo & Chr(34) & " database file." & vbCrLf & vbCrLf & "The error seems to have occured while reading record " & ParameterOne & vbCrLf & vbCrLf & SystemCode.Message & vbCrLf & vbCrLf & "You will need to restore this database from backup and re-import any logs imported since last save."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()

                'Show full error report - on button press (uses same form just appends new info to textbox.)
                If DialogResult = Windows.Forms.DialogResult.Yes Then
                    ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.Vertical
                    ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                    ErrorHandlerForm.ErrorTrapMessageTEXTBOX.AppendText(vbCrLf & vbCrLf & vbCrLf & "FULL SYSTEM ERROR REPORT..." & vbCrLf & vbCrLf & SystemCode.ToString)
                    ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                    DialogResult = ErrorHandlerForm.ShowDialog()
                    ErrorHandlerForm.ErrorTrapMessageTEXTBOX.ScrollBars = ScrollBars.None
                End If
                DatabaseManagmentFunctions.CloseFile() ' Close and clear file
            '-----------------------------------------------------------------------------------------



            Case 502 'FROM OPEN DATABASE ROUTINE - Read File Is Out Of Sync and May Be Corrupted
                MessageBox.Show("FROM OPEN DATABASE ROUTINE - DATABASE FILE BEING OPENED IS OUT OF SYNC - MAY BE CORRUPTED AND NEED A BACKUP RESTORE -")
                End


            Case 601 ' DATABASE MANAGEMENT FUNCTIONS MODULE - Unexpected Error Saving Database file

                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR SAVING DATABASE"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Text = "Full Report"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "DiaBase has encountered a problem while saving the " & Chr(34) & Me.OpenDatabaseLABEL.Text & Chr(34) & " database file." & vbCrLf & vbCrLf & "The error seems to have occured while writing record " & ParameterOne & vbCrLf & vbCrLf & SystemCode.Message & vbCrLf & vbCrLf & "You may need to re-import any logs you have added to this database since last save."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()

            Case 701 ' DATABASE MANAGEMENT FUNCTIONS MODULE - Unexpected Error Renaming Database file

                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR RENAMING DATABASE"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Text = "Full Report"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "DiaBase has encountered a problem while Renaming the " & Chr(34) & ParameterTwo & Chr(34) & " database file." & vbCrLf & vbCrLf & SystemCode.Message & vbCrLf & vbCrLf & "Please report this error if the problem persists."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()

            Case 801 ' DATABASE MANAGEMENT FUNCTIONS MODULE - Unexpected Error Deleting Database file

                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR DELETING DATABASE"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Text = "Full Report"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "DiaBase has encountered a problem while Deleting the " & Chr(34) & ParameterTwo & Chr(34) & " database file." & vbCrLf & vbCrLf & SystemCode.Message & vbCrLf & vbCrLf & "Please report this error if the problem persists."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()

            Case 901 'RESTORE FROM BACKUP ROUTINE - MAIN FORM - Backup file does not exist for the current database - cannot restore from baxckup

                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "SYSTEM ERROR RESTORING DATABASE"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Confirm"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "DiaBase cannot locate the backup file for the current database. Be sure to check the backup checkboxes in settings to auto backup before most functions to ensure a viable backup file always exists." & vbCrLf & vbCrLf & "Restore current database from backup has been aborted."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()

            Case 1001 'IMPORT ITEMS TO DATABASE - EXPORT FORM - Unexpected Error - Item Export Has Failed

                ErrorHandlerForm.ErrorTrapHeaderLABEL.Text = "UNEXPECTED ERROR EXPORTING ITEMS"
                ErrorHandlerForm.ErrorTrapYesBUTTON.Visible = False
                ErrorHandlerForm.ErrorTrapNoBUTTON.Visible = True
                ErrorHandlerForm.ErrorTrapNoBUTTON.Text = "Cancel Import"
                ErrorHandlerForm.ErrorTrapMessageTEXTBOX.Text = "DiaBase cannot export the selected items. The export has failed... " & vbCrLf & vbCrLf & "SYSTEM ERROR:" & vbCrLf & vbCrLf & SystemCode.message & vbCrLf & vbCrLf & "Please report this error if the problem persists."
                ErrorHandlerForm.StartPosition = FormStartPosition.CenterParent
                DialogResult = ErrorHandlerForm.ShowDialog()




        End Select
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------
    'User List CONTEXT MENU - REMOVE ITEMS - DELETES SELECTED ITEMS FROM THE USER LIST (this should be in user list funxctions module)
    '---------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearItemsToolStripMenuItem.Click
        'Clear all items from User Ref List When Selected from context menu

        Dim ReturnLocation As Integer = -1 '                                               Remember current item location to return highlight bar to after delete

        If UserLISTBOX.SelectedIndex > 0 Then ReturnLocation = UserLISTBOX.SelectedIndex ' puts the highlight bar one above the fist item deleted in the 
        '                                                                                ' block after clear is finished for easier user refrence
        '                                                                                ' if there is only one item before delete no item is highlighted

        'Itterate and remove all selected itmes in user list
        If UserLISTBOX.SelectedIndices.Count = 1 Then
            UserObjects.RemoveAt(UserLISTBOX.SelectedIndex)
            UserLISTBOX.Items.RemoveAt(UserLISTBOX.SelectedIndex)
        Else
            For ItemIndex = UserLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                UserObjects.RemoveAt(ItemIndex)
                UserLISTBOX.Items.RemoveAt(ItemIndex)
            Next
        End If

        'Attempts to select the prior selected item after delete else selects nothing (selecting nothing stops list from moving if all else fails)
        Me.ClearItemStats()
        UserLISTBOX.SelectedIndex = -1
        UserLISTBOX.SelectedIndex = ReturnLocation

    End Sub



    '-----------------------------------------------------------------------------------------------------------------------------------------
    'DISPLAYS THE STATISTICS FOR THE CURRNETLY SELECTED ITEM IN THE ALL ITMES LIST (user list has its own routine in UserListFunctions)
    '-----------------------------------------------------------------------------------------------------------------------------------------

    Sub DisplayItemStats(ItemIndex)

        If ItemIndex < 0 Or ItemIndex >= ItemObjects.Count Then ClearItemStats() : Return

        ItemStatsRICHTEXTBOX.Clear() : ItemNameRICHTEXTBOX.Clear() 'moved this here as occassionally getting double display nfi why

        'Display mule details
        MuleRealmTEXTBOX.Text = ItemObjects(ItemIndex).ItemRealm
        MuleAccountTEXTBOX.Text = ItemObjects(ItemIndex).MuleAccount
        MuleNameTEXTBOX.Text = ItemObjects(ItemIndex).MuleName
        If AppSettings.HideMulePass = False Then MulePasswordTEXTBOX.UseSystemPasswordChar = False
        If AppSettings.HideMulePass = True Then MulePasswordTEXTBOX.UseSystemPasswordChar = True
        MulePasswordTEXTBOX.Text = ItemObjects(ItemIndex).MulePass
        If ItemObjects(ItemIndex).HardCore = True Then CoreTypeTEXTBOX.Text = "HardCore"
        If ItemObjects(ItemIndex).HardCore = False Then CoreTypeTEXTBOX.Text = "SoftCore"
        If ItemObjects(ItemIndex).Ladder = False Then LadderTEXTBOX.Text = "Non Ladder"
        If ItemObjects(ItemIndex).Ladder = True Then LadderTEXTBOX.Text = "Ladder"

        Dim DisplayColour As String = ItemObjects(ItemIndex).ItemQuality
        Dim ColourCount1 As Integer = ItemObjects(ItemIndex).ItemQuality.Length








        'Normal And Superior - White
        If (DisplayColour = "Normal" Or DisplayColour = "Superior") And ItemObjects(ItemIndex).RuneWord = False Then
            If ItemObjects(ItemIndex).ItemBase = "Rune" Then
                ItemNameRICHTEXTBOX.SelectionColor = Color.Orange
                ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Quest and Special Items - Orange
            If ItemObjects(ItemIndex).ItemBase = "Quest" Then
                ItemNameRICHTEXTBOX.SelectionColor = Color.Orange
                ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Runeword Items - White
            If ItemObjects(ItemIndex).ItemName.IndexOf("Rune") = -1 And ItemObjects(ItemIndex).ItemBase <> "Quest" Then
                ItemNameRICHTEXTBOX.SelectionColor = Color.White
                ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If
        End If

        'Magic items - Blue
        If DisplayColour = "Magic" Then
            ItemNameRICHTEXTBOX.SelectionColor = Color.DodgerBlue
            ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Rares - Yellow
        If DisplayColour = "Rare" Then
            ItemNameRICHTEXTBOX.SelectionColor = Color.Yellow
            ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Crafted - Gold
        If DisplayColour = "Crafted" Then
            ItemNameRICHTEXTBOX.SelectionColor = Color.DarkGoldenrod
            ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Set Items - Green
        If DisplayColour = "Set" Then
            ItemNameRICHTEXTBOX.SelectionColor = Color.Green
            ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Uniques - Light Gold
        If DisplayColour = "Unique" Or ItemObjects(ItemIndex).RuneWord = True Then
            ItemNameRICHTEXTBOX.SelectionColor = Color.BurlyWood
            ItemNameRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        ItemNameRICHTEXTBOX.SelectAll() : ItemNameRICHTEXTBOX.SelectionAlignment = HorizontalAlignment.Center






        'RuneWord String - Still Not Reporting Right
        'If ItemObjects(ItemIndex).RuneWord = True Then
        'ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
        'ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).Stat1 & vbCrLf
        'End If
        'ItemStatsRICHTEXTBOX.AppendText(vbCrLf) 'Spacer line after Item Name and class Always 

        ColourCount1 = ItemStatsRICHTEXTBOX.TextLength 'Used to Count number of lines to calculate selection to colour text selection for the Basic Info Block - this var represents the starting point to colour

        'White text for basic info Block - Line spaceing added between each section (if needed)
        If ItemObjects(ItemIndex).OneHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText("One Hand Damage: " & ItemObjects(ItemIndex).OneHandDamageMin & " to" & ItemObjects(ItemIndex).OneHandDamageMax & vbCrLf)
        If ItemObjects(ItemIndex).TwoHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText("Two Hand Damage: " & ItemObjects(ItemIndex).TwoHandDamageMin & " to" & ItemObjects(ItemIndex).TwoHandDamageMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).OneHandDamageMax > 0 Or ItemObjects(ItemIndex).TwoHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Item Defensive Values
        If ItemObjects(ItemIndex).Defense > 0 Then ItemStatsRICHTEXTBOX.AppendText("Defense: " & ItemObjects(ItemIndex).Defense & vbCrLf)
        If ItemObjects(ItemIndex).ChanceToBlock > 0 Then ItemStatsRICHTEXTBOX.AppendText("Chance To Block: " & ItemObjects(ItemIndex).ChanceToBlock & "%" & vbCrLf)
        If ItemObjects(ItemIndex).DurabilityMin > 0 Then ItemStatsRICHTEXTBOX.AppendText("Durability: " & ItemObjects(ItemIndex).DurabilityMin & " of" & ItemObjects(ItemIndex).DurabilityMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).Defense > 0 Or ItemObjects(ItemIndex).ChanceToBlock > 0 Or ItemObjects(ItemIndex).DurabilityMin > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Item Requirements
        If ItemObjects(ItemIndex).RequiredStrength > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Strength: " & ItemObjects(ItemIndex).RequiredStrength & vbCrLf)
        If ItemObjects(ItemIndex).RequiredDexterity > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Dexterity: " & ItemObjects(ItemIndex).RequiredDexterity & vbCrLf)
        If ItemObjects(ItemIndex).RequiredLevel > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Level: " & ItemObjects(ItemIndex).RequiredLevel & vbCrLf)

        'Required Character Displayed Red and in Brackets to copy the ingame D2 themed layout
        If ItemObjects(ItemIndex).RequiredCharacter <> Nothing Then
            ItemStatsRICHTEXTBOX.SelectedText = "[" & ItemObjects(ItemIndex).RequiredCharacter & " Only]" & vbCrLf
        End If

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).RequiredStrength > 0 Or ItemObjects(ItemIndex).RequiredDexterity > 0 Or ItemObjects(ItemIndex).RequiredCharacter <> Nothing And ItemObjects(ItemIndex).RequiredLevel > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Item Attack Class And Speed
        If ItemObjects(ItemIndex).AttackClass <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).AttackClass & " Class") : If ItemObjects(ItemIndex).AttackSpeed <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(" - " & ItemObjects(ItemIndex).AttackSpeed & vbCrLf) Else ItemStatsRICHTEXTBOX.AppendText(vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).AttackClass <> Nothing Or ItemObjects(ItemIndex).AttackSpeed <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
            If ItemObjects(ItemIndex).RequiredLevel > 0 And ItemObjects(ItemIndex).AttackClass = Nothing And ItemObjects(ItemIndex).AttackSpeed = Nothing And ItemObjects(ItemIndex).RequiredCharacter = Nothing And ItemObjects(ItemIndex).RequiredDexterity = 0 And ItemObjects(ItemIndex).RequiredStrength = 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Colour Above Displayed Basic Info Text Block White
        Dim ColourCount2 As Integer = ItemStatsRICHTEXTBOX.TextLength - ColourCount1 'Calculate difference between Basic Info Block and Other Stats to Color Basic Info - this var represents the finishing point to colour (rob: adding a ditty to avoid colouring the ilevel blue, keeping it white so its clearly seperate from the stats block)
        ItemStatsRICHTEXTBOX.Select(ColourCount1, ColourCount2)
        ItemStatsRICHTEXTBOX.SelectionColor = Color.White

        'Unique Attributes Block as Default Blue (No need to colour these as they are blue by default)
        If ItemObjects(ItemIndex).Stat1 <> Nothing And ItemObjects(ItemIndex).RuneWord = False Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat1 & vbCrLf)
        If ItemObjects(ItemIndex).Stat2 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat2 & vbCrLf)
        If ItemObjects(ItemIndex).Stat3 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat3 & vbCrLf)
        If ItemObjects(ItemIndex).Stat4 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat4 & vbCrLf)
        If ItemObjects(ItemIndex).Stat5 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat5 & vbCrLf)
        If ItemObjects(ItemIndex).Stat6 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat6 & vbCrLf)
        If ItemObjects(ItemIndex).Stat7 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat7 & vbCrLf)
        If ItemObjects(ItemIndex).Stat8 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat8 & vbCrLf)
        If ItemObjects(ItemIndex).Stat9 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat9 & vbCrLf)
        If ItemObjects(ItemIndex).Stat10 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat10 & vbCrLf)
        If ItemObjects(ItemIndex).Stat11 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat11 & vbCrLf)
        If ItemObjects(ItemIndex).Stat12 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat12 & vbCrLf)
        If ItemObjects(ItemIndex).Stat13 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat13 & vbCrLf)
        If ItemObjects(ItemIndex).Stat14 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat14 & vbCrLf)
        If ItemObjects(ItemIndex).Stat15 <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).Stat15 & vbCrLf)


        ItemStatsRICHTEXTBOX.AppendText(vbCrLf & "Item Level: " & ItemObjects(ItemIndex).Itemlevel)

        'Select All and Center Justify Text Allignment in the ItemStatsRICHTEXTBOX - Display Items Routine is DONE :)
        ItemStatsRICHTEXTBOX.SelectAll()
        ItemStatsRICHTEXTBOX.SelectionAlignment = HorizontalAlignment.Center

        ItemSkinPICTUREBOX.Load("Skins\" + ImageArray(ItemObjects(ItemIndex).ItemImage) + ".jpg")

        'THIS DITTY CHANGES THE "Item Level: 00" LINE FROM BLUE TO WHITE (looks nicer and seperates it from the unique attribs block)
        'NOTE TO MYSELF: is something not right with the runes display, seems to add extra spaces, only does it to runes, atm nfi wtf, perhaps logging differs to other unique attribs blocks is my guess REMEMBER TO FIX!
        Dim linecount As Integer = 0
        For Each Line In ItemStatsRICHTEXTBOX.Lines
            If Line.IndexOf("Item Level:") > -1 Then ItemStatsRICHTEXTBOX.Select(ItemStatsRICHTEXTBOX.Text.Length - Len(Line), Len(Line))
            linecount = linecount + 1
        Next
        ItemStatsRICHTEXTBOX.SelectionColor = Color.White


    End Sub


    '----------------------------------------------------------------------------------------------------------------------------------------
    'lists All Mule Accounts in database and displays attached realm and mule
    '----------------------------------------------------------------------------------------------------------------------------------------

    Private Sub BuildMuleListMainMenu_Click(sender As Object, e As EventArgs)

        ' just a tempory listing - may keep as an app feature? Just needed some way of getting muleaccounts back :)
        ' removed previous code as it was working for some reason
        TradeListRICHTEXTBOX.Clear()
        For index = 0 To ItemObjects.Count - 1
            Dim temp = ItemObjects(index).ItemRealm & " / " & ItemObjects(index).MuleAccount & " / " & ItemObjects(index).MulePass

            If TradeListRICHTEXTBOX.Text.Contains(temp) = True Then Continue For

            TradeListRICHTEXTBOX.AppendText(temp & vbCrLf)

        Next

        DupesList(False)
        Return

    End Sub

    Private Sub LoadMuleMainMenu_Click(sender As Object, e As EventArgs) Handles LoadItemMuleMENUITEM.Click
        Dim Iindex = AllItemsLISTBOX.SelectedIndex
        If Iindex = -1 Then Return
        WriteLoaderFile(Iindex)
        'loadD2(Iindex)
    End Sub

    Private Sub SetAllNonLadderMainMenu_Click(sender As Object, e As EventArgs) Handles SetAllNonLadderMENUITEM.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        TimerStartPauseButton(sender, e)

        'Setup User Input Form For user warning
        UserInput.Text = "ALERT ALERT"
        UserInput.UserInputHeaderLABEL.Text = "RESET"

        UserInput.UserInputTEXTBOX.Text = Nothing
        UserInput.UserInputMessageLABEL.Text = "Are you sure you want to set all items in current database to Non Ladder?"
        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "OK"

        UserInput.DatabaseManagerBorder1LABEL.Visible = False
        UserInput.DatabaseManagerBorder2LABEL.Visible = False
        UserInput.DatabaseManagerBorder3LABEL.Visible = False
        UserInput.DatabaseManagerBorder4LABEL.Visible = False
        UserInput.UserInputTEXTBOX.Visible = False

        UserInput.UserInputTEXTBOX.SelectionStart = 0 : UserInput.UserInputTEXTBOX.SelectionLength = Len(UserInput.UserInputTEXTBOX.Text)
        UserInput.UserInputTEXTBOX.Select()

        'Confirms the backup proceedure and proceeds on confirmation
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Dim DialogResult = UserInput.ShowDialog
        If DialogResult = Windows.Forms.DialogResult.Yes Then
            CreateBackup(AppSettings.CurrentDatabase)
            If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
            For index = 0 To ItemObjects.Count - 1
                ItemObjects(index).Ladder = False
            Next
        End If
        TimerStartPauseButton(sender, e)
    End Sub

    'sub to set to nonladder based on mule logged date
    Private Sub SetLadderByDateMainMenu_Click(sender As Object, e As EventArgs) Handles SetLadderByDateMENUITEM.Click
        'may need try catch here
        Dim resetdate As Date = Date.ParseExact(AppSettings.ResetDate, "d/m/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        Dim temp As Date

        For index = 0 To ItemObjects.Count - 1
            ItemObjects(index).ImportDate = ItemObjects(index).ImportDate.Replace("\", "/")
            temp = Date.ParseExact(ItemObjects(index).ImportDate, "d/m/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            If temp > resetdate And ItemObjects(index).Ladder = True Then
                ItemObjects(index).Ladder = False
            End If
        Next

    End Sub

    Private Sub SendToTradeListSearchMenu_Click(sender As Object, e As EventArgs) Handles SendToTradeListSearchMenu.Click
        If AutoLoggerRunning = True Then Return

        ImportTimer.Stop()

        If SearchLISTBOX.SelectedIndices.Count > 0 Then
            Dim a As Integer = 0
            Dim count As Integer = 0

            For index = 0 To SearchLISTBOX.SelectedIndices.Count - 1

                a = SearchReferenceList(SearchLISTBOX.SelectedIndices(index))
                Dim Temp = ItemObjects(a).ItemName
                If ItemObjects(a).ItemBase = "Rune" Or ItemObjects(a).ItemBase = "Gem" Or ItemObjects(a).ItemName.IndexOf("Token") > -1 Or ItemObjects(a).ItemName.IndexOf("Key of") > -1 Or ItemObjects(a).ItemName.IndexOf("Essence") > -1 Then
                    If ItemObjects(a).ItemName.IndexOf("Token") > -1 Then Temp = "Token"
                    TradeListRICHTEXTBOX.AppendText(Temp & vbCrLf & vbCrLf)
                Else
                    SendToTradeList(a)
                End If
            Next

        End If

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListboxTABCONTROL.SelectTab(2)


        ItemTallyTEXTBOX.Text = SearchReferenceList.Count & " - Entries" ' simpler way to indicate count
        ImportTimer.Start()

    End Sub

    Private Sub ClearItemSearchCMenu_Click(sender As Object, e As EventArgs) Handles ClearItemSearchCMenu.Click
        If AutoLoggerRunning = True Then Return

        ImportTimer.Stop()

        Dim FocusOnExit As Integer = SearchLISTBOX.SelectedIndex
        Dim a As Integer
        For index = SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
            a = SearchLISTBOX.SelectedIndices(index)
            SearchLISTBOX.Items.RemoveAt(a)
            SearchReferenceList.RemoveAt(a)
        Next
        SearchLISTBOX.SelectedItem = -1

        'KEEPS FocusOnExit WITHIN LISTBOX BOUNDARYS
        If FocusOnExit >= (SearchLISTBOX.Items.Count) Then FocusOnExit = SearchLISTBOX.Items.Count - 1
        If SearchLISTBOX.Items.Count = 1 Then FocusOnExit = 0
        If SearchLISTBOX.Items.Count = 0 Then FocusOnExit = -1
        SearchLISTBOX.SelectedIndex = FocusOnExit
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"
    End Sub

    Private Sub DeleteItemSearchCMenu_Click(sender As Object, e As EventArgs) Handles DeleteItemSearchCMenu.Click
        Dim a As Integer
        Dim b As Integer
        Dim FocusOnExit As Integer = SearchLISTBOX.SelectedIndex

        For index = SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
            a = SearchLISTBOX.SelectedIndices(index)
            b = SearchReferenceList(a)
            SearchLISTBOX.Items.RemoveAt(a)
            AllItemsLISTBOX.Items.RemoveAt(b)
            ItemObjects.RemoveAt(b)
            SearchReferenceList.RemoveAt(a)
            'For x = a To SearchReferenceList.Count - 1' redundant ?
            '    SearchReferenceList(x) = SearchReferenceList(x) - 1
            'Next
        Next
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"

        'SET THE DELETED OBJECT LOCATION IN THE LIST AS THE HIGHLIGHTED ITEM ON RETURN FROM DETETE
        If FocusOnExit >= (SearchLISTBOX.Items.Count) Then FocusOnExit = SearchLISTBOX.Items.Count - 1
        If SearchLISTBOX.Items.Count = 1 Then FocusOnExit = 0
        If SearchLISTBOX.Items.Count = 0 Then FocusOnExit = -1
        SearchLISTBOX.SelectedIndex = FocusOnExit
    End Sub

    Private Sub SelectAllSearchCMenu_Click(sender As Object, e As EventArgs) Handles SelectAllSearchCMenu.Click
        SendMessage(SearchLISTBOX.Handle, &H185, New IntPtr(1), New IntPtr(-1))
        TriggerUpdate.SetValue(SearchLISTBOX.SelectedItems, True)
        TriggerIndexChanged.Invoke(SearchLISTBOX, New Object() {New EventArgs})
    End Sub

    Private Sub SendAllItemsToUserListSearchMenu_Click(sender As Object, e As EventArgs) Handles SendAllItemsToUserListSearchMenu.Click
        SearchItemsToUserList()
    End Sub


    'Populate search dropdown routine - triggers when field value changes (Third Time Lucky??)
    Private Sub SearchFieldCOMBOBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchFieldCOMBOBOX.SelectedIndexChanged
        Me.SearchWordCOMBOBOX.Items.Clear() 'clear old list out ready for new selection

        Select Case (SearchFieldCOMBOBOX.Text)
            Case "Item Quality"
                SearchWordCOMBOBOX.Items.Add("Normal")
                SearchWordCOMBOBOX.Items.Add("Superior")
                SearchWordCOMBOBOX.Items.Add("Set")
                SearchWordCOMBOBOX.Items.Add("Rare")
                SearchWordCOMBOBOX.Items.Add("Magic")
                SearchWordCOMBOBOX.Items.Add("Unique")
                SearchWordCOMBOBOX.Items.Add("Crafted")

        End Select

    End Sub


    Private Sub WestRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles WestRealmCHECKBOX.CheckedChanged
        If WestRealmCHECKBOX.Checked = True Then
            EastRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub EastRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles EastRealmCHECKBOX.CheckedChanged
        If EastRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub AsiaRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles AsiaRealmCHECKBOX.CheckedChanged
        If AsiaRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            EastRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub EuropeRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles EuropeRealmCHECKBOX.CheckedChanged
        If EuropeRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            EastRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub SendToUserListItemsCMenu_Click(sender As Object, e As EventArgs) Handles SendToUserListItemsCMenu.Click
        SelectedItemsToUserList()
        If UserLISTBOX.Items.Count > 0 And UserLISTBOX.SelectedIndex = -1 Then UserLISTBOX.SelectedIndex = 0

    End Sub

    Private Sub SendAllToTradeListSearchMenu_Click(sender As Object, e As EventArgs) Handles SendAllToTradeListSearchMenu.Click
        If AutoLoggerRunning = True Then Return

        ImportTimer.Stop()

        If SearchLISTBOX.Items.Count > 0 Then
            Dim a As Integer = 0
            Dim count As Integer = 0

            For index = 0 To SearchLISTBOX.Items.Count - 1

                a = SearchReferenceList(index)
                Dim Temp = ItemObjects(a).ItemName
                If ItemObjects(a).ItemBase = "Rune" Or ItemObjects(a).ItemBase = "Gem" Or ItemObjects(a).ItemName.IndexOf("Token") > -1 Or ItemObjects(a).ItemName.IndexOf("Key of") > -1 Or ItemObjects(a).ItemName.IndexOf("Essence") > -1 Then
                    If ItemObjects(a).ItemName.IndexOf("Token") > -1 Then Temp = "Token"
                    TradeListRICHTEXTBOX.AppendText(Temp & vbCrLf & vbCrLf)
                Else
                    SendToTradeList(a)
                End If
            Next

        End If

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListboxTABCONTROL.SelectTab(2)


        ItemTallyTEXTBOX.Text = SearchReferenceList.Count & " - Entries" ' simpler way to indicate count
        ImportTimer.Start()
    End Sub

    Private Sub CombineDupesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CombineDupesToolStripMenuItem.Click
        DupesList(True)
    End Sub

    Private Sub CopyAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyAllToolStripMenuItem.Click
        My.Computer.Clipboard.SetText(TradeListRICHTEXTBOX.Text)
    End Sub


    Private Sub ClearAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearAllToolStripMenuItem.Click
        TradeListRICHTEXTBOX.Clear()
    End Sub

    Private Sub AddItemItemsCMenu_Click(sender As Object, e As EventArgs) Handles AddItemItemsCMenu.Click

    End Sub

    Private Sub UserLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles UserLISTBOX.MouseDown




        'Display User List Context Menu
        If e.Button = Windows.Forms.MouseButtons.Right Then
            If UserLISTBOX.Items.Count > 0 Then
                Me.UserListCONTEXTMENUSTRIP.Show(Control.MousePosition)
            End If
        End If
    End Sub

    Private Sub ExportItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportItemsToolStripMenuItem.Click
        ImportTimer.Stop()
        Export.ShowDialog()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        ImportTimer.Start()
    End Sub

    'logic arguments for context  sensitive options in user list context menu
    Private Sub UserListCONTEXTMENUSTRIP_Opened(sender As Object, e As EventArgs) Handles UserListCONTEXTMENUSTRIP.Opened

        '================================================================================================================================================================================================
        'ROBS NOTE TO HIMSELF:  CHECK THE BELOW ARGUMENTS FUNCTIONALITY ACTUALLY WORKS AS INTENDED - TO DISABLE MANIPULATION OF ITEMS THAT ARE IN THE USER LIST BUT ARE NOT PART OF THE CURRENT DATABASE
        '================================================================================================================================================================================================

        'enables or disables send to trade list depending if current database holds the selected item
        If UserLISTBOX.SelectedIndex <> -1 Then If DatabaseFileNameTEXTBOX.Text = UserObjects(UserLISTBOX.SelectedIndex).DatabaseFilename Then SendToTradeListToolStripMenuItem1.Enabled = True Else SendToTradeListToolStripMenuItem1.Enabled = False

        'enables or disables delete depending if current database holds the selected item
        If UserLISTBOX.SelectedIndex <> -1 Then If DatabaseFileNameTEXTBOX.Text = UserObjects(UserLISTBOX.SelectedIndex).DatabaseFilename Then DeleteItemsToolStripMenuItem1.Enabled = True Else DeleteItemsToolStripMenuItem1.Enabled = False

        'enables or disables export items depending if current database holds the selected item
        If UserLISTBOX.SelectedIndex <> -1 Then If DatabaseFileNameTEXTBOX.Text = UserObjects(UserLISTBOX.SelectedIndex).DatabaseFilename Then ExportItemsToolStripMenuItem.Enabled = True Else ExportItemsToolStripMenuItem.Enabled = False


    End Sub

    '--------------------------------------------------------------------------------------------------------------------------
    'NEDS GEM EFFECT HANDLER - Plays A Funky D2 Sound Effect And Then Hides The Gem For Rest Of Session - A Little Fun From Ned
    '--------------------------------------------------------------------------------------------------------------------------
    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles HiddenGemPICTUREBOX.Click
        My.Computer.Audio.Play(My.Resources.DiabloTaunt, AudioPlayMode.Background)
        HiddenGemPICTUREBOX.Hide()
    End Sub

    'CLEARS OUT ALL TEST IN THE IMPORT LOG - Also Refeshes Session Start String At Top Of Log
    Private Sub ClearImportLogMENUITEM_Click(sender As Object, e As EventArgs) Handles ClearImportLogMENUITEM.Click
        ImportLogRICHTEXTBOX.Text = Date.Today & " @ " & TimeOfDay & " - Session Start, AutoLogger Ready." & vbCrLf
    End Sub

    Private Sub HideDupesMENUITEM_Click(sender As Object, e As EventArgs) Handles HideDupesMENUITEM.Click
        If HideDupesMENUITEM.Checked = True Then
            HideDupesMENUITEM.Checked = False
            AppSettings.HideDupes = False
            Return
        End If

        'Update LineBreaks settings to true
        If HideDupesMENUITEM.Checked = False Then
            HideDupesMENUITEM.Checked = True
            AppSettings.HideDupes = True
        End If
    End Sub
End Class
