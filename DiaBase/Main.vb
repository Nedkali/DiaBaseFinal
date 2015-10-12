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

    Sub StartTimer()
        Timercount = 0
        TimerSeconds = AppSettings.AutoLoggingDelay * 60
        ToolStripProgressBar1.Maximum = 100
        ImportTimer.Interval = 1000
        ImportTimer.Start()
        RichTextBox1.Text = "AutoLogging Enabled" & vbCrLf
    End Sub


    'IMPORT DELAY COUNTER - HANDLES IMPORT DLEAY AUTOLOGGING SYSTEM - BRANCES TO IMPORT ROUTINE WHEN DELAY IS UP
    Private Sub TimerEventProcessor(ByVal myObject As Object, ByVal MyEventArgs As EventArgs) Handles ImportTimer.Tick
        Timercount = Timercount + 1
        If Timercount > TimerSeconds Then
            Timercount = 0
            ImportTimer.Stop()
            RichTextBox1.Text = "Checking for new Log Files" & vbCrLf
            AutoLoggerRunning = True
            ImportLogFiles(False)
            AutoLoggerRunning = False
            ImportTimer.Start()
            RichTextBox1.AppendText("AutoLogging Enabled" & vbCrLf)
        End If

        Dim Timerprogress As Integer = Math.Round((Timercount / TimerSeconds) * 100)
        ToolStripProgressBar1.Value = Timerprogress
        ToolStripStatusLabel2.Text = "< " & Math.Ceiling((TimerSeconds - Timercount) / 60) & " mins"
    End Sub

    'TIMER BUTTON pause & restart routine
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles TimerControl.Click
        If TimerControl.Text = "Pause" Then
            ImportTimer.Stop()
            TimerRestart = True
            TimerControl.Text = "Start"
            RichTextBox1.Text = "AutoLogger Paused" & vbCrLf
            Return
        End If

        If TimerControl.Text = "Start" And TimerRestart = True Then
            ImportTimer.Start()
            TimerControl.Text = "Pause"
            RichTextBox1.Text = "AutoLogging Enabled" & vbCrLf
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
        Me.Text = VersionAndRevision                            'Display Version and Revision Number in Main Form Title Bar

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
            TimerControl.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

        'Set foucus on the main listbox
        AllItemsLISTBOX.Select() : ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Total Items"

        'Play the introduction laugh
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.BigDLaugh, AudioPlayMode.Background)

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SHOWN MAIN FORM EVENT HANDLER  - Checks Paths and Branches to settings if path check fails
    '                               - Loads Default Database File (if there is one)
    '                               - If database is loaded also populates main listbox
    '                               - Focus On 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Main_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown


        'Verify The current Paths In Settings Are Correct - Auto Open Settings Window If Verify Fails
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = False Or My.Computer.FileSystem.FileExists(AppSettings.DefaultDatabase) = False Then
            Settings.ShowDialog()
        End If

        'Branch to Load Default Database and populate main listbox once all setting proceedures are absolutly completed with all potential path errors handled
        If AppSettings.DefaultDatabase <> Nothing Then OpenDatabase(AppSettings.DefaultDatabase) : PopulateAllItemsLISTBOX()

        StartTimer()

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
        Dim a As Integer = AllItemsLISTBOX.SelectedIndex
        If a > -1 Then
            AllItemsLISTBOX.SelectedIndex = a
            DisplayItemStats(a)
            If AllItemsLISTBOX.SelectedItems.Count > 1 Then ItemTallyTEXTBOX.Text = AllItemsLISTBOX.SelectedItems.Count & " - Selected Items" : Return
        End If
        ItemTallyTEXTBOX.Text = ItemObjects.Count & " - Total Items"
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SEARCH LISTBOX INDEX CHANGED HANDLER       - Focuses Main Listbox Selected Item to Match the Serch Lists Currently Selected Item 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SearchLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SearchLISTBOX.SelectedIndexChanged
        Dim a As Integer = SearchLISTBOX.SelectedIndex
        If a > -1 Then
            AllItemsLISTBOX.SelectedIndex = SearchReferenceList(a)
            DisplayItemStats(SearchReferenceList(a))
            If SearchLISTBOX.SelectedItems.Count > 1 Then ItemTallyTEXTBOX.Text = SearchLISTBOX.SelectedItems.Count & " - Selected Items" : Return
        End If
        ItemTallyTEXTBOX.Text = SearchReferenceList.Count & " - Total Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'USER LISTBOX INDEX CHANGED HANDLER         - Branches to Display Item Statistics Routine
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub UserLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles UserLISTBOX.SelectedIndexChanged

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
        ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Total Items"
        DatabaseFileNameLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()
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
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Matches"
        DatabaseFileNameLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()
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
        ItemTallyTEXTBOX.Text = TradeListRICHTEXTBOX.Lines.Count & " - Trade Entries" ' this needs to be changed to a routine that counts the BLANK lines between trade entries
        DatabaseFileNameLABEL.Hide()
        DatabaseFileNameTEXTBOX.Hide()
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
        ItemTallyTEXTBOX.Text = (UserLISTBOX.Items.Count & " - User Entries")
        DatabaseFileNameLABEL.Show()
        DatabaseFileNameTEXTBOX.Show()
    End Sub
















    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DISPLAY SETTINGS FORM - Settings Window Handles All Global Config Functions for the Entire Application
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SettingsToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem1.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        Settings.ShowDialog()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Button3_Click(sender, e)

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - DISPLAY DATABASE MANAGER FORM - Handles Open Create Delete Rename Database files all in one place for (hopefully an improvment)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        DatabaseManager.ShowDialog()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Button3_Click(sender, e)

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

            End If
        Else
            e.Cancel = True 'Automatically cancels Exit Event If Autologger Is Running (Avoids Potential Import Errors)
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR- CLOSE APPLICATION HANDLER - Shuts down main form
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - IMPORT NOW FUNCTION - Activated the autologger on demand as opposed to waiting for the delay to time out
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ImportNowMenuItem_Click(sender As Object, e As EventArgs) Handles ImportNowMenuItem.Click
        AutoLoggerRunning = True
        Timercount = 0
        RichTextBox1.Text = "Checking for New Logs..." & vbCrLf
        ImportLogFiles(False)
        AutoLoggerRunning = False
        If AllItemsLISTBOX.Items.Count > 0 Then AllItemsLISTBOX.SelectedIndex = 0
        ListboxTABCONTROL.SelectTab(0)
        ItemTallyTEXTBOX.Text = ItemObjects.Count & " - Total Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - SAVE DATABASE FUNCTION - branches to save routine to save the current database 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        WriteToFile(0, AppSettings.DefaultDatabase, False)
        Button3_Click(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CREATE BACKUP BUTTONPRESS HANDLER - Branches to backup routine to create backup of current file 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)

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
        Button3_Click(sender, e)

    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - RESTORE FROM BACKUP BUTTON PRESS HANDLER - Branches to backup routine to restore the current file from backup 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub RestoreBackupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreBackupToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
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
        Button3_Click(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - EDIT ITEM FUNCTION - Displays edit item form and broances to EditItem.vb if 1 or more items are selected in the main listbox
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub EditExistingItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditExistingItemToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        If AllItemsLISTBOX.SelectedIndex = -1 Then
            ErrorHandler(2, 0, 0, 0) : Return
        End If
        iEdit = AllItemsLISTBOX.SelectedIndex
        Button3_Click(sender, e)
        ItemEdit.ShowDialog()
        Button3_Click(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Sort ItemObjects alphabetically
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        ItemTallyTEXTBOX.Text = ("Sorting A to Z)")
        ItemObjects.Sort(Function(x, y) x.ItemName.CompareTo(y.ItemName))
        PopulateAllItemsLISTBOX()
        Button3_Click(sender, e)
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - 'SENDS HIGHLIGHTED ITEMS TO THE TRADE LIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SendToTradeListToolStripMenuItem3_Click(sender As Object, e As EventArgs)

        If AutoLoggerRunning = True Then
            RichTextBox1.AppendText("Please wait Import in progress")
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
        DupesList(True)

        'SET TRADELIST HIGHLIGHT AND SELECT TRADE LIST TAB
        ListControlTabBUTTON.BackColor = Color.Black
        SearchListControlTabBUTTON.BackColor = Color.Black
        TradesListControlTabBUTTON.BackColor = Color.DimGray
        ListboxTABCONTROL.SelectTab(2)

        'SHORT ROUTINE TO COUNT TRADE ITEMS IN RICHTEXT3 BY COUNTING THE GAPS BETWEEN THE ITEMS (SUBTRACTS 1 DUE TO EMPTY LINE AT END OF TEXT) 
        Dim TradeItemCounter As Integer = 0
        For Each item In TradeListRICHTEXTBOX.Lines
            If item = Nothing Then TradeItemCounter = TradeItemCounter + 1
        Next
        If TradeItemCounter = 0 Then TradeItemCounter = 1
        ItemTallyTEXTBOX.Text = (TradeItemCounter - 1) & " - Trade Entries"
        ImportTimer.Start()

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - SELLECT ALL ITEMS
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        SendMessage(AllItemsLISTBOX.Handle, &H185, New IntPtr(1), New IntPtr(-1))
        TriggerUpdate.SetValue(AllItemsLISTBOX.SelectedItems, True)
        TriggerIndexChanged.Invoke(AllItemsLISTBOX, New Object() {New EventArgs})

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - OPENS ADD NEW ITEM FORM
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub AddNewItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddNewItemToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)

        Button3_Click(sender, e)
        ItemAdd.ShowDialog()
        Button3_Click(sender, e)

        Button3_Click(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - DELETE SELECTED ITEM/S
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DeleteItemToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteItemToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)

        Dim FocusOnExit As Integer = AllItemsLISTBOX.SelectedIndex

        'check for backup on edits set to true if so then backup now
        If AppSettings.BackupBeforeEdits = True Then CreateBackup(AppSettings.CurrentDatabase)
        If AllItemsLISTBOX.SelectedIndices.Count > 0 Then
            UnDoCount.Add(AllItemsLISTBOX.SelectedIndices.Count - 1)
            For index = AllItemsLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                Dim a As Integer = AllItemsLISTBOX.SelectedIndices(index)
                AllItemsLISTBOX.Items.RemoveAt(a)
                UnDo.Add(ItemObjects(a)) : UnDoPos.Add(a) : ToolStripMenuItem6.Enabled = True
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
            ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Total Items"

            'SET THE DELETED OBJECT LOCATION IN THE LIST AS THE HIGHLIGHTED ITEM ON RETURN FROM DETETE
            If FocusOnExit >= (AllItemsLISTBOX.Items.Count) Then FocusOnExit = AllItemsLISTBOX.Items.Count - 1
            If AllItemsLISTBOX.Items.Count = 1 Then FocusOnExit = 0
            If AllItemsLISTBOX.Items.Count = 0 Then FocusOnExit = -1
            AllItemsLISTBOX.SelectedIndex = FocusOnExit

        End If

        Button3_Click(sender, e)

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEAR SEARCH LISTING
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearSearchListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ClearSearchListToolStripMenuItem1.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        SearchLISTBOX.Items.Clear()
        SearchReferenceList.Clear()
        ItemTallyTEXTBOX.Text = SearchLISTBOX.Items.Count & " - Total Items"
        Button3_Click(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEARS TRADELIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearTradeListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearTradeListToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        TradeListRICHTEXTBOX.Clear()
        Button3_Click(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - CLEARS USERLIST
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ClearUserListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearUserListToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)
        UserLISTBOX.Items.Clear()
        Button3_Click(sender, e)
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Add blank lines to item displayed
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DisplayLineBreaksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisplayLineBreaksToolStripMenuItem.Click
        If DisplayLineBreaksToolStripMenuItem.Checked = True Then
            DisplayLineBreaksToolStripMenuItem.Checked = False
            DisplayItemStats(AllItemsLISTBOX.SelectedIndex)
            Return
        End If
        If DisplayLineBreaksToolStripMenuItem.Checked = False Then
            DisplayLineBreaksToolStripMenuItem.Checked = True
            DisplayItemStats(AllItemsLISTBOX.SelectedIndex)
        End If

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Rebuild database from Archived logs
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub RebuildDefaultDBaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RebuildDefaultDBaseToolStripMenuItem.Click
        AutoLoggerRunning = True
        RichTextBox1.Text = "Checking for New Logs" & vbCrLf
        ImportLogFiles(True)
        PopulateAllItemsLISTBOX()
        AutoLoggerRunning = False
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Hide dupes in search results
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        If ToolStripMenuItem4.Checked = True Then
            ToolStripMenuItem4.Checked = False
            AppSettings.HideDupes = False
            Return
        End If
        If ToolStripMenuItem4.Checked = False Then
            ToolStripMenuItem4.Checked = True
            AppSettings.HideDupes = True
        End If
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'MENU BAR - Undo delete - works but restores one at a time atm
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
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
            ToolStripMenuItem6.Enabled = False
        End If
        ItemTallyTEXTBOX.Text = AllItemsLISTBOX.Items.Count & " - Total Items"
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'All Item List MENU - rerouted to main menu option
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SortListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SortListToolStripMenuItem.Click
        ToolStripMenuItem3_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem1.Click
        DeleteItemToolStripMenuItem_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub EditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EditToolStripMenuItem.Click
        EditExistingItemToolStripMenuItem_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub SendToTradeListToolStripMenuItem3_Click_1(sender As Object, e As EventArgs) Handles SendToTradeListToolStripMenuItem3.Click
        SendToTradeListToolStripMenuItem3_Click(sender, e) 'links to main menu option
    End Sub

    Private Sub SelectAllToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem1.Click
        ToolStripMenuItem5_Click(sender, e) 'links to main menu option
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Search button - calls search routine
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub SearchBUTTON_Click(sender As Object, e As EventArgs) Handles SearchBUTTON.Click
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

        End Select
    End Sub

    Sub DisplayItemStats(ItemIndex)
        If ItemIndex < 0 Or ItemIndex >= ItemObjects.Count Then
            MuleRealmTEXTBOX.Clear()
            MuleAccountTEXTBOX.Clear()
            MuleNameTEXTBOX.Clear()
            MulePasswordTEXTBOX.Clear()
            CoreTypeTEXTBOX.Clear()
            ItemStatsRICHTEXTBOX.Clear()
            Return
        End If

        ItemStatsRICHTEXTBOX.Clear() 'moved this here as occassionally getting double display nfi why

        'Display mule details
        MuleRealmTEXTBOX.Text = ItemObjects(ItemIndex).ItemRealm
        MuleAccountTEXTBOX.Text = ItemObjects(ItemIndex).MuleAccount
        MuleNameTEXTBOX.Text = ItemObjects(ItemIndex).MuleName
        If AppSettings.HideMulePass = False Then MulePasswordTEXTBOX.UseSystemPasswordChar = False
        If AppSettings.HideMulePass = True Then MulePasswordTEXTBOX.UseSystemPasswordChar = True
        MulePasswordTEXTBOX.Text = ItemObjects(ItemIndex).MulePass

        If ItemObjects(ItemIndex).HardCore = True Then CoreTypeTEXTBOX.Text = "HardCore"
        If ItemObjects(ItemIndex).HardCore = False Then CoreTypeTEXTBOX.Text = "SoftCore"

        Dim DisplayColour As String = ItemObjects(ItemIndex).ItemQuality
        Dim ColourCount1 As Integer = ItemObjects(ItemIndex).ItemQuality.Length


        'Normal And SAuperior - White
        If (DisplayColour = "Normal" Or DisplayColour = "Superior") And ItemObjects(ItemIndex).RuneWord = False Then
            If ItemObjects(ItemIndex).ItemBase = "Rune" Then
                ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
                ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Quest and Special Items - Orange
            If ItemObjects(ItemIndex).ItemBase = "Quest" Then
                ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
                ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Runeword Items - White
            If ItemObjects(ItemIndex).ItemName.IndexOf("Rune") = -1 And ItemObjects(ItemIndex).ItemBase <> "Quest" Then
                ItemStatsRICHTEXTBOX.SelectionColor = Color.White
                ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
            End If
        End If

        'Magic items - Blue
        If DisplayColour = "Magic" Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.DodgerBlue
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Rares - Yellow
        If DisplayColour = "Rare" Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.Yellow
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Crafted - Gold
        If DisplayColour = "Crafted" Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.DarkGoldenrod
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Set Items - Green
        If DisplayColour = "Set" Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.Green
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Uniques - Light Gold
        If DisplayColour = "Unique" Or ItemObjects(ItemIndex).RuneWord = True Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.BurlyWood
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).ItemName & vbCrLf
        End If

        If ItemObjects(ItemIndex).RuneWord = True Then
            ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
            ItemStatsRICHTEXTBOX.SelectedText = ItemObjects(ItemIndex).Stat1 & vbCrLf
        End If

        ItemStatsRICHTEXTBOX.AppendText(vbCrLf) 'Spacer line after Item Name and class Always 


        ColourCount1 = ItemStatsRICHTEXTBOX.TextLength 'Used to Count number of lines to calculate selection to colour text selection for the Basic Info Block - this var represents the starting point to colour

        'White text for basic info Block - Line spaceing added between each section (if needed)
        If ItemObjects(ItemIndex).OneHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText("One Hand Damage: " & ItemObjects(ItemIndex).OneHandDamageMin & " to " & ItemObjects(ItemIndex).OneHandDamageMax & vbCrLf)
        If ItemObjects(ItemIndex).TwoHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText("Two Hand Damage: " & ItemObjects(ItemIndex).TwoHandDamageMin & " to " & ItemObjects(ItemIndex).TwoHandDamageMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksToolStripMenuItem.Checked = True Then
            If ItemObjects(ItemIndex).OneHandDamageMax > 0 Or ItemObjects(ItemIndex).TwoHandDamageMax > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        If ItemObjects(ItemIndex).Defense > 0 Then ItemStatsRICHTEXTBOX.AppendText("Defense: " & ItemObjects(ItemIndex).Defense & vbCrLf)
        If ItemObjects(ItemIndex).ChanceToBlock > 0 Then ItemStatsRICHTEXTBOX.AppendText("Chance To Block: " & ItemObjects(ItemIndex).ChanceToBlock & "%" & vbCrLf)
        If ItemObjects(ItemIndex).DurabilityMin > 0 Then ItemStatsRICHTEXTBOX.AppendText("Durability: " & ItemObjects(ItemIndex).DurabilityMin & " of " & ItemObjects(ItemIndex).DurabilityMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksToolStripMenuItem.Checked = True Then
            If ItemObjects(ItemIndex).Defense > 0 Or ItemObjects(ItemIndex).ChanceToBlock > 0 Or ItemObjects(ItemIndex).DurabilityMin > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        If ItemObjects(ItemIndex).RequiredStrength > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Strength: " & ItemObjects(ItemIndex).RequiredStrength & vbCrLf)
        If ItemObjects(ItemIndex).RequiredDexterity > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Dexterity: " & ItemObjects(ItemIndex).RequiredDexterity & vbCrLf)
        If ItemObjects(ItemIndex).RequiredCharacter <> Nothing Then
            ItemStatsRICHTEXTBOX.SelectedText = "[" & ItemObjects(ItemIndex).RequiredCharacter & " Only]" & vbCrLf
        End If

        If ItemObjects(ItemIndex).RequiredLevel > 0 Then ItemStatsRICHTEXTBOX.AppendText("Required Level: " & ItemObjects(ItemIndex).RequiredLevel & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksToolStripMenuItem.Checked = True Then
            If ItemObjects(ItemIndex).RequiredStrength > 0 Or ItemObjects(ItemIndex).RequiredDexterity > 0 Or ItemObjects(ItemIndex).RequiredCharacter = 0 And ItemObjects(ItemIndex).RequiredLevel > 0 Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        If ItemObjects(ItemIndex).AttackClass <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(ItemObjects(ItemIndex).AttackClass & " Class") : If ItemObjects(ItemIndex).AttackSpeed <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(" - " & ItemObjects(ItemIndex).AttackSpeed & vbCrLf) Else ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        'ADD lINE SPACING BASED ON OPTION SETTING
        If DisplayLineBreaksToolStripMenuItem.Checked = True Then
            If ItemObjects(ItemIndex).AttackClass <> Nothing Or ItemObjects(ItemIndex).AttackSpeed <> Nothing Then ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Colour Above Displayed Basic Info Text Block White
        Dim ColourCount2 As Integer = ItemStatsRICHTEXTBOX.TextLength - ColourCount1 'Calculate difference between Basic Info Block and Other Stats to Color Basic Info - this var represents the finishing point to colour
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


        ItemStatsRICHTEXTBOX.AppendText(vbCrLf & "Item Level: " & ItemObjects(ItemIndex).Itemlevel & vbCrLf)

        'Select All and Center Justify Text Allignment in the ItemStatsRICHTEXTBOX - Display Items Routine is DONE :)
        ItemStatsRICHTEXTBOX.SelectAll()
        ItemStatsRICHTEXTBOX.SelectionAlignment = HorizontalAlignment.Center

        ItemSkinPICTUREBOX.Load("Skins\" + ImageArray(ItemObjects(ItemIndex).ItemImage) + ".jpg")

    End Sub


    '----------------------------------------------------------------------------------------------------------------------------------------
    'lists All Mule Accounts in database and displays attached realm and mule
    '----------------------------------------------------------------------------------------------------------------------------------------

    Private Sub BuildMuleListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BuildMuleListToolStripMenuItem.Click

        ' just a tempory listing - may keep as an app feature? Just needed some way of getting muleaccounts back :)
        ' removed previous code as it was working for some reason
        For index = 0 To ItemObjects.Count - 1
            Dim temp = ItemObjects(index).ItemRealm & " / " & ItemObjects(index).MuleAccount & " / " & ItemObjects(index).MulePass

            If TradeListRICHTEXTBOX.Text.Contains(temp) = True Then Continue For

            TradeListRICHTEXTBOX.AppendText(temp & vbCrLf)

        Next

        DupesList(False)
        Return

    End Sub

    Private Sub LoadMuleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadMuleToolStripMenuItem.Click
        Dim Iindex = AllItemsLISTBOX.SelectedIndex
        If Iindex = -1 Then Return
        WriteLoaderFile(Iindex)
        'loadD2(Iindex)
    End Sub

    Private Sub AllItemsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AllItemsToolStripMenuItem.Click
        If AutoLoggerRunning = True Then ErrorHandler(1, 0, 0, 0) : Return
        Button3_Click(sender, e)

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
        Button3_Click(sender, e)
    End Sub
    Private Sub ByDateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByDateToolStripMenuItem.Click

        'AppSettings.ResetDate = "26/4/2015"
        Dim resetdate As Date = Date.ParseExact(AppSettings.ResetDate, "d/m/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
        Dim temp As Date
        Dim findoldest As Date = Date.Now()
        MessageBox.Show("Please Wait")
        Dim a = 0
        For index = 0 To ItemObjects.Count - 1
            ItemObjects(index).ImportDate = ItemObjects(index).ImportDate.Replace("\", "/")
            temp = Date.ParseExact(ItemObjects(index).ImportDate, "d/m/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            If temp > resetdate And ItemObjects(index).Ladder = True Then
                ItemObjects(index).Ladder = False
                a += 1
            End If
            If findoldest > temp Then
                findoldest = temp
            End If
        Next
        MessageBox.Show("Changed = " & a)
    End Sub


End Class
