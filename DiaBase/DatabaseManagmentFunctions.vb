
'========================================================================================================================================================================
'DATABASE MANAGMENT FUNCTIONS DATABASE - Handles all file functions for database files within this module
'========================================================================================================================================================================
Imports System.IO
Module DatabaseManagmentFunctions
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'ReadSettingFile SUBROUTINE TO READ THE CONFIG VALUE WITHIN SETTINGS FILE AND APPLY THEN TO THEIR GOLBAL CONFIG VARS - should be safe to use this sub anywhere its needed
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub ReadSettingsFile()

        Try
            If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Settings.cfg") = True Then
                Dim ReadFile As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Settings.cfg")
                AppSettings.EtalPath = ReadFile.ReadLine                'Etal Path
                AppSettings.DefaultDatabase = ReadFile.ReadLine         'Startup Database
                AppSettings.SoundMute = ReadFile.ReadLine               'Mute Sound Setting bool
                AppSettings.RemoveMuleDupes = ReadFile.ReadLine         'Remove mule dupe bool
                AppSettings.HideMulePass = ReadFile.ReadLine            'Hide Password Bool
                AppSettings.BackupBeforeImports = ReadFile.ReadLine     'Backup before imports bool 
                AppSettings.BackupBeforeEdits = ReadFile.ReadLine       'Backup before item edits bool
                AppSettings.AutoLoggingDelay = ReadFile.ReadLine        'Autologger delay
                AppSettings.DefaultRealm = ReadFile.ReadLine            'Mute Sound Setting bool
                AppSettings.DefaultPassword = ReadFile.ReadLine
                AppSettings.ResetDate = ReadFile.ReadLine               'Mute Sound Setting bool
                AppSettings.DisplayLineBreaks = ReadFile.ReadLine       'Stat display spaceing 
                ReadFile.Close()
            Else : Main.ErrorHandler(100, 0, 0, 0)                      'Our File Not Found Error Handler - pass to error handler
            End If

        Catch ex As Exception                                           'All Other Unforseen System Error Handler - pass to error handler
            Main.ErrorHandler(101, ex, 0, 0)
        End Try

        'Apply LineBreak Bool Value To Menu Control Checkstate
        If AppSettings.DisplayLineBreaks = True Then Main.DisplayLineBreaksMENUITEM.Checked = True
        If AppSettings.DisplayLineBreaks = False Then Main.DisplayLineBreaksMENUITEM.Checked = False

        'Apply Realm Search Checkbox Values
        If AppSettings.DefaultRealm = "USEast" Then Main.EastRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "USWest" Then Main.WestRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "Asia" Then Main.AsiaRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "Europe" Then Main.EuropeRealmCHECKBOX.Checked = True
    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' SAVE SETTINGS FILE - Updates AppSettings Class Variables to Settings.CFG file in the Install Directory - Used by Settings Form, First Run Routine, and Exit App Handler
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub SaveSettingsFile()
        Try

            'Check for valid entrys (avoid potential crash with null entries) and swap app runstate values (where nessicary) to settings class variables
            If AppSettings.DefaultDatabase = Nothing Then AppSettings.DefaultDatabase = AppSettings.InstallPath & "\Databases\Default.TXT"

            'technically only one of the following checkboxes should be checked, but not on first run i dont think.
            If Main.EastRealmCHECKBOX.Checked = True Then AppSettings.DefaultRealm = "USEast"
            If Main.WestRealmCHECKBOX.Checked = True Then AppSettings.DefaultRealm = "USWest"
            If Main.AsiaRealmCHECKBOX.Checked = True Then AppSettings.DefaultRealm = "Asia"
            If Main.EuropeRealmCHECKBOX.Checked = True Then AppSettings.DefaultRealm = "Europe"

            'applys the menu option checkstate to its partnered settings clas variable
            AppSettings.DisplayLineBreaks = Main.DisplayLineBreaksMENUITEM.CheckState

            'Writes the settings to file Settings.CFG Assumes no null entries exist at this point or lines will be skipped during save will crash when attempting to read it later
            Dim WriteFile As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(AppSettings.InstallPath + "\Settings.cfg", False)
            WriteFile.WriteLine(AppSettings.EtalPath)
            WriteFile.WriteLine(AppSettings.DefaultDatabase)
            WriteFile.WriteLine(AppSettings.SoundMute)
            WriteFile.WriteLine(AppSettings.RemoveMuleDupes)
            WriteFile.WriteLine(AppSettings.HideMulePass)
            WriteFile.WriteLine(AppSettings.BackupBeforeImports)
            WriteFile.WriteLine(AppSettings.BackupBeforeEdits)
            WriteFile.WriteLine(AppSettings.AutoLoggingDelay)
            WriteFile.WriteLine(AppSettings.DefaultRealm)
            WriteFile.WriteLine(AppSettings.DefaultPassword)
            WriteFile.WriteLine(AppSettings.ResetDate)
            WriteFile.WriteLine(AppSettings.DisplayLineBreaks)
            WriteFile.Close()
        Catch ex As Exception
            'Branch to error handler with unique code and system error code if save fails for whatever reason
            Main.ErrorHandler(301, ex, 0, 0)
        End Try

        'Update Working Vars...
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\Configs\USWest\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "NED"
        If (My.Computer.FileSystem.DirectoryExists(String.Concat(AppSettings.EtalPath, "\Scripts\AMS\MuleInventory"))) = True Then AppSettings.EtalVersion = "PUB"
        If AppSettings.EtalVersion = "NED" Then Main.Text = VersionAndRevision & " - Running Red Dragon Compataibility Mode"
        If AppSettings.EtalVersion = "PUB" Then Main.Text = VersionAndRevision & " - Running Black Empress Compatibility Mode"
    End Sub


    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'OPEN DATABASE ROUTINE  - Clears Out Old ItemObject Database
    '                       - Opens the selected database
    '                       - While Loading A File The Routine Also Checks Item Records in File Being Read Are In Sync With No Missing Field Values
    '                       - Also Branches to Error Handler On all Other Unexpected errors
    '                       -
    '                       -
    '--------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub OpenDatabase(DatabaseFilePath)
        Dim OpenDatabase As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(DatabaseFilePath)
        AppSettings.CurrentDatabase = DatabaseFilePath
        Dim CheckSpacerFlag As String = "--------------------"
        Dim temp As String = "" 'using this for debugging purposes
        Dim CountRecordsForErrorEvents As Integer = 0
        CloseFile()

        Try
            Do While OpenDatabase.EndOfStream = False
                Dim NewItem As New ItemDatabase '                           Define NewItem As A New Object for ItemDatabase Class
                CountRecordsForErrorEvents += 1
                CheckSpacerFlag = OpenDatabase.ReadLine()
                If CheckSpacerFlag <> "--------------------" Then Exit Do ' Check Spacer Flag Ditty
                NewItem.ItemName = OpenDatabase.ReadLine
                NewItem.Itemlevel = OpenDatabase.ReadLine
                NewItem.ItemRealm = OpenDatabase.ReadLine
                NewItem.ItemBase = OpenDatabase.ReadLine
                NewItem.ItemQuality = OpenDatabase.ReadLine
                NewItem.RequiredCharacter = OpenDatabase.ReadLine
                NewItem.EtherealItem = OpenDatabase.ReadLine
                NewItem.Sockets = OpenDatabase.ReadLine
                NewItem.RuneWord = OpenDatabase.ReadLine
                NewItem.ThrowDamageMin = OpenDatabase.ReadLine
                NewItem.ThrowDamageMax = OpenDatabase.ReadLine
                NewItem.OneHandDamageMin = OpenDatabase.ReadLine
                NewItem.OneHandDamageMax = OpenDatabase.ReadLine
                NewItem.TwoHandDamageMin = OpenDatabase.ReadLine
                NewItem.TwoHandDamageMax = OpenDatabase.ReadLine
                NewItem.Defense = OpenDatabase.ReadLine
                NewItem.ChanceToBlock = OpenDatabase.ReadLine
                NewItem.QuantityMin = OpenDatabase.ReadLine
                NewItem.QuantityMax = OpenDatabase.ReadLine
                NewItem.DurabilityMin = OpenDatabase.ReadLine
                NewItem.DurabilityMax = OpenDatabase.ReadLine
                NewItem.RequiredStrength = OpenDatabase.ReadLine
                NewItem.RequiredDexterity = OpenDatabase.ReadLine
                NewItem.RequiredLevel = OpenDatabase.ReadLine
                NewItem.AttackClass = OpenDatabase.ReadLine
                NewItem.AttackSpeed = OpenDatabase.ReadLine
                NewItem.Stat1 = OpenDatabase.ReadLine
                NewItem.Stat2 = OpenDatabase.ReadLine
                NewItem.Stat3 = OpenDatabase.ReadLine
                NewItem.Stat4 = OpenDatabase.ReadLine
                NewItem.Stat5 = OpenDatabase.ReadLine
                NewItem.Stat6 = OpenDatabase.ReadLine
                NewItem.Stat7 = OpenDatabase.ReadLine
                If NewItem.Stat7 = " " Then NewItem.Stat7 = Nothing
                NewItem.Stat8 = OpenDatabase.ReadLine
                If NewItem.Stat8 = " " Then NewItem.Stat8 = Nothing
                NewItem.Stat9 = OpenDatabase.ReadLine
                If NewItem.Stat9 = " " Then NewItem.Stat9 = Nothing
                NewItem.Stat10 = OpenDatabase.ReadLine
                If NewItem.Stat10 = " " Then NewItem.Stat10 = Nothing
                NewItem.Stat11 = OpenDatabase.ReadLine
                If NewItem.Stat11 = " " Then NewItem.Stat11 = Nothing
                NewItem.Stat12 = OpenDatabase.ReadLine
                If NewItem.Stat12 = " " Then NewItem.Stat12 = Nothing
                NewItem.Stat13 = OpenDatabase.ReadLine
                NewItem.Stat14 = OpenDatabase.ReadLine
                NewItem.Stat15 = OpenDatabase.ReadLine
                NewItem.MuleName = OpenDatabase.ReadLine
                NewItem.MuleAccount = OpenDatabase.ReadLine
                NewItem.MulePass = OpenDatabase.ReadLine
                NewItem.PickitAccount = OpenDatabase.ReadLine
                NewItem.HardCore = OpenDatabase.ReadLine
                NewItem.Ladder = OpenDatabase.ReadLine
                NewItem.Expansion = OpenDatabase.ReadLine
                NewItem.UserField = OpenDatabase.ReadLine
                NewItem.ItemImage = OpenDatabase.ReadLine
                NewItem.ImportTime = OpenDatabase.ReadLine
                NewItem.ImportDate = OpenDatabase.ReadLine
                ItemObjects.Add(NewItem)
            Loop
            OpenDatabase.Close()
            If CheckSpacerFlag <> "--------------------" Then Main.ErrorHandler(502, 0, 0, 0) 'Branch to error handler if item records in file are out of sync
            'Branch To Populate ListBox Routine If Load Worked Fine

            'CATCH ERRORS WHEN OPENING DATABASES AND PASS THEM TO THE ERROR HANDLER
        Catch ex As Exception
            OpenDatabase.Close()                                                        'Close the OpenDatabase Read File Channel
            Main.ErrorHandler(501, ex, CountRecordsForErrorEvents, My.Computer.FileSystem.GetName(DatabaseFilePath).ToString)                                                  'Branch to error handler on any other unexpected error
        End Try

        'Setup New Database On Main Form
        PopulateAllItemsLISTBOX()
        If ItemObjects.Count > 0 Then Main.AllItemsLISTBOX.SelectedIndex = 0

        Main.OpenDatabaseLABEL.Text = Replace(My.Computer.FileSystem.GetName(AppSettings.CurrentDatabase), ".txt", "")

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'POPULATE ALL ITEMS LISTBOX     - Clear out Old Items From AllItemsLISTBOX
    '                               - Iterates all items names in Item Database And appys them as items into AllItemsLISTBOX
    '                               - Update Item Tally Counter TEXTBOX
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub PopulateAllItemsLISTBOX()
        Main.AllItemsLISTBOX.Items.Clear()
        For ItemIndex = 0 To ItemObjects.Count - 1
            Main.AllItemsLISTBOX.Items.Add(ItemObjects(ItemIndex).ItemName)
        Next
        Main.ItemTallyTEXTBOX.Text = ItemObjects.Count & " - Items"

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'CLOSE FILE ROUTINE - clears Item Database and all main form controls
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub CloseFile()
        ItemObjects.Clear()

        Main.AllItemsLISTBOX.Items.Clear()
        Main.SearchLISTBOX.Items.Clear()
        Main.TradeListRICHTEXTBOX.Clear()

        Main.ClearItemStats()
        Main.OpenDatabaseLABEL.Text = ""

        Main.AllItemsLISTBOX.Select()
        Main.ListboxTABCONTROL.SelectTab(0)
        Main.ListControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        Main.SearchListControlTabBUTTON.BackgroundImage = Nothing
        Main.TradesListControlTabBUTTON.BackgroundImage = Nothing
        Main.UserRefControlTabBUTTON.BackgroundImage = Nothing
        Main.ItemTallyTEXTBOX.Text = Main.AllItemsLISTBOX.Items.Count & " - Items"
        Main.DatabaseFileLABEL.Hide()
        Main.DatabaseFileNameTEXTBOX.Hide()

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SAVES THE CURRENT DATABASE TO FILE ROUTINE 
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub SaveDatabase(DatabaseFile)
        Try
            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(DatabaseFile, False)
            For x = 0 To ItemObjects.Count - 1
                LogWriter.WriteLine("--------------------")
                LogWriter.WriteLine(ItemObjects(x).ItemName)
                LogWriter.WriteLine(ItemObjects(x).ItemBase)
                LogWriter.WriteLine(ItemObjects(x).ItemQuality)
                LogWriter.WriteLine(ItemObjects(x).RequiredCharacter)
                LogWriter.WriteLine(ItemObjects(x).EtherealItem)
                LogWriter.WriteLine(ItemObjects(x).Sockets)
                LogWriter.WriteLine(ItemObjects(x).RuneWord)
                LogWriter.WriteLine(ItemObjects(x).ThrowDamageMin)
                LogWriter.WriteLine(ItemObjects(x).ThrowDamageMax)
                LogWriter.WriteLine(ItemObjects(x).OneHandDamageMin)
                LogWriter.WriteLine(ItemObjects(x).OneHandDamageMax)
                LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMin)
                LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMax)
                LogWriter.WriteLine(ItemObjects(x).Defense)
                LogWriter.WriteLine(ItemObjects(x).ChanceToBlock)
                LogWriter.WriteLine(ItemObjects(x).QuantityMin)
                LogWriter.WriteLine(ItemObjects(x).QuantityMax)
                LogWriter.WriteLine(ItemObjects(x).DurabilityMin)
                LogWriter.WriteLine(ItemObjects(x).DurabilityMax)
                LogWriter.WriteLine(ItemObjects(x).RequiredStrength)
                LogWriter.WriteLine(ItemObjects(x).RequiredDexterity)
                LogWriter.WriteLine(ItemObjects(x).RequiredLevel)
                LogWriter.WriteLine(ItemObjects(x).AttackClass)
                LogWriter.WriteLine(ItemObjects(x).AttackSpeed)
                LogWriter.WriteLine(ItemObjects(x).Stat1)
                LogWriter.WriteLine(ItemObjects(x).Stat2)
                LogWriter.WriteLine(ItemObjects(x).Stat3)
                LogWriter.WriteLine(ItemObjects(x).Stat4)
                LogWriter.WriteLine(ItemObjects(x).Stat5)
                LogWriter.WriteLine(ItemObjects(x).Stat6)
                LogWriter.WriteLine(ItemObjects(x).Stat7)
                LogWriter.WriteLine(ItemObjects(x).Stat8)
                LogWriter.WriteLine(ItemObjects(x).Stat9)
                LogWriter.WriteLine(ItemObjects(x).Stat10)
                LogWriter.WriteLine(ItemObjects(x).Stat11)
                LogWriter.WriteLine(ItemObjects(x).Stat12)
                LogWriter.WriteLine(ItemObjects(x).Stat13)
                LogWriter.WriteLine(ItemObjects(x).Stat14)
                LogWriter.WriteLine(ItemObjects(x).Stat15)
                LogWriter.WriteLine(ItemObjects(x).MuleName)
                LogWriter.WriteLine(ItemObjects(x).MuleAccount)
                LogWriter.WriteLine(ItemObjects(x).MulePass)
                LogWriter.WriteLine(ItemObjects(x).PickitAccount)
                LogWriter.WriteLine(ItemObjects(x).HardCore)
                LogWriter.WriteLine(ItemObjects(x).Ladder)
                LogWriter.WriteLine(ItemObjects(x).Expansion)
                LogWriter.WriteLine(ItemObjects(x).UserField)
                LogWriter.WriteLine(ItemObjects(x).ItemImage)
                LogWriter.WriteLine(ItemObjects(x).ImportTime)
                LogWriter.WriteLine(ItemObjects(x).ImportDate)
            Next
            LogWriter.Close()

        Catch ex As Exception
            'ymessages = "File Write Error" : MyMessageBox()
        End Try

    End Sub


    '-----------------------------------------------------------------------------------------------------------------------------------------------
    ' To use call by using SaveLoggedItems(integer, name of file, Append true/false)
    '-----------------------------------------------------------------------------------------------------------------------------------------------
    Public Sub WriteToFile(ByVal itemstart, fName, bAppend)
        Dim CountRecordsForErrorReports As Integer = 0
        Main.ImportLogRICHTEXTBOX.AppendText("Saving to file " & fName & vbCrLf)

        Try
            Dim LogWriter = My.Computer.FileSystem.OpenTextFileWriter(fName, bAppend)
            For x = itemstart To ItemObjects.Count - 1
                CountRecordsForErrorReports = CountRecordsForErrorReports + 1
                LogWriter.WriteLine("--------------------")
                LogWriter.WriteLine(ItemObjects(x).ItemName)
                LogWriter.WriteLine(ItemObjects(x).Itemlevel)
                LogWriter.WriteLine(ItemObjects(x).ItemRealm)
                LogWriter.WriteLine(ItemObjects(x).ItemBase)
                LogWriter.WriteLine(ItemObjects(x).ItemQuality)
                LogWriter.WriteLine(ItemObjects(x).RequiredCharacter)
                LogWriter.WriteLine(ItemObjects(x).EtherealItem)
                LogWriter.WriteLine(ItemObjects(x).Sockets)
                LogWriter.WriteLine(ItemObjects(x).RuneWord)
                LogWriter.WriteLine(ItemObjects(x).ThrowDamageMin)
                LogWriter.WriteLine(ItemObjects(x).ThrowDamageMax)
                LogWriter.WriteLine(ItemObjects(x).OneHandDamageMin)
                LogWriter.WriteLine(ItemObjects(x).OneHandDamageMax)
                LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMin)
                LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMax)
                LogWriter.WriteLine(ItemObjects(x).Defense)
                LogWriter.WriteLine(ItemObjects(x).ChanceToBlock)
                LogWriter.WriteLine(ItemObjects(x).QuantityMin)
                LogWriter.WriteLine(ItemObjects(x).QuantityMax)
                LogWriter.WriteLine(ItemObjects(x).DurabilityMin)
                LogWriter.WriteLine(ItemObjects(x).DurabilityMax)
                LogWriter.WriteLine(ItemObjects(x).RequiredStrength)
                LogWriter.WriteLine(ItemObjects(x).RequiredDexterity)
                LogWriter.WriteLine(ItemObjects(x).RequiredLevel)
                LogWriter.WriteLine(ItemObjects(x).AttackClass)
                LogWriter.WriteLine(ItemObjects(x).AttackSpeed)
                LogWriter.WriteLine(ItemObjects(x).Stat1)
                LogWriter.WriteLine(ItemObjects(x).Stat2)
                LogWriter.WriteLine(ItemObjects(x).Stat3)
                LogWriter.WriteLine(ItemObjects(x).Stat4)
                LogWriter.WriteLine(ItemObjects(x).Stat5)
                LogWriter.WriteLine(ItemObjects(x).Stat6)
                LogWriter.WriteLine(ItemObjects(x).Stat7)
                LogWriter.WriteLine(ItemObjects(x).Stat8)
                LogWriter.WriteLine(ItemObjects(x).Stat9)
                LogWriter.WriteLine(ItemObjects(x).Stat10)
                LogWriter.WriteLine(ItemObjects(x).Stat11)
                LogWriter.WriteLine(ItemObjects(x).Stat12)
                LogWriter.WriteLine(ItemObjects(x).Stat13)
                LogWriter.WriteLine(ItemObjects(x).Stat14)
                LogWriter.WriteLine(ItemObjects(x).Stat15)
                LogWriter.WriteLine(ItemObjects(x).MuleName)
                LogWriter.WriteLine(ItemObjects(x).MuleAccount)
                LogWriter.WriteLine(ItemObjects(x).MulePass)
                LogWriter.WriteLine(ItemObjects(x).PickitAccount)
                LogWriter.WriteLine(ItemObjects(x).HardCore)
                LogWriter.WriteLine(ItemObjects(x).Ladder)
                LogWriter.WriteLine(ItemObjects(x).Expansion)
                LogWriter.WriteLine(ItemObjects(x).UserField)
                LogWriter.WriteLine(ItemObjects(x).ItemImage)
                LogWriter.WriteLine(ItemObjects(x).ImportTime)
                LogWriter.WriteLine(ItemObjects(x).ImportDate)
            Next
            LogWriter.Close()

        Catch ex As Exception
            Main.ErrorHandler(601, ex, CountRecordsForErrorReports, 0)

        End Try
        Main.ImportLogRICHTEXTBOX.AppendText("Items Saved = " & (ItemObjects.Count - itemstart) & vbCrLf)

    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'RENAME SELECTED DATABASE ROUTINE   - Selected file check is carried out before branching here
    '                                   - Sets up USer Input Form For User To Supply New File Name
    '                                   - Renames Selected Database File in DATABASES dir
    '                                   - Checks For current file and renames main form label
    '                                   - Checks if Backup file exists and renames it also
    '                                   -
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Sub RenameDatabase(DatabasePath)

        'Setup User Input Form For Use With Rename Database Function
        UserInput.Text = "Rename Selected Database"
        UserInput.UserInputHeaderLABEL.Text = "ENTER NEW DATABASE NAME"
        UserInput.UserInputMessageLABEL.Text = "Please type a new file name into the text box below to continue."
        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "Rename"
        UserInput.UserInputTEXTBOX.Text = DatabaseManager.ManagerDatabasesLISTBOX.SelectedItem
        UserInput.UserInputTEXTBOX.Select()

        UserInput.DatabaseManagerBorder1LABEL.Visible = True
        UserInput.DatabaseManagerBorder2LABEL.Visible = True
        UserInput.DatabaseManagerBorder3LABEL.Visible = True
        UserInput.DatabaseManagerBorder4LABEL.Visible = True
        UserInput.UserInputTEXTBOX.Visible = True


        'Gets New Database Filename With UserImputForm 
        Dim DialogResult = UserInput.ShowDialog
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        If DialogResult = Windows.Forms.DialogResult.Yes Then

            Try

                'Renames Database File HERE
                My.Computer.FileSystem.RenameFile(DatabasePath, UserInput.UserInputTEXTBOX.Text + ".txt")

                'Renames Current File Label On Main Form (If It Matches The Database Being Renamed)
                If Main.OpenDatabaseLABEL.Text = (Replace(My.Computer.FileSystem.GetName(DatabasePath), ".txt", "")) Then Main.OpenDatabaseLABEL.Text = UserInput.UserInputTEXTBOX.Text

                'Renames The Databases Backup File (If One Exists in Backup Dir)
                If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\Backup\" + Replace(My.Computer.FileSystem.GetName(DatabasePath), ".txt", "") + ".bak") = True Then
                    My.Computer.FileSystem.RenameFile(AppSettings.InstallPath + "\Databases\Backup\" + Replace(My.Computer.FileSystem.GetName(DatabasePath), ".txt", "") + ".bak", UserInput.UserInputTEXTBOX.Text + ".txt")
                End If

                'Renames Settings Default file if it matches
                If AppSettings.DefaultDatabase = DatabasePath Then AppSettings.DefaultDatabase = AppSettings.InstallPath + "\Databases\" + UserInput.UserInputTEXTBOX.Text + ".txt" : SaveSettingsFile()
                If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
            Catch ex As Exception
                Main.ErrorHandler(701, ex, 0, 0)
            End Try

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'CREATE BACKUP ROUTINE - Creates a backup of the supplied database file and saves it in the backup dir (if only they were all this small)
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub CreateBackup(DatabaseFile)
        If My.Computer.FileSystem.FileExists(DatabaseFile) = False Then  '
            'error handler
            Return
        End If

        Dim FileToBackup = My.Computer.FileSystem.GetName(DatabaseFile) '                                               remove path and assign filename to variable
        Dim backupFileName = AppSettings.InstallPath + "\Databases\Backup\" + Replace(FileToBackup, ".txt", ".bak") '   assign backup path to filename and add .bak extension
        If My.Computer.FileSystem.FileExists(backupFileName) = True Then  '
            My.Computer.FileSystem.DeleteFile(backupFileName)
        End If

        'Writes the new file by copying the current with an new .BAK extension
        My.Computer.FileSystem.CopyFile(DatabaseFile, backupFileName)
    End Sub
    '=====================================================================================================================================================================
End Module
