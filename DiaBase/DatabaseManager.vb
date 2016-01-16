Imports System.IO

'=========================================================================================================================================================================
'DATABASE MANAGER FORM - Handles All database manipulation such as open rename create and delete (may add save too tho it would be redundant)
'=========================================================================================================================================================================
Public Class DatabaseManager

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DATABASE MANAGER LOAD HANDLER  - Applys Diablo II Font To Headers And Buttons
    '                               - Gets all files from the InstallPath\Databases\ Directory 
    '                               - Populates List with Database File Names
    '                               - Populates Combobox DropDown List With Database File Names
    '                               - NOTE: Only includes files from InstallPath\Databases\ directory if they have a .TXT or .txt extension, others files are skipped
    '                               - Applys The current Database as the Default Database File Name Selected In the combobox
    '                               - Plays Audio D2 Donk When UnMuted
    '                               -
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DatabaseManager_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Left = Main.Left + 75 : Me.Top = Main.Top + 70         'Set Central window location


        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then
            'Apply Diablo II Header Font - 14 point
            ManagerHeaderLABEL.Font = New Font(pfc.Families(0), 14, FontStyle.Regular)

            'Apply Diablo II Function Buttons Font - 9 point
            ManagerRenameBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerDeleteBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerOpenBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerCancelBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerCreateBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerSummaryBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerRefreshBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ManagerMergeBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

        RefreshDatabaseList()
     

        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background) '                       Audio D2 Dong
    End Sub



    '---------------------------------------------------------------------------------------------------------------------------------------------
    'CLOSE DATABASE MANAGER BUTTON HANDLER - Simply closes the database manager form with a dong
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ManagerCancelBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerCancelBUTTON.Click, ManagerCancelBUTTON.Click
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        Me.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'RENAME DATABASE BUTTON HANDLER - Renames selected database filenames
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ManagerRenameBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerRenameBUTTON.Click
        If AutoLoggerRunning = False Then
            If DatabaseManagerSavedDatabasesLISTBOX.SelectedItem <> Nothing Then
                If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt") = True Then
                    If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
                    RenameDatabase(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt")
                End If

                RefreshDatabaseList()
            End If
        End If
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'DELETE DATABASE BUTTON HANDLER - Deletes selected database
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ManagerDeleteBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerDeleteBUTTON.Click
        If AutoLoggerRunning = False Then
            If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt") = True And Main.OpenDatabaseLABEL.Text <> DatabaseManagerSavedDatabasesLISTBOX.SelectedItem Then
                If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
                DeleteDatabase(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt")
                RefreshDatabaseList()
            End If
        End If
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'OPEN DATABSE BUTTON HANDLER       
    '                                   - Verify Selected File Exists
    '                                   - Verifys Selected File Is an actual DiaBase Database File
    '                                   - Check Save Current File Checkbox and apply save if nessicary
    '                                   - 
    '                                   - 
    '                                   - 
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ManagerOpenBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerOpenBUTTON.Click

        If AutoLoggerRunning = False And DatabaseManagerSavedDatabasesLISTBOX.SelectedIndex > -1 Then
            If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt") = True Then

                OpenDatabase(AppSettings.InstallPath + "\Databases\" + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".txt") 'Branch To Open Database Routine
                Me.Close()
            Else
                Main.ErrorHandler(401, 0, 0, 0)       'File Does Not Exist Error Branch To Error Handler

            End If
        End If
        If ItemObjects.Count > 0 Then Main.AllItemsLISTBOX.SelectedIndex = 0
        Me.Close()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'DATABASE MANAGER REFRESH BUTTON HANDLER - Branches to Update Saved Database Listbox routine On Darabase Manager Form
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ManagerRefreshBUTTON.Click
        If AutoLoggerRunning = False Then RefreshDatabaseList()
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------
    'REFRESH DATABASE ROUTINE - gets all saved database filenames and populates database  managaer listbox with file names
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Sub RefreshDatabaseList()
        'Get All Database Files From InstallPath\Databases Directory
        Dim AllSavedDatabaseFileNames As Array = Nothing
        AllSavedDatabaseFileNames = Directory.GetFiles(AppSettings.InstallPath + "\Databases\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        'Populate Listbox items and Combobox Drop Down List items
        DatabaseManagerSavedDatabasesLISTBOX.Items.Clear() '                                                                             Delete old database file name lists
        For Each DatabaseFileName In AllSavedDatabaseFileNames
            If DatabaseFileName.indexof(".txt") > -1 Then '                        Check for correct .TXT or .txt file extension
                If DatabaseFileName.indexof(".txt") > -1 Then DatabaseFileName = Replace(DatabaseFileName, ".txt", "") '    Remove lower case extension if there is one
                DatabaseManagerSavedDatabasesLISTBOX.Items.Add(DatabaseFileName)  '                                                      Apply Cropped Database File Name to lists
            End If
        Next
    End Sub


    '---------------------------------------------------------------------------------------------------------------------------------------------
    'CREATE NEW DATABASE FILE ROUTINE
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ManagerCreateBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerCreateBUTTON.Click
        Dim OnExistsError As Boolean = False 'Used to display a file exist error when looping to display a secong textbox to supply another file name (this way one display form routine handles both jobs)

FileExistsErrorLoop:

        'Setup User Input Form For Use With The Create New Database Function
        UserInput.Text = "Create New Database File"
        UserInput.UserInputHeaderLABEL.Text = "ENTER UNIQUE DATABASE NAME"
        If OnExistsError = False Then
            UserInput.UserInputTEXTBOX.Text = Nothing
            UserInput.UserInputMessageLABEL.Text = "To create a new database please type a unique file name into the text box below and select CREATE to continue." + vbCrLf + vbCrLf + " Do not include file path or file extension. DiaBase will manage these for you."
        End If
        If OnExistsError = True Then
            UserInput.UserInputMessageLABEL.Text = "THE SUPPLIED FILE NAME ALREADY EXISTS..." + vbCrLf + vbCrLf + "To create a new database please type a UNIQUE file name into the text box below and select CREATE to continue." + vbCrLf + vbCrLf + " Do not include file path or file extension."

        End If

        OnExistsError = False

        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "Create"


        UserInput.DatabaseManagerBorder1LABEL.Visible = True
        UserInput.DatabaseManagerBorder2LABEL.Visible = True
        UserInput.DatabaseManagerBorder3LABEL.Visible = True
        UserInput.DatabaseManagerBorder4LABEL.Visible = True
        UserInput.UserInputTEXTBOX.Visible = True

        UserInput.UserInputTEXTBOX.SelectionStart = 0 : UserInput.UserInputTEXTBOX.SelectionLength = Len(UserInput.UserInputTEXTBOX.Text)
        UserInput.UserInputTEXTBOX.Select()



        'Gets New Database Filename With UserImputForm 
        Dim DialogResult = UserInput.ShowDialog
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        If DialogResult = Windows.Forms.DialogResult.Yes Then

            If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\" + UserInput.UserInputTEXTBOX.Text + ".txt") = True Then
                OnExistsError = True
                'CREATES THE NEW DATABASE FILE RIGHT HERE - New file is simple an empty text file with the first items spacer included as the top line of the file
            Else
                Dim CreateDatabase As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(AppSettings.InstallPath + "\Databases\" + UserInput.UserInputTEXTBOX.Text + ".txt", False)
                CreateDatabase.Close()
            End If
        End If

        If OnExistsError = True Then GoTo FileExistsErrorLoop Else RefreshDatabaseList()


    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'DELETE DATABASE ROUTINE - Removes the selected database file
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub DeleteDatabase(DatabaseFile)

        'Setup User Input Form For Use With Confirm Delete Database Function
        UserInput.Text = "Delete Selected Database"
        UserInput.UserInputHeaderLABEL.Text = "DELETE DATABASE CONFIRMATION"
        UserInput.UserInputMessageLABEL.Text = "DATA LOSS WARNING:" + vbCrLf + vbCrLf + "You are about to delete the " + Chr(34) + DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + Chr(34) + " database file and its associated backup file." + vbCrLf + vbCrLf + " Please select DELETE to continue or CANCEL to abort."
        UserInput.UserInputNoBUTTON.Text = "Cancel"
        UserInput.UserInputYesBUTTON.Text = "Delete"

        'Hide Unused Controls
        UserInput.DatabaseManagerBorder1LABEL.Visible = False
        UserInput.DatabaseManagerBorder2LABEL.Visible = False
        UserInput.DatabaseManagerBorder3LABEL.Visible = False
        UserInput.DatabaseManagerBorder4LABEL.Visible = False
        UserInput.UserInputTEXTBOX.Visible = False

        'ON YES CONFIRMATION deletes file here
        Dim DialogResult = UserInput.ShowDialog
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
        If DialogResult = Windows.Forms.DialogResult.Yes Then

            Try

                'Database Dir
                My.Computer.FileSystem.DeleteFile(DatabaseFile)

                'Backup Dir
                If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\Backup\" + Replace(My.Computer.FileSystem.GetName(DatabaseFile), ".txt", ".bak")) = True Then
                    My.Computer.FileSystem.DeleteFile(AppSettings.InstallPath + "\Databases\Backup\" + Replace(My.Computer.FileSystem.GetName(DatabaseFile), ".txt", ".bak"))
                End If

            Catch ex As Exception
                Main.ErrorHandler(801, ex, 0, 0)
            End Try

        End If
    End Sub

    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'BRANCH TO DATABASE INFO FORM ON INFO BUTTON PRESS - Evaluation Routines are Within the Database Info Forms module
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ManagerSummaryBUTTON_Click(sender As Object, e As EventArgs) Handles ManagerSummaryBUTTON.Click

        'Branch to DatabaseInfo Form & routines if a database is selected if not do nothing
        If DatabaseManagerSavedDatabasesLISTBOX.SelectedIndex <> -1 Then
            DatabaseInfo.ClearOldData()
            DatabaseInfo.Show()
            DatabaseInfo.DatabaseInfoTABCONTROL.SelectTab(0)
            DatabaseInfo.DatabaseInfoCloseBUTTON.Select()
        End If


    End Sub
End Class