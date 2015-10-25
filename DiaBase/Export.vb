Imports System.IO

Public Class Export
    Private Sub Export_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetFileList()
    End Sub

    Private Sub GetFileList()
        Dim temp = AppSettings.CurrentDatabase.Split("\")
        Dim temp1 = temp(temp.Length - 1)
        'Get All Database Files From InstallPath\Databases Directory
        Dim AllSavedDatabaseFileNames As Array = Nothing
        AllSavedDatabaseFileNames = Directory.GetFiles(AppSettings.InstallPath + "\Databases\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        'Populate Listbox items and Combobox Drop Down List items
        SavedDatabasesLISTBOX.Items.Clear() '                                                                             Delete old database file name lists
        For Each DatabaseFileName In AllSavedDatabaseFileNames
            If DatabaseFileName.indexof(".txt") > -1 And DatabaseFileName <> temp1 Then '                        Check for correct .TXT or .txt file extension
                If DatabaseFileName.indexof(".txt") > -1 Then DatabaseFileName = Replace(DatabaseFileName, ".txt", "") '    Remove lower case extension if there is one
                SavedDatabasesLISTBOX.Items.Add(DatabaseFileName)  '                                                      Apply Cropped Database File Name to lists
            End If
        Next
    End Sub

    'EXPORT ITEMS TO SECOND DATABASE ROUTINE

    Private Sub ExportButton_Click(sender As Object, e As EventArgs) Handles ExportNowBUTTON.Click
        If Main.UserLISTBOX.SelectedIndices.Count < 1 Then Return

        'Just incase entry includes path and or file extension this should successfully take them back out
        If DatabaseFilenameTEXTBOX.Text <> "" Then
            My.Computer.FileSystem.GetName(DatabaseFilenameTEXTBOX.Text)
            DatabaseFilenameTEXTBOX.Text = Replace(DatabaseFilenameTEXTBOX.Text, ".txt", "")

        End If

        Dim x As Integer = 0
            Dim temp = DatabaseFilenameTEXTBOX.Text

            temp = AppSettings.InstallPath + "\Databases\" + temp + ".txt"

        'Create database file if it doesnt already exist yet
        If My.Computer.FileSystem.FileExists(temp) = False Then
            Dim CreateFile As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(AppSettings.InstallPath + "\DataBases\" & DatabaseFilenameTEXTBOX.Text & " .txt", False)
            CreateFile.Close()
        End If

        'Run a backup of current database if backup before edits is set to true in app settings class and form
        If AppSettings.BackupBeforeEdits = True Then DatabaseManagmentFunctions.SaveDatabase(AppSettings.InstallPath & "\Databases\Backup\" & Main.OpenDatabaseLABEL.Text & ".BAK")


            Try
                Dim LogWriter As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(temp, True)

            For index = 0 To Main.UserLISTBOX.SelectedIndices.Count - 1
                x = UserListReferenceList(index)
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

            'Delete imported items from old dtatabase if checkbox IS NOT checked
            If DontDeleteItemsCHECKBOX.Checked = False Then
                    Dim a As Integer
                    Dim b As Integer
                    For index = Main.SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                        a = Main.UserLISTBOX.SelectedIndices(index)
                        b = UserListReferenceList(a)
                        Main.UserLISTBOX.Items.RemoveAt(a)
                        Main.AllItemsLISTBOX.Items.RemoveAt(b)
                        ItemObjects.RemoveAt(b)
                        SearchReferenceList.RemoveAt(a)
                    Next
                    Main.SearchLISTBOX.SelectedItem = -1
                End If



                Return
            Catch ex As Exception
                'Pass Failed import error instance to error hanler 
                Main.ErrorHandler(1001, ex, 0, 0) '1001 - 1099 Haldles all expected importing errors (if needed)
                Return
            End Try



        'checks open destination database checkbox then save this database and load the destionation database if nessicary
        If OpenDatabaseCHECKBOX.Checked = True Then
            DatabaseManagmentFunctions.SaveDatabase(AppSettings.CurrentDatabase) 'branch to save routine to save source dbase before loading destination dbase NOTE: I MAY PUT A CHECKBOX IN FOR THIS

            'clean out old items from the last loaded database
            DatabaseManagmentFunctions.CloseFile()

            'load up destination one and display first item and set current database label in top right of form1
            AppSettings.CurrentDatabase = Application.StartupPath + "\Database\" + DatabaseFilenameTEXTBOX.Text + ".txt"
            DatabaseManagmentFunctions.OpenDatabase(AppSettings.CurrentDatabase)


            Main.ListboxTABCONTROL.SelectTab(0)
            Main.ListControlTabBUTTON.BackColor = Color.DimGray
            Main.SearchListControlTabBUTTON.BackColor = Color.Black
            Main.TradesListControlTabBUTTON.BackColor = Color.Black
            Main.ItemTallyTEXTBOX.Text = Main.AllItemsLISTBOX.Items.Count & " - Total Items"

            'ensure dabase containing filename field is hidden for re-focusing on main list box
            Main.DatabaseFileLABEL.Hide()
            Main.DatabaseFileNameTEXTBOX.Hide()



        End If



    End Sub






    'Updates textbox with selected file and keeps textbox selected ready for new filename 
    Private Sub SavedDatabasesLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SavedDatabasesLISTBOX.SelectedIndexChanged
        If SavedDatabasesLISTBOX.SelectedIndex <> -1 Then
            DatabaseFilenameTEXTBOX.Text = SavedDatabasesLISTBOX.SelectedItem
            DatabaseFilenameTEXTBOX.Select()

        End If
    End Sub

    'Keeps Textbox selected (ready for new filename) after interaction with the database listbox 
    Private Sub SavedDatabasesLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles SavedDatabasesLISTBOX.MouseDown
        DatabaseFilenameTEXTBOX.Select()
        DatabaseFilenameTEXTBOX.SelectionLength = 0

    End Sub

    Private Sub MoveItemsCancelBUTTON_Click(sender As Object, e As EventArgs) Handles ExportCancelBUTTON.Click
        Me.Close()
    End Sub

    'Auto select filename textbox
    Private Sub Export_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        DatabaseFilenameTEXTBOX.Select()
    End Sub
End Class