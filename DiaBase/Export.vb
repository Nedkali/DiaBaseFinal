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
        SavedDatabasesLISTBOX.Items.Clear() '                                                                               Delete old database file name lists
        For Each DatabaseFileName In AllSavedDatabaseFileNames
            If DatabaseFileName.indexof(".txt") > -1 And DatabaseFileName <> temp1 Then '                                   Check for correct .TXT or .txt file extension
                If DatabaseFileName.indexof(".txt") > -1 Then DatabaseFileName = Replace(DatabaseFileName, ".txt", "") '    Remove lower case extension if there is one
                SavedDatabasesLISTBOX.Items.Add(DatabaseFileName)  '                                                        Apply Cropped Database File Name to lists
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

        'Dim x As Integer = 0
        'Dim Temp = DatabaseFilenameTEXTBOX.Text : Temp = AppSettings.InstallPath + "\Databases\" + Temp + ".txt"

        'Run a backup of current database if backup before edits is set to true in app settings class and form
        If AppSettings.BackupBeforeEdits = True Then DatabaseManagmentFunctions.SaveDatabase(AppSettings.InstallPath & "\Databases\Backup\" & Main.OpenDatabaseLABEL.Text & ".BAK")

        Try
            'Dim LogWriter As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(temp, True)
            ' For index = 0 To Main.UserLISTBOX.SelectedIndices.Count - 1
            'x = UserListReferenceList(index)
            'LogWriter.WriteLine("--------------------")
            'LogWriter.WriteLine(ItemObjects(x).ItemName)
            'LogWriter.WriteLine(ItemObjects(x).Itemlevel)
            'LogWriter.WriteLine(ItemObjects(x).ItemRealm)
            'LogWriter.WriteLine(ItemObjects(x).ItemBase)
            'LogWriter.WriteLine(ItemObjects(x).ItemQuality)
            'LogWriter.WriteLine(ItemObjects(x).RequiredCharacter)
            'LogWriter.WriteLine(ItemObjects(x).EtherealItem)
            'LogWriter.WriteLine(ItemObjects(x).Sockets)
            'LogWriter.WriteLine(ItemObjects(x).RuneWord)
            'LogWriter.WriteLine(ItemObjects(x).ThrowDamageMin)
            'LogWriter.WriteLine(ItemObjects(x).ThrowDamageMax)
            'LogWriter.WriteLine(ItemObjects(x).OneHandDamageMin)
            'LogWriter.WriteLine(ItemObjects(x).OneHandDamageMax)
            'LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMin)
            'LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMax)
            'LogWriter.WriteLine(ItemObjects(x).Defense)
            'LogWriter.WriteLine(ItemObjects(x).ChanceToBlock)
            'LogWriter.WriteLine(ItemObjects(x).QuantityMin)
            'LogWriter.WriteLine(ItemObjects(x).QuantityMax)
            'LogWriter.WriteLine(ItemObjects(x).DurabilityMin)
            'LogWriter.WriteLine(ItemObjects(x).DurabilityMax)
            'LogWriter.WriteLine(ItemObjects(x).RequiredStrength)
            'LogWriter.WriteLine(ItemObjects(x).RequiredDexterity)
            'LogWriter.WriteLine(ItemObjects(x).RequiredLevel)
            'LogWriter.WriteLine(ItemObjects(x).AttackClass)
            'LogWriter.WriteLine(ItemObjects(x).AttackSpeed)
            'LogWriter.WriteLine(ItemObjects(x).Stat1)
            'LogWriter.WriteLine(ItemObjects(x).Stat2)
            'LogWriter.WriteLine(ItemObjects(x).Stat3)
            'LogWriter.WriteLine(ItemObjects(x).Stat4)
            'LogWriter.WriteLine(ItemObjects(x).Stat5)
            'LogWriter.WriteLine(ItemObjects(x).Stat6)
            'LogWriter.WriteLine(ItemObjects(x).Stat7)
            'LogWriter.WriteLine(ItemObjects(x).Stat8)
            'LogWriter.WriteLine(ItemObjects(x).Stat9)
            'LogWriter.WriteLine(ItemObjects(x).Stat10)
            'LogWriter.WriteLine(ItemObjects(x).Stat11)
            'LogWriter.WriteLine(ItemObjects(x).Stat12)
            'LogWriter.WriteLine(ItemObjects(x).Stat13)
            'LogWriter.WriteLine(ItemObjects(x).Stat14)
            'LogWriter.WriteLine(ItemObjects(x).Stat15)
            'LogWriter.WriteLine(ItemObjects(x).MuleName)
            'LogWriter.WriteLine(ItemObjects(x).MuleAccount)
            'LogWriter.WriteLine(ItemObjects(x).MulePass)
            'LogWriter.WriteLine(ItemObjects(x).PickitAccount)
            'LogWriter.WriteLine(ItemObjects(x).HardCore)
            'LogWriter.WriteLine(ItemObjects(x).Ladder)
            'LogWriter.WriteLine(ItemObjects(x).Expansion)
            'LogWriter.WriteLine(ItemObjects(x).UserField)
            'LogWriter.WriteLine(ItemObjects(x).ItemImage)
            'LogWriter.WriteLine(ItemObjects(x).ImportTime)
            'LogWriter.WriteLine(ItemObjects(x).ImportDate)
            'Next
            'LogWriter.Close()


            'Delete imported items from old dtatabase if checkbox IS NOT checked
            If DontDeleteItemsCHECKBOX.Checked = False Then

                Dim ItemIndex As Integer = 0
                Dim CheckResult As Boolean = Nothing

                For Each UserIndex In Main.UserLISTBOX.SelectedIndices
                    If Main.UserLISTBOX.Items.Item(UserIndex) = ItemObjects(ItemIndex).ItemName Then VerifyStats(ItemIndex, UserIndex, CheckResult)
                    If CheckResult = True Then Main.UserLISTBOX.Items.RemoveAt(ItemIndex)
                    ItemIndex = ItemIndex + 1

                Next

                '  Dim a As Integer
                '  Dim b As Integer
                '  For index = Main.AllItemsLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                '  a = Main.UserLISTBOX.SelectedIndices(index)
                '  b = UserListReferenceList(a)
                '  Main.UserLISTBOX.Items.RemoveAt(a)
                '  Main.AllItemsLISTBOX.Items.RemoveAt(b)
                '  ItemObjects.RemoveAt(b)
                '  SearchReferenceList.RemoveAt(a)
                '  Next
                Main.SearchLISTBOX.SelectedItem = -1
            End If

            'checks open destination database checkbox then save this database and load the destionation database if nessicary
            If OpenDatabaseCHECKBOX.Checked = True Then
                DatabaseManagmentFunctions.SaveDatabase(AppSettings.CurrentDatabase) 'branch to save routine to save source dbase before loading destination dbase NOTE: I MAY PUT A CHECKBOX IN FOR THIS

                'clean out old items from the last loaded database
                DatabaseManagmentFunctions.CloseFile()

                'load up destination one and display first item and set current database label in top right of form1
                AppSettings.CurrentDatabase = Application.StartupPath + "\Databases\" + DatabaseFilenameTEXTBOX.Text + ".txt"
                DatabaseManagmentFunctions.OpenDatabase(AppSettings.CurrentDatabase)

                Me.Close()
                Main.ListboxTABCONTROL.SelectTab(0)
                Main.ListControlTabBUTTON.BackColor = Color.DimGray
                Main.SearchListControlTabBUTTON.BackColor = Color.Black
                Main.TradesListControlTabBUTTON.BackColor = Color.Black
                Main.ItemTallyTEXTBOX.Text = Main.AllItemsLISTBOX.Items.Count & " - Total Items"

                'ensure dabase containing filename field is hidden for re-focusing on main list box
                Main.DatabaseFileLABEL.Hide()
                Main.DatabaseFileNameTEXTBOX.Hide()
                If Main.AllItemsLISTBOX.Items.Count > 0 Then Main.AllItemsLISTBOX.SelectedIndex = 0
            End If

            Return
        Catch ex As Exception
            'Pass Failed import error instance to error hanler 
            Main.ErrorHandler(1001, ex, 0, 0) '1001 - 1099 Haldles all expected importing errors (if needed)
            Return
        End Try

        Me.Close()



    End Sub


    'USED BY REMOVE EXPORTED ITEMS FROM SOURCE DATABASE ROUTINE - checks item to delete from database is the exact item by verifyings all stats are the same
    Sub VerifyStats(ItemIndex, UserIndex, CheckResult)

        CheckResult = True

        For Each Item In ItemObjects

            If ItemObjects(ItemIndex).Itemlevel <> UserObjects(UserIndex).Itemlevel Then CheckResult = False
            If ItemObjects(ItemIndex).ItemRealm <> UserObjects(UserIndex).ItemRealm Then CheckResult = False
            If ItemObjects(ItemIndex).ItemBase <> UserObjects(UserIndex).ItemBase Then CheckResult = False
            If ItemObjects(ItemIndex).ItemQuality <> UserObjects(UserIndex).ItemQuality Then CheckResult = False
            If ItemObjects(ItemIndex).RequiredCharacter <> UserObjects(UserIndex).RequiredCharacter Then CheckResult = False
            If ItemObjects(ItemIndex).EtherealItem <> UserObjects(UserIndex).EtherealItem Then CheckResult = False
            If ItemObjects(ItemIndex).Sockets <> UserObjects(UserIndex).Sockets Then CheckResult = False
            If ItemObjects(ItemIndex).RuneWord <> UserObjects(UserIndex).RuneWord Then CheckResult = False
            If ItemObjects(ItemIndex).ThrowDamageMin <> UserObjects(UserIndex).ThrowDamageMin Then CheckResult = False
            If ItemObjects(ItemIndex).ThrowDamageMax <> UserObjects(UserIndex).ThrowDamageMax Then CheckResult = False
            If ItemObjects(ItemIndex).OneHandDamageMin <> UserObjects(UserIndex).OneHandDamageMin Then CheckResult = False
            If ItemObjects(ItemIndex).OneHandDamageMax <> UserObjects(UserIndex).OneHandDamageMax Then CheckResult = False
            If ItemObjects(ItemIndex).TwoHandDamageMin <> UserObjects(UserIndex).TwoHandDamageMin Then CheckResult = False
            If ItemObjects(ItemIndex).TwoHandDamageMax <> UserObjects(UserIndex).TwoHandDamageMax Then CheckResult = False
            If ItemObjects(ItemIndex).Defense <> UserObjects(UserIndex).Defense Then CheckResult = False
            If ItemObjects(ItemIndex).ChanceToBlock <> UserObjects(UserIndex).ChanceToBlock Then CheckResult = False
            If ItemObjects(ItemIndex).QuantityMin <> UserObjects(UserIndex).QuantityMin Then CheckResult = False
            If ItemObjects(ItemIndex).QuantityMax <> UserObjects(UserIndex).QuantityMax Then CheckResult = False
            If ItemObjects(ItemIndex).DurabilityMin <> UserObjects(UserIndex).DurabilityMin Then CheckResult = False
            If ItemObjects(ItemIndex).DurabilityMax <> UserObjects(UserIndex).DurabilityMax Then CheckResult = False
            If ItemObjects(ItemIndex).RequiredStrength <> UserObjects(UserIndex).RequiredStrength Then CheckResult = False
            If ItemObjects(ItemIndex).RequiredDexterity <> UserObjects(UserIndex).RequiredDexterity Then CheckResult = False
            If ItemObjects(ItemIndex).RequiredLevel <> UserObjects(UserIndex).RequiredLevel Then CheckResult = False
            If ItemObjects(ItemIndex).AttackClass <> UserObjects(UserIndex).AttackClass Then CheckResult = False
            If ItemObjects(ItemIndex).AttackSpeed <> UserObjects(UserIndex).AttackSpeed Then CheckResult = False
            If ItemObjects(ItemIndex).Stat1 <> UserObjects(UserIndex).Stat1 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat2 <> UserObjects(UserIndex).Stat2 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat3 <> UserObjects(UserIndex).Stat3 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat4 <> UserObjects(UserIndex).Stat4 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat5 <> UserObjects(UserIndex).Stat5 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat6 <> UserObjects(UserIndex).Stat6 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat7 <> UserObjects(UserIndex).Stat7 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat8 <> UserObjects(UserIndex).Stat8 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat9 <> UserObjects(UserIndex).Stat9 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat10 <> UserObjects(UserIndex).Stat10 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat11 <> UserObjects(UserIndex).Stat11 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat12 <> UserObjects(UserIndex).Stat12 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat13 <> UserObjects(UserIndex).Stat13 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat14 <> UserObjects(UserIndex).Stat14 Then CheckResult = False
            If ItemObjects(ItemIndex).Stat15 <> UserObjects(UserIndex).Stat15 Then CheckResult = False
            If ItemObjects(ItemIndex).MuleName <> UserObjects(UserIndex).MuleName Then CheckResult = False
            If ItemObjects(ItemIndex).MuleAccount <> UserObjects(UserIndex).MuleAccount Then CheckResult = False
            If ItemObjects(ItemIndex).MulePass <> UserObjects(UserIndex).MulePass Then CheckResult = False
            If ItemObjects(ItemIndex).PickitAccount <> UserObjects(UserIndex).PickitAccount Then CheckResult = False
            If ItemObjects(ItemIndex).HardCore <> UserObjects(UserIndex).HardCore Then CheckResult = False
            If ItemObjects(ItemIndex).Ladder <> UserObjects(UserIndex).Ladder Then CheckResult = False
            If ItemObjects(ItemIndex).Expansion <> UserObjects(UserIndex).Expansion Then CheckResult = False
            If ItemObjects(ItemIndex).UserField <> UserObjects(UserIndex).UserField Then CheckResult = False
            If ItemObjects(ItemIndex).ItemImage <> UserObjects(UserIndex).ItemImage Then CheckResult = False
            If ItemObjects(ItemIndex).ImportTime <> UserObjects(UserIndex).ImportTime Then CheckResult = False
            If ItemObjects(ItemIndex).ImportDate <> UserObjects(UserIndex).ImportDate Then CheckResult = False
            If CheckResult = False Then Return
        Next
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