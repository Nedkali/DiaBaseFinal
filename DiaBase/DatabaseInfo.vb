Public Class DatabaseInfo

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SETUP LOCAL VARS AND ARRAY LISTS
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    Dim Count As Integer '                                          Index Tracker Var, Used Within FOR/NEXT And DO/LOOP Routines
    Dim Temp As String '                                            Temporary memory location For easier use of raw file data
    Dim ItemBaseList As List(Of String) = New List(Of String) '     List of all item bases
    Dim ItemBaseGroups As List(Of String) = New List(Of String) '   List of all item bases with duplicates removed
    Dim ItemBaseValues As List(Of String) = New List(Of String) '   List of  tallied percentile ratios for each item base group found
    Dim TotalItemsInSelected As Integer = 0 '                       Total Items Tally Var, is mainly used for header tally

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Show Handler - Used to clear out old data and run evaluation on forms first open
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DatabaseInfo_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        'Call Evaluation Routines for first run...

        ClearOldData() 'Delete form of old data routine
        GetItemTotal() 'Tally all items in selected             (for header)
        GetItemBases() 'Base Values For Each Item In Database   (for header)


        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        'UPDATES HEADER INFORMATION - Database Filename and Total Items
        '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        If DatabaseManager.DatabaseManagerSavedDatabasesLISTBOX.SelectedIndex <> -1 Then
            DatabaseInfoSelectedTEXTBOX.Text = DatabaseManager.DatabaseManagerSavedDatabasesLISTBOX.SelectedItem ' Database filename
            DatabaseInfoTotalTEXTBOX.Text = TotalItemsInSelected
        End If
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Clear out old data rows and textboxes from last selected database evaluation
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub ClearOldData()
        DatabaseInfoDATAGRIDVIEW.Rows.Clear()
        DatabaseInfoSelectedTEXTBOX.Text = Nothing
        DatabaseInfoTotalTEXTBOX.Text = Nothing
    End Sub

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Count all items in selected database
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub GetItemTotal()
        Dim CountItems = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Databases\" + DatabaseManager.DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".TXT")
        Do Until CountItems.EndOfStream
            If CountItems.ReadLine = "--------------------" Then TotalItemsInSelected = TotalItemsInSelected + 1
        Loop
        CountItems.Close()
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'GETS PERCENTAGE RATIOS AND TALLYS OF ALL ITEM BASES MENTIONED IN SELECTED DATABASE
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub GetItemBases()

        'Want to display info form as a modal dialog box to allow the selected database to be changed from the manager form. Then have
        'this form update data as a new database is selected. Want both forms to work together if i can to avoid writing a second
        'database selector in the info form, or the need to shuffle back and foward betewwn the two forms to evaluate more than
        'one database at a time...
        'Should be able to even set it up to work with the currently loaded database in real time - eg, as database changes the info
        'form could update as the selected database is saved. will need to add a file changed check on the info routine to trigger
        ' the forms data to update

        'START - Open Database File. Uses Database Manager Listbox Selected item as filename selector.
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Count = 0
        Dim CheckBases = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Databases\" + DatabaseManager.DatabaseManagerSavedDatabasesLISTBOX.SelectedItem + ".TXT")
        CheckBases.ReadLine() '                                                                     Skip Checkflag / Record Seperator Line

        'Read All Item Base Fields From The Selected Databases File
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Do Until CheckBases.EndOfStream
            CheckBases.ReadLine() : CheckBases.ReadLine() : CheckBases.ReadLine() '                 Align to Item Base Field as next field to be read - skip first three fields
            Temp = CheckBases.ReadLine() : ItemBaseList.Add(Temp) '                                 Read file and assign to array
            Count = 0 : Do Until Count = 49 : CheckBases.ReadLine() : Count = Count + 1 : Loop '    Align to next items item base field - skips 48 fields and checkflag
        Loop
        CheckBases.Close()

        'Seperates All Item Base Entries to a list of single entries - Removes dupes by creating new list. Old list now used for ratio
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        For Each item In ItemBaseList
            If ItemBaseGroups.Contains(item) = False Then ItemBaseGroups.Add(item)
        Next

        'Get Tallys for each item base listed.                      Ratio is displayed as a percentage of all items in the database
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        For Each ItemGroup In ItemBaseGroups '                      Each Individual Group
            Count = 0
            For Each ItemBase In ItemBaseList '                     In all All Collated Groups
                If ItemGroup = ItemBase Then Count = Count + 1 '    Total Tally Counter
            Next
            ItemBaseValues.Add(Count)
        Next

        'FINISH - Update Item Base Data Grid Layout Page.           Print compiled data to screen
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Count = 0
        For Each item In ItemBaseGroups
            DatabaseInfoDATAGRIDVIEW.Rows.Add(item, ItemBaseValues(Count), ItemBaseValues(Count) / ItemBaseList.Count * 100)
            Count = Count + 1
        Next
    End Sub

    'Closes Database Info Form
    Private Sub DatabaseInfoCloseBUTTON_Click(sender As Object, e As EventArgs) Handles DatabaseInfoCloseBUTTON.Click
        Me.Close()
    End Sub
End Class