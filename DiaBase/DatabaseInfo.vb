Public Class DatabaseInfo

    Private Sub DatabaseInfo_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        Dim Count As Integer = 0 '                                      Index Tracker Var, Used Within FOR/NEXT And DO/LOOP Routines
        Dim Temp As String = Nothing '                                  Temporary memory location For easier manipulation of file data
        Dim ItemBaseList As List(Of String) = New List(Of String) '     List of each items item base value
        Dim ItemBaseGroups As List(Of String) = New List(Of String) '   List of each item base found in the selected database
        Dim ItemBaseValues As List(Of String) = New List(Of String) '   List of tallied ratios for each item base group (ThisBase / AllBases * 100)

        DatabaseInfoDATAGRIDVIEW.Rows.Clear() '                         Clear out old data rows from last selected database evaluation



        'Call Evaluation Routines for first run

        GetItemBases()


    End Sub



    '==============================================================================================GETS PERCENTAGE RATIOS AND TALLYS OF ALL ITEM BASES MENTIONED IN SELECTED DATABASE===
    Sub GetItemBases()

        '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        'Want to display info form as a modal dialog box to allow the selected database to be changed from the manager form. Then have
        'this form update data as a new database is selected. Want both forms to work together if i can to avoid writing a second
        'database selector in the info form, or the need to shuffle back and foward betewwn the two forms to evaluate more than
        'one database at a time...
        'Should be able to even set it up to work with the currently loaded database in real time - eg, as database changes the info
        'form could update as the selected database is saved. will need to add a file changed check on the info routine to trigger
        ' the forms data to update

        '---COLLATE ITEM BASE INFORMATION-----------------------------------------------------------------------------------------------------------------------------------------------    

        'sub ItemBaseRatios:



        'Setup Working Vars And Define Mewmory Arrays (lists Of T)
        Dim Count As Integer '                                          Index Tracker Var, Used Within FOR/NEXT And DO/LOOP Routines
        Dim Temp As String '                                            Temporary memory location For easier use of raw file data
        Dim ItemBaseList As List(Of String) = New List(Of String) '     List of T - of all item bases
        Dim ItemBaseGroups As List(Of String) = New List(Of String) '   List of T - of all item bases with duper removed
        Dim ItemBaseValues As List(Of String) = New List(Of String) '   List of T - of tallied ratios for each item base group

        Count = 0 : DatabaseInfoDATAGRIDVIEW.Rows.Clear() 'Cear out old data rows from last selected database

        'START - Open Database File. Uses Database Manager Listbox Selected item as filename selector.

        Dim CheckBases = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Databases\" + DatabaseManager.ManagerDatabasesLISTBOX.SelectedItem + ".TXT")
        CheckBases.ReadLine() '                                                                     Skip Checkflag / Record Seperator Line

        'Read All Item Base Fields From The Selected Databases File

        Do Until CheckBases.EndOfStream
            CheckBases.ReadLine() : CheckBases.ReadLine() : CheckBases.ReadLine() '                 Align to Item Base Field as next field to be read - skip first three fields
            Temp = CheckBases.ReadLine() : ItemBaseList.Add(Temp)
            Count = 0 : Do Until Count = 49 : CheckBases.ReadLine() : Count = Count + 1 : Loop '    Align to next items item base field - skips 48 fields and checkflag
        Loop

        CheckBases.Close()

        'Seperates All Item Base Entries to a list of single entries - Removes dupes by creating new list. Old list now used for ratio

        For Each item In ItemBaseList
            If ItemBaseGroups.Contains(item) = False Then ItemBaseGroups.Add(item)
        Next


        'Get Tallys for each item base listed.                                                      Ratio is displayed as a percentage of all items in the database

        For Each ItemGroup In ItemBaseGroups '                       Each Individual Group
            Count = 0
            For Each ItemBase In ItemBaseList '                      In all All Collated Groups
                If ItemGroup = ItemBase Then Count = Count + 1 '     Total Tally Counter
            Next
            ItemBaseValues.Add(Count)
        Next

        'FINISH - Update Item Base Data Grid Layout Page.                                                    Print compiled data to screen

        For Each item In ItemBaseGroups
            DatabaseInfoDATAGRIDVIEW.Rows.Add(item, ItemBaseValues(Count), ItemBaseValues(Count) / ItemBaseList.Count * 100)
            Count = Count + 1
        Next

    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------






















End Class