Public Class DatabaseInfo
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'SET FORMS X,Y Co-ordinates and  Height, witdth from AppSettings value
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DatabaseInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Height = AppSettings.XSizeInfo
        Me.Width = AppSettings.YSizeInfo
        Me.Location = New Point(AppSettings.XPosInfo, AppSettings.YPosInfo)

        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)


    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Closes Database Info Form
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DatabaseInfoCloseBUTTON_Click(sender As Object, e As EventArgs) Handles DatabaseInfoCloseBUTTON.Click
        Me.Close()
        If AppSettings.SoundMute = False Then My.Computer.Audio.Play(My.Resources.d2Dong, AudioPlayMode.Background)
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Refresh Database Info Form
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub DatabaseInfoRefreshBUTTON_Click(sender As Object, e As EventArgs) Handles DatabaseInfoRefreshBUTTON.Click

        'Branch to DatabaseInfo Form & routines if a database is selected if not do nothing
        If DatabaseManager.DatabaseManagerSavedDatabasesLISTBOX.SelectedIndex <> -1 Then
            Me.Show()
            Me.ClearOldData()
            Me.GetItemTotal() 'item total must come before header info for it to be displayed correctly
            Me.GetHeaderInfo()
            Me.GetItemBases()
            Me.DatabaseInfoTABCONTROL.SelectTab(0)
            Me.DatabaseInfoCloseBUTTON.Select()
        End If
    End Sub
    Sub ClearInfoButtons()
        InfoBaseBUTTON.BackgroundImage = Nothing
        InfoQualityBUTTON.BackgroundImage = Nothing
        IntoRuneBUTTON.BackgroundImage = Nothing
        InfoRealmBUTTON.BackgroundImage = Nothing
        InfoMuleBUTTON.BackgroundImage = Nothing
        InfoAccountBUTTON.BackgroundImage = Nothing
        InfoFileBUTTON.BackgroundImage = Nothing
    End Sub




    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Database Info Form Tab Control Buttons, D2 Style buttons. button press focuses on eachg tabpage
    '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'Item Base
    Private Sub InfoBaseBUTTON_Click(sender As Object, e As EventArgs) Handles InfoBaseBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(0)
        ClearInfoButtons()
        InfoBaseBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'Item Quality
    Private Sub InfoQualityBUTTON_Click(sender As Object, e As EventArgs) Handles InfoQualityBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(1)
        ClearInfoButtons()
        InfoQualityBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'Runes
    Private Sub IntoRuneBUTTON_Click(sender As Object, e As EventArgs) Handles IntoRuneBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(2)
        ClearInfoButtons()
        IntoRuneBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'Realms
    Private Sub InfoRealmBUTTON_Click(sender As Object, e As EventArgs) Handles InfoRealmBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(3)
        ClearInfoButtons()
        InfoRealmBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'Mules
    Private Sub InfoMuleBUTTON_Click(sender As Object, e As EventArgs) Handles InfoMuleBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(4)
        ClearInfoButtons()
        InfoMuleBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'Accounts
    Private Sub InfoAccountBUTTON_Click(sender As Object, e As EventArgs) Handles InfoAccountBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(5)
        ClearInfoButtons()
        InfoAccountBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    'File
    Private Sub InfoFileBUTTON_Click(sender As Object, e As EventArgs) Handles InfoFileBUTTON.Click
        DatabaseInfoTABCONTROL.SelectTab(6)
        ClearInfoButtons()
        InfoFileBUTTON.BackgroundImage = My.Resources.ButtonBackground
        DatabaseInfoCloseBUTTON.Select()
    End Sub

    '-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'UPDATES HEADER INFORMATION - Database Filename and Total Items
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub GetHeaderInfo()
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
        Dim CountItems = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Databases\" + Me.DatabaseInfoSelectedTEXTBOX.Text + ".TXT")
        TotalItemsInSelected = 0
        Do Until CountItems.EndOfStream
            If CountItems.ReadLine = "--------------------" Then TotalItemsInSelected = TotalItemsInSelected + 1
        Loop
        CountItems.Close()
        Me.DatabaseInfoTotalTEXTBOX.Text = TotalItemsInSelected
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'GETS PERCENTAGE RATIOS AND TALLYS OF ALL ITEM BASES MENTIONED IN SELECTED DATABASE
    '-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub GetItemBases()

        'START - Opens Database File. Uses Database Manager Listbox Selected item as filename selector.
        '----------------------------------------------------------------------------------------------
        Count = 0 : ItemBaseList.Clear() : ItemBaseGroups.Clear() : ItemBaseValues.Clear()
        Dim CheckBases = My.Computer.FileSystem.OpenTextFileReader(AppSettings.InstallPath + "\Databases\" + Me.DatabaseInfoSelectedTEXTBOX.Text + ".TXT")
        CheckBases.ReadLine() '                                                                     Skip Checkflag / Record Seperator Line

        'Read All Item Base Fields From The Selected Databases File
        '----------------------------------------------------------
        Do Until CheckBases.EndOfStream
            CheckBases.ReadLine() : CheckBases.ReadLine() : CheckBases.ReadLine() '                 Align to Item Base Field as next field to be read - skip first three fields
            Temp = CheckBases.ReadLine() : ItemBaseList.Add(Temp) '                                 Read file and assign to array
            Count = 0 : Do Until Count = 49 : CheckBases.ReadLine() : Count = Count + 1 : Loop '    Align to next items item base field - skips 48 fields and checkflag
        Loop
        CheckBases.Close()

        'Seperates All Item Base Entries to a list of single entries - Removes dupes by creating new list. Old list now used for ratio
        '-----------------------------------------------------------------------------------------------------------------------------
        For Each item In ItemBaseList
            If ItemBaseGroups.Contains(item) = False Then ItemBaseGroups.Add(item)
        Next

        'Get Tallys for each item base listed.                      Ratio is displayed as a percentage of all items in the database
        '--------------------------------------------------------------------------------------------------------------------------
        For Each ItemGroup In ItemBaseGroups '                      Each Individual Group
            Count = 0
            For Each ItemBase In ItemBaseList '                     In all All Collated Groups
                If ItemGroup = ItemBase Then Count = Count + 1 '    Total Tally Counter
            Next
            ItemBaseValues.Add(Count)
        Next

        'FINISH - Update Item Base Data Grid Layout Page.           Print compiled data to screen
        '----------------------------------------------------------------------------------------
        Count = 0
        For Each item In ItemBaseGroups
            DatabaseInfoDATAGRIDVIEW.Rows.Add(item, ItemBaseValues(Count), ItemBaseValues(Count) / ItemBaseList.Count * 100)
            Count = Count + 1
        Next
    End Sub

    '----------------------------------------------------
    'Close Routine to Get Form Size and Position Values
    '----------------------------------------------------
    Private Sub DatabaseInfo_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        AppSettings.XSizeInfo = Me.Height : AppSettings.YSizeInfo = Me.Width
        AppSettings.XPosInfo = Me.Location.X : AppSettings.YPosInfo = Me.Location.Y

    End Sub


    'set d2 fonts
    Private Sub DatabaseInfo_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        '    ListControlTabBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)


        'tab buttons
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "/Extras/DiabloFont1.TTF") = True Then
            InfoBaseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            InfoQualityBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            IntoRuneBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            InfoRealmBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            InfoMuleBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            InfoAccountBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            InfoFileBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)

            'function buttons
            DatabaseInfoCloseBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            DatabaseInfoRefreshBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)


            'labels
            DatabaseInfoSelectedLABEL.Font = New Font(pfc.Families(0), 11, FontStyle.Regular)
            DatabaseInfoTotalLABEL.Font = New Font(pfc.Families(0), 11, FontStyle.Regular)

        End If
    End Sub
End Class