Module UserListFunctions
    Public Sub SelectedItemsToUserList()
        If AutoLoggerRunning = True Then Return

        Main.ImportTimer.Stop()

        If Main.AllItemsLISTBOX.SelectedIndices.Count > 0 Then
            Dim a As Integer = 0
            Dim count As Integer = 0
            For index = 0 To Main.AllItemsLISTBOX.SelectedIndices.Count - 1
                transferobjects(Main.AllItemsLISTBOX.SelectedIndices(index))
                UserListReferenceList.Add(Main.AllItemsLISTBOX.SelectedIndices(index))
            Next
            Main.AllItemsLISTBOX.SelectedIndex = -1
        End If


        Main.ListboxTABCONTROL.SelectTab(3)
        Main.UserRefControlTabBUTTON.BackgroundImage = My.Resources.ButtonBackground
        Main.SearchListControlTabBUTTON.BackgroundImage = Nothing
        Main.ListControlTabBUTTON.BackgroundImage = Nothing
        Main.TradesListControlTabBUTTON.BackgroundImage = Nothing
        Main.LadderTEXTBOX.Show()
        Main.Ladder.Show()
        Main.ImportTimer.Start()
    End Sub

    Public Sub SearchItemsToUserList()
        'CHECK FOR LOGGER ACTIVE
        If AutoLoggerRunning = True Then Return
        Main.ImportTimer.Stop()

        Dim a As Integer = 0

        If Main.SearchLISTBOX.Items.Count > 0 And Main.SearchLISTBOX.SelectedIndex <> -1 Then 'Check items exist in the search list - Only does the following if they do
            For Each ItemIndex In Main.SearchLISTBOX.SelectedIndices                    'Setup Loop to 
                UserListReferenceList.Add(SearchReferenceList(ItemIndex))           'Add Item index value(in object database) for stats reference
                transferobjects(SearchReferenceList(ItemIndex))                     'routine for exchange data in database management functions
            Next

            Main.UserLISTBOX.SelectedItem = Main.SearchLISTBOX.SelectedItem                   'Select first moved item placed in user list
            Main.ListboxTABCONTROL.SelectTab(3)                                          'Auto Select User List Tab

        End If
        Main.ImportTimer.Start()
    End Sub

    Public Sub transferobjects(ByVal a)

        Dim AddToUserList As New UserListDatabase
        AddToUserList.ItemRealm = ItemObjects(a).ItemRealm
        AddToUserList.ItemName = ItemObjects(a).ItemName
        AddToUserList.ItemBase = ItemObjects(a).ItemBase
        AddToUserList.ItemQuality = ItemObjects(a).ItemQuality
        AddToUserList.RequiredCharacter = ItemObjects(a).RequiredCharacter
        AddToUserList.EtherealItem = ItemObjects(a).EtherealItem
        AddToUserList.Sockets = ItemObjects(a).Sockets
        AddToUserList.RuneWord = ItemObjects(a).RuneWord
        AddToUserList.ThrowDamageMin = ItemObjects(a).ThrowDamageMin
        AddToUserList.ThrowDamageMax = ItemObjects(a).ThrowDamageMax
        AddToUserList.OneHandDamageMin = ItemObjects(a).OneHandDamageMin
        AddToUserList.OneHandDamageMax = ItemObjects(a).OneHandDamageMax
        AddToUserList.TwoHandDamageMin = ItemObjects(a).TwoHandDamageMin
        AddToUserList.TwoHandDamageMax = ItemObjects(a).TwoHandDamageMax
        AddToUserList.Defense = ItemObjects(a).Defense
        AddToUserList.ChanceToBlock = ItemObjects(a).ChanceToBlock
        AddToUserList.QuantityMin = ItemObjects(a).QuantityMin
        AddToUserList.QuantityMax = ItemObjects(a).QuantityMax
        AddToUserList.DurabilityMin = ItemObjects(a).DurabilityMin
        AddToUserList.DurabilityMax = ItemObjects(a).DurabilityMax
        AddToUserList.RequiredStrength = ItemObjects(a).RequiredStrength
        AddToUserList.RequiredDexterity = ItemObjects(a).RequiredDexterity
        AddToUserList.RequiredLevel = ItemObjects(a).RequiredLevel
        AddToUserList.AttackClass = ItemObjects(a).AttackClass
        AddToUserList.AttackSpeed = ItemObjects(a).AttackSpeed
        AddToUserList.Stat1 = ItemObjects(a).Stat1
        AddToUserList.Stat2 = ItemObjects(a).Stat2
        AddToUserList.Stat3 = ItemObjects(a).Stat3
        AddToUserList.Stat4 = ItemObjects(a).Stat4
        AddToUserList.Stat5 = ItemObjects(a).Stat5
        AddToUserList.Stat6 = ItemObjects(a).Stat6
        AddToUserList.Stat7 = ItemObjects(a).Stat7
        AddToUserList.Stat8 = ItemObjects(a).Stat8
        AddToUserList.Stat9 = ItemObjects(a).Stat9
        AddToUserList.Stat10 = ItemObjects(a).Stat10
        AddToUserList.Stat11 = ItemObjects(a).Stat11
        AddToUserList.Stat12 = ItemObjects(a).Stat12
        AddToUserList.Stat13 = ItemObjects(a).Stat13
        AddToUserList.Stat14 = ItemObjects(a).Stat14
        AddToUserList.Stat15 = ItemObjects(a).Stat15
        AddToUserList.MuleName = ItemObjects(a).MuleName
        AddToUserList.MuleAccount = ItemObjects(a).MuleAccount
        AddToUserList.MulePass = ItemObjects(a).MulePass
        AddToUserList.HardCore = ItemObjects(a).HardCore
        AddToUserList.Ladder = ItemObjects(a).Ladder
        AddToUserList.Expansion = ItemObjects(a).Expansion
        AddToUserList.PickitAccount = ItemObjects(a).PickitAccount
        AddToUserList.UserField = ItemObjects(a).UserField
        AddToUserList.ItemImage = ItemObjects(a).ItemImage
        AddToUserList.ImportDate = ItemObjects(a).ImportDate
        AddToUserList.ImportTime = ItemObjects(a).ImportTime

        Dim temp = AppSettings.CurrentDatabase.Split("\")

        AddToUserList.DatabaseFilename = temp(temp.Length - 1)
        UserObjects.Add(AddToUserList)
        Main.UserLISTBOX.Items.Add(AddToUserList.ItemName)
    End Sub

    Public Sub DisplaySelectedUserListItem()
        If Main.UserLISTBOX.SelectedIndex = -1 Then Return

        Main.AllItemsLISTBOX.SelectedItem = -1

        Dim ItemIndex = Main.UserLISTBOX.SelectedIndex

        If ItemIndex < 0 Or ItemIndex >= UserObjects.Count Then
            Main.MuleRealmTEXTBOX.Clear()
            Main.MuleAccountTEXTBOX.Clear()
            Main.MuleNameTEXTBOX.Clear()
            Main.MulePasswordTEXTBOX.Clear()
            Main.CoreTypeTEXTBOX.Clear()
            Main.ItemStatsRICHTEXTBOX.Clear()
            Return
        End If

        Main.DatabaseFileLABEL.Show()
        Main.ItemStatsRICHTEXTBOX.Clear() 'moved this here as occassionally getting double display nfi why

        'Display mule details
        Main.MuleRealmTEXTBOX.Text = UserObjects(ItemIndex).ItemRealm
        Main.MuleAccountTEXTBOX.Text = UserObjects(ItemIndex).MuleAccount
        Main.MuleNameTEXTBOX.Text = UserObjects(ItemIndex).MuleName
        If AppSettings.HideMulePass = False Then Main.MulePasswordTEXTBOX.UseSystemPasswordChar = False
        If AppSettings.HideMulePass = True Then Main.MulePasswordTEXTBOX.UseSystemPasswordChar = True
        Main.MulePasswordTEXTBOX.Text = UserObjects(ItemIndex).MulePass
        If UserObjects(ItemIndex).HardCore = True Then Main.CoreTypeTEXTBOX.Text = "HardCore"
        If UserObjects(ItemIndex).HardCore = False Then Main.CoreTypeTEXTBOX.Text = "SoftCore"
        If ItemObjects(ItemIndex).Ladder = False Then Main.LadderTEXTBOX.Text = "Non Ladder" '  <-----------ExportErrorHereOccasionally
        If ItemObjects(ItemIndex).Ladder = True Then Main.LadderTEXTBOX.Text = "Ladder" '       <-----------ExportErrorHereOccasionally

        Main.DatabaseFileNameTEXTBOX.Text = UserObjects(ItemIndex).DatabaseFilename

        Dim DisplayColour As String = UserObjects(ItemIndex).ItemQuality
        Dim ColourCount1 As Integer = UserObjects(ItemIndex).ItemQuality.Length


        'Normal And SAuperior - White
        If (DisplayColour = "Normal" Or DisplayColour = "Superior") And UserObjects(ItemIndex).RuneWord = False Then
            If UserObjects(ItemIndex).ItemBase = "Rune" Then
                Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
                Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Quest and Special Items - Orange
            If UserObjects(ItemIndex).ItemBase = "Quest" Then
                Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
                Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
            End If

            'Runeword Items - White
            If UserObjects(ItemIndex).ItemName.IndexOf("Rune") = -1 And UserObjects(ItemIndex).ItemBase <> "Quest" Then
                Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.White
                Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
            End If
        End If

        'Magic items - Blue
        If DisplayColour = "Magic" Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.DodgerBlue
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Rares - Yellow
        If DisplayColour = "Rare" Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.Yellow
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Crafted - Gold
        If DisplayColour = "Crafted" Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.DarkGoldenrod
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Set Items - Green
        If DisplayColour = "Set" Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.Green
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
        End If

        'Uniques - Light Gold
        If DisplayColour = "Unique" Or UserObjects(ItemIndex).RuneWord = True Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.BurlyWood
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).ItemName & vbCrLf
        End If

        'RuneWord String - Still Not Reporting Right
        If UserObjects(ItemIndex).RuneWord = True Then
            Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.Orange
            Main.ItemStatsRICHTEXTBOX.SelectedText = UserObjects(ItemIndex).Stat1 & vbCrLf
        End If

        Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf) '          Spacer line after Item Name and class Always 
        ColourCount1 = Main.ItemStatsRICHTEXTBOX.TextLength '   Used to Count number of lines to calculate selection to colour text selection for the Basic Info Block - this var represents the starting point to colour

        'White text for basic info Block - Line spaceing added between each section (if needed)
        If UserObjects(ItemIndex).OneHandDamageMax > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("One Hand Damage: " & UserObjects(ItemIndex).OneHandDamageMin & " to " & UserObjects(ItemIndex).OneHandDamageMax & vbCrLf)
        If UserObjects(ItemIndex).TwoHandDamageMax > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Two Hand Damage: " & UserObjects(ItemIndex).TwoHandDamageMin & " to " & UserObjects(ItemIndex).TwoHandDamageMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If Main.DisplayLineBreaksMENUITEM.Checked = True Then
            If UserObjects(ItemIndex).OneHandDamageMax > 0 Or UserObjects(ItemIndex).TwoHandDamageMax > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Item Defensive Values
        If UserObjects(ItemIndex).Defense > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Defense: " & UserObjects(ItemIndex).Defense & vbCrLf)
        If UserObjects(ItemIndex).ChanceToBlock > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Chance To Block: " & UserObjects(ItemIndex).ChanceToBlock & "%" & vbCrLf)
        If UserObjects(ItemIndex).DurabilityMin > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Durability: " & UserObjects(ItemIndex).DurabilityMin & " of " & UserObjects(ItemIndex).DurabilityMax & vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If Main.DisplayLineBreaksMENUITEM.Checked = True Then
            If UserObjects(ItemIndex).Defense > 0 Or UserObjects(ItemIndex).ChanceToBlock > 0 Or UserObjects(ItemIndex).DurabilityMin > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf)
        End If

        'Item Requirement Values
        If UserObjects(ItemIndex).RequiredStrength > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Required Strength: " & UserObjects(ItemIndex).RequiredStrength & vbCrLf)
        If UserObjects(ItemIndex).RequiredDexterity > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Required Dexterity: " & UserObjects(ItemIndex).RequiredDexterity & vbCrLf)
        If UserObjects(ItemIndex).RequiredLevel > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText("Required Level: " & UserObjects(ItemIndex).RequiredLevel & vbCrLf)

        'Required Character Displayed Red and in Brackets to copy the ingame D2 themed layout
        If UserObjects(ItemIndex).RequiredCharacter <> Nothing Then
            Main.ItemStatsRICHTEXTBOX.SelectedText = "[" & UserObjects(ItemIndex).RequiredCharacter & " Only]" & vbCrLf
        End If

        'ADD lINE SPACING BASED ON OPTION SETTING
        If Main.DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).RequiredStrength > 0 Or ItemObjects(ItemIndex).RequiredDexterity > 0 Or ItemObjects(ItemIndex).RequiredCharacter <> Nothing And ItemObjects(ItemIndex).RequiredLevel > 0 Then Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf) '                                                                                                   <-----------ExportErrorHereOccasionally
        End If

        'Item Attack Class And Speed
        If UserObjects(ItemIndex).AttackClass <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).AttackClass & " Class") : If UserObjects(ItemIndex).AttackSpeed <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(" - " & UserObjects(ItemIndex).AttackSpeed & vbCrLf) Else Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf)

        'ADD lINE SPACING BASED ON OPTION SETTING
        If Main.DisplayLineBreaksMENUITEM.Checked = True Then
            If ItemObjects(ItemIndex).AttackClass <> Nothing Or ItemObjects(ItemIndex).AttackSpeed <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf) '                                                                                                                                                                                                    <-----------ExportErrorHereOccasionally
            If ItemObjects(ItemIndex).RequiredLevel > 0 And ItemObjects(ItemIndex).AttackClass = Nothing And ItemObjects(ItemIndex).AttackSpeed = Nothing And ItemObjects(ItemIndex).RequiredCharacter = Nothing And ItemObjects(ItemIndex).RequiredDexterity = 0 And ItemObjects(ItemIndex).RequiredStrength = 0 Then Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf) '<-----------ExportErrorHereOccasionally
        End If

        'Colour Above Displayed Basic Info Text Block White
        Dim ColourCount2 As Integer = Main.ItemStatsRICHTEXTBOX.TextLength - ColourCount1 'Calculate difference between Basic Info Block and Other Stats to Color Basic Info - this var represents the finishing point to colour
        Main.ItemStatsRICHTEXTBOX.Select(ColourCount1, ColourCount2)
        Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.White

        'Unique Attributes Block as Default Blue (No need to colour these as they are blue by default)
        If UserObjects(ItemIndex).Stat1 <> Nothing And UserObjects(ItemIndex).RuneWord = False Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat1 & vbCrLf)
        If UserObjects(ItemIndex).Stat2 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat2 & vbCrLf)
        If UserObjects(ItemIndex).Stat3 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat3 & vbCrLf)
        If UserObjects(ItemIndex).Stat4 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat4 & vbCrLf)
        If UserObjects(ItemIndex).Stat5 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat5 & vbCrLf)
        If UserObjects(ItemIndex).Stat6 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat6 & vbCrLf)
        If UserObjects(ItemIndex).Stat7 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat7 & vbCrLf)
        If UserObjects(ItemIndex).Stat8 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat8 & vbCrLf)
        If UserObjects(ItemIndex).Stat9 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat9 & vbCrLf)
        If UserObjects(ItemIndex).Stat10 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat10 & vbCrLf)
        If UserObjects(ItemIndex).Stat11 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat11 & vbCrLf)
        If UserObjects(ItemIndex).Stat12 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat12 & vbCrLf)
        If UserObjects(ItemIndex).Stat13 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat13 & vbCrLf)
        If UserObjects(ItemIndex).Stat14 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat14 & vbCrLf)
        If UserObjects(ItemIndex).Stat15 <> Nothing Then Main.ItemStatsRICHTEXTBOX.AppendText(UserObjects(ItemIndex).Stat15 & vbCrLf)


        Main.ItemStatsRICHTEXTBOX.AppendText(vbCrLf & "Item Level: " & UserObjects(ItemIndex).Itemlevel & vbCrLf)

        'Select All and Center Justify Text Allignment in the ItemStatsRICHTEXTBOX - Display Items Routine is DONE :)
        Main.ItemStatsRICHTEXTBOX.SelectAll()
        Main.ItemStatsRICHTEXTBOX.SelectionAlignment = HorizontalAlignment.Center

        Main.ItemSkinPICTUREBOX.Load("Skins\" + ImageArray(UserObjects(ItemIndex).ItemImage) + ".jpg")

        'THIS DITTY CHANGES THE "Item Level: 00" LINE FROM BLUE TO WHITE (looks nicer and seperates it from the unique attribs block)
        'NOTE TO MYSELF: is something not right with the runes display, seems to add extra spaces, only does it to runes, atm nfi wtf, perhaps logging differs to other unique attribs blocks is my guess REMEMBER TO FIX!
        Dim linecount As Integer = 0
        For Each Line In Main.ItemStatsRICHTEXTBOX.Lines
            If Line.IndexOf("Item Level:") > -1 Then Main.ItemStatsRICHTEXTBOX.Select(Main.ItemStatsRICHTEXTBOX.Text.Length - Len(Line), Len(Line))
            linecount = linecount + 1
        Next
        Main.ItemStatsRICHTEXTBOX.SelectionColor = Color.White
    End Sub
End Module
