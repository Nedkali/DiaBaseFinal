Public Class ItemAdd


    Private Sub AllTextBoxes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AddItemNameTEXTBOX.KeyDown, AddItemBaseCOMBOBOX.KeyDown, AddItemQualityCOMBOBOX.KeyDown,
            AddItemLevelTEXTBOX.KeyDown, AddItemLadderCheckBox.KeyDown, AddItemRunewordCheckBox.KeyDown, AddItemEtherealItemCHECKBOX.KeyDown, AddItemSocketsNumericUpDown.KeyDown, AddItemThrowDamageMinTEXTBOX.KeyDown,
            AddItemThrowDamageMaxTEXTBOX.KeyDown, AddItemOneHandDamageMinTEXTBOX.KeyDown, AddItemOneHandDamageMaxTEXTBOX.KeyDown, AddItemTwoHandDamageMinTEXTBOX.KeyDown, AddItemTwoHandDamageMaxTEXTBOX.KeyDown,
            AddItemQuantityMinTEXTBOX.KeyDown, AddItemQuantityMaxTEXTBOX.KeyDown, AddItemDurabilityMinTEXTBOX.KeyDown, AddItemDurabilityMaxTEXTBOX.KeyDown, AddItemDefenseTEXTBOX.KeyDown, AddItemChanceToBlockTEXTBOX.KeyDown,
            AddItemRequiredStrengthTEXTBOX.KeyDown, AddItemRequiredDexterityTEXTBOX.KeyDown, AddItemRequiredLevelTEXTBOX.KeyDown, AddItemReqCharCOMBOBOX.KeyDown, AddItemAttackClassCOMBOBOX.KeyDown,
            AddItemAttackSpeedCOMBOBOX.KeyDown, AddItemStat1TEXTBOX.KeyDown, AddItemStat2TEXTBOX.KeyDown, AddItemStat3TEXTBOX.KeyDown, AddItemStat4TEXTBOX.KeyDown, AddItemStat5TEXTBOX.KeyDown,
            AddItemStat6TEXTBOX.KeyDown, AddItemStat7TEXTBOX.KeyDown, AddItemStat8TEXTBOX.KeyDown, AddItemStat9TEXTBOX.KeyDown, AddItemStat10TEXTBOX.KeyDown, AddItemStat11TEXTBOX.KeyDown,
            AddItemStat12TEXTBOX.KeyDown, AddItemStat13TEXTBOX.KeyDown, AddItemStat14TEXTBOX.KeyDown, AddItemStat15TEXTBOX.KeyDown, AddItemImageTEXTBOX.KeyDown, AddItemRealmCOMBOBOX.KeyDown,
            AddItemMuleAccountCOMBOBOX.KeyDown, AddItemMulePassCOMBOBOX.KeyDown, AddItemMuleNameCOMBOBOX.KeyDown, AddItemCoreTypeCOMBOBOX.KeyDown, AddItemPickitBotCOMBOBOX.KeyDown, AddItemImportDateTEXTBOX.KeyDown,
            AddItemImportTimeTEXTBOX.KeyDown, AddItemUserReferenceTEXTBOX.KeyDown, AddItemSaveChangesBUTTON.KeyDown

        'Checks For ENTER Keypress and Switches Focus using tabbing order (carrage returns to next field after each entry)
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(True)
        End If
    End Sub
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'LOAD HANDLER - POPULATES THE QUICK EDIT COMBOBOX LISTS WITH ENTRIES THAT ALREADY EXIST IN THE DATABASE
    '-------------------------------------------------------------------------------------------------------------------------------------------

    Private Sub ItemAdd_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AddItemMuleAccountCOMBOBOX.Items.Clear()
        AddItemMuleNameCOMBOBOX.Items.Clear()
        AddItemMulePassCOMBOBOX.Items.Clear()
        AddItemBaseCOMBOBOX.Items.Clear()
        AddItemQualityCOMBOBOX.Items.Clear()
        AddItemAttackClassCOMBOBOX.Items.Clear()
        AddItemPickitBotCOMBOBOX.Items.Clear()


        For Each ItemObjectItem As ItemDatabase In ItemObjects
            If ItemObjectItem.MuleName <> "" Then
                If AddItemMuleNameCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then AddItemMuleNameCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
            End If
            If ItemObjectItem.MuleAccount <> "" Then
                If AddItemMuleAccountCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then AddItemMuleAccountCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
            End If
            If ItemObjectItem.MulePass <> "" Then
                If AddItemMulePassCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then AddItemMulePassCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
            End If
            If ItemObjectItem.AttackClass <> "" Then
                If AddItemAttackClassCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then AddItemAttackClassCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
            End If
            If ItemObjectItem.ItemBase <> "" Then
                If AddItemBaseCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then AddItemBaseCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
            End If
            If ItemObjectItem.ItemQuality <> "" Then
                If AddItemQualityCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then AddItemQualityCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
            End If
            If ItemObjectItem.PickitAccount <> "" Then
                If AddItemPickitBotCOMBOBOX.Items.Contains(ItemObjectItem.PickitAccount) = False Then AddItemPickitBotCOMBOBOX.Items.Add(ItemObjectItem.PickitAccount)
            End If
        Next
    End Sub

    '-------------------------------------------------------------------------------------------------------------------------------------------
    'ADD ITEM SHOWN HANDLER - SETS UP AND CLEARS FIELD CONTROLS
    '-------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub ItemAdd_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        InitializeItemFields()
    End Sub
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'CLEARS OUT OLD ADD ITEM FIELDS
    '-------------------------------------------------------------------------------------------------------------------------------------------
    Sub InitializeItemFields()
        AddItemNameTEXTBOX.Text = Nothing
        AddItemBaseCOMBOBOX.Text = Nothing
        AddItemQualityCOMBOBOX.SelectedIndex = 0
        AddItemRunewordCheckBox.Checked = False
        AddItemSocketsNumericUpDown.Value = 0
        AddItemEtherealItemCHECKBOX.Checked = False
        AddItemThrowDamageMinTEXTBOX.Text = "0"
        AddItemThrowDamageMaxTEXTBOX.Text = "0"
        AddItemOneHandDamageMinTEXTBOX.Text = "0"
        AddItemOneHandDamageMaxTEXTBOX.Text = "0"
        AddItemTwoHandDamageMinTEXTBOX.Text = "0"
        AddItemTwoHandDamageMaxTEXTBOX.Text = "0"
        AddItemQuantityMinTEXTBOX.Text = "0"
        AddItemQuantityMaxTEXTBOX.Text = "0"
        AddItemDurabilityMinTEXTBOX.Text = "0"
        AddItemDurabilityMaxTEXTBOX.Text = "0"
        AddItemDefenseTEXTBOX.Text = "0"
        AddItemChanceToBlockTEXTBOX.Text = "0"
        AddItemRequiredStrengthTEXTBOX.Text = "0"
        AddItemRequiredDexterityTEXTBOX.Text = "0"
        AddItemRequiredLevelTEXTBOX.Text = "0"
        AddItemAttackClassCOMBOBOX.Text = Nothing
        AddItemAttackSpeedCOMBOBOX.Text = Nothing
        AddItemReqCharCOMBOBOX.Text = Nothing
        AddItemImageTEXTBOX.Text = "0"
        AddItemUserReferenceTEXTBOX.Text = ""
        AddItemStat1TEXTBOX.Text = Nothing
        AddItemStat2TEXTBOX.Text = Nothing
        AddItemStat3TEXTBOX.Text = Nothing
        AddItemStat4TEXTBOX.Text = Nothing
        AddItemStat5TEXTBOX.Text = Nothing
        AddItemStat6TEXTBOX.Text = Nothing
        AddItemStat7TEXTBOX.Text = Nothing
        AddItemStat8TEXTBOX.Text = Nothing
        AddItemStat9TEXTBOX.Text = Nothing
        AddItemStat10TEXTBOX.Text = Nothing
        AddItemStat11TEXTBOX.Text = Nothing
        AddItemStat12TEXTBOX.Text = Nothing
        AddItemStat13TEXTBOX.Text = Nothing
        AddItemStat14TEXTBOX.Text = Nothing
        AddItemStat15TEXTBOX.Text = Nothing
        AddItemLevelTEXTBOX.Text = "1"
        AddItemMuleNameCOMBOBOX.Text = Nothing
        AddItemMuleAccountCOMBOBOX.Text = Nothing
        AddItemPickitBotCOMBOBOX.Text = Nothing
        AddItemMulePassCOMBOBOX.Text = Nothing
        AddItemCoreTypeCOMBOBOX.Text = "SoftCore"
        AddItemRealmCOMBOBOX.Text = Nothing
        AddItemLadderCheckBox.Checked = True
        AddItemImportTimeTEXTBOX.Text = Nothing
        AddItemImportDateTEXTBOX.Text = Nothing

        'Wheres the hide pass routine lol?
        'If AppSettings.HideMulePass = True Then AddItemMulePassCOMBOBOX.Text = HidePassRoutine(ItemObjects(ItemIndexNumber).MulePass) Else AddItemMulePassCOMBOBOX.Text =  ""(ItemIndexNumber).MulePass







    End Sub
    '-------------------------------------------------------------------------------------------------------------------------------------------
    'SAVES THE NEWLY CRETED ITEM TO DATABASE AND ADDS IT TO THE ALL ITEMS LIST
    '-------------------------------------------------------------------------------------------------------------------------------------------
    Private Sub AddItemSaveChangesBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemSaveChangesBUTTON.Click

        If AddItemNameTEXTBOX.Text <> "" Then

            Dim NewItem As New ItemDatabase '                           Defines NewItem As A New Object for ItemDatabase Class
            NewItem.ItemName = AddItemNameTEXTBOX.Text
            NewItem.ItemBase = AddItemBaseCOMBOBOX.Text
            NewItem.ItemQuality = AddItemQualityCOMBOBOX.Text
            NewItem.RuneWord = AddItemRunewordCheckBox.Checked
            NewItem.Sockets = AddItemSocketsNumericUpDown.Value
            NewItem.EtherealItem = AddItemEtherealItemCHECKBOX.Checked
            NewItem.ThrowDamageMin = Convert.ToInt32(AddItemThrowDamageMinTEXTBOX.Text)
            NewItem.ThrowDamageMax = AddItemThrowDamageMaxTEXTBOX.Text
            NewItem.OneHandDamageMin = AddItemOneHandDamageMinTEXTBOX.Text
            NewItem.OneHandDamageMax = AddItemOneHandDamageMaxTEXTBOX.Text
            NewItem.TwoHandDamageMin = AddItemTwoHandDamageMinTEXTBOX.Text
            NewItem.TwoHandDamageMax = AddItemTwoHandDamageMaxTEXTBOX.Text
            NewItem.QuantityMin = AddItemQuantityMinTEXTBOX.Text
            NewItem.QuantityMax = AddItemQuantityMaxTEXTBOX.Text
            NewItem.DurabilityMin = AddItemDurabilityMinTEXTBOX.Text
            NewItem.DurabilityMax = AddItemDurabilityMaxTEXTBOX.Text
            NewItem.Defense = AddItemDefenseTEXTBOX.Text
            NewItem.ChanceToBlock = AddItemChanceToBlockTEXTBOX.Text
            NewItem.RequiredStrength = AddItemRequiredStrengthTEXTBOX.Text
            NewItem.RequiredDexterity = AddItemRequiredDexterityTEXTBOX.Text
            NewItem.RequiredLevel = AddItemRequiredLevelTEXTBOX.Text
            NewItem.AttackClass = AddItemAttackClassCOMBOBOX.Text
            NewItem.AttackSpeed = AddItemAttackSpeedCOMBOBOX.Text
            NewItem.RequiredCharacter = AddItemReqCharCOMBOBOX.Text
            NewItem.ItemImage = AddItemImageTEXTBOX.Text
            NewItem.UserField = AddItemUserReferenceTEXTBOX.Text
            NewItem.Stat1 = AddItemStat1TEXTBOX.Text
            NewItem.Stat2 = AddItemStat2TEXTBOX.Text
            NewItem.Stat3 = AddItemStat3TEXTBOX.Text
            NewItem.Stat4 = AddItemStat4TEXTBOX.Text
            NewItem.Stat5 = AddItemStat5TEXTBOX.Text
            NewItem.Stat6 = AddItemStat6TEXTBOX.Text
            NewItem.Stat7 = AddItemStat7TEXTBOX.Text
            NewItem.Stat8 = AddItemStat8TEXTBOX.Text
            NewItem.Stat9 = AddItemStat9TEXTBOX.Text
            NewItem.Stat10 = AddItemStat10TEXTBOX.Text
            NewItem.Stat11 = AddItemStat11TEXTBOX.Text
            NewItem.Stat12 = AddItemStat12TEXTBOX.Text
            NewItem.Stat13 = AddItemStat13TEXTBOX.Text
            NewItem.Stat14 = AddItemStat14TEXTBOX.Text
            NewItem.Stat15 = AddItemStat15TEXTBOX.Text
            NewItem.Itemlevel = AddItemLevelTEXTBOX.Text
            NewItem.MuleName = AddItemMuleNameCOMBOBOX.Text
            NewItem.MuleAccount = AddItemMuleAccountCOMBOBOX.Text
            NewItem.PickitAccount = AddItemPickitBotCOMBOBOX.Text
            NewItem.MulePass = AddItemMulePassCOMBOBOX.Text
            NewItem.ImportDate = AddItemImportDateTEXTBOX.Text
            If AddItemCoreTypeCOMBOBOX.Text = "HardCore" Then NewItem.HardCore = True
            If AddItemCoreTypeCOMBOBOX.Text = "SoftCore" Then NewItem.HardCore = False
            NewItem.ItemRealm = AddItemRealmCOMBOBOX.Text
            NewItem.Ladder = AddItemLadderCheckBox.Checked
            NewItem.ImportTime = AddItemImportTimeTEXTBOX.Text
            ItemObjects.Add(NewItem)
            Main.AllItemsLISTBOX.Items.Add(NewItem.ItemName)
            Me.Close()

            'Selects the newly added item to make the added item clearly visible and so easier to user verify
            Main.AllItemsLISTBOX.SelectedIndex = -1 : Main.AllItemsLISTBOX.SelectedIndex = Main.AllItemsLISTBOX.Items.Count - 1
            Main.DisplayItemStats(Main.AllItemsLISTBOX.Items.Count - 1)

        Else

            'If No Item Name Entry Then Pass to Error Handler - ROBS NOTE TO MYSLEF: Also check for any other required fields ned has added
            Main.ErrorHandler(1101, 0, 0, 0)

            AddItemNameTEXTBOX.Select()

        End If

    End Sub

    'CLEAR ALL FIELDS BUTTON _ RESETS ALL ENTRSY IN ALL FIELDS
    Private Sub InitalizeFieldsBUTTON_Click(sender As Object, e As EventArgs) Handles InitalizeFieldsBUTTON.Click
        InitializeItemFields()
    End Sub



End Class