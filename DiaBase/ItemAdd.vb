Public Class ItemAdd

    Private Sub AddEdit_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        InitializeItemFields()



    End Sub


    Sub InitializeItemFields()

        AddItemNameTEXTBOX.Text = ""
        AddItemBaseCOMBOBOX.Text = ""
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
        AddItemAttackClassCOMBOBOX.Text = ""
        AddItemAttackSpeedCOMBOBOX.Text = ""
        AddItemReqCharCOMBOBOX.Text = ""
        AddItemImageTEXTBOX.Text = "0"
        AddItemUserReferenceTEXTBOX.Text = ""
        AddItemStat1TEXTBOX.Text = ""
        AddItemStat2TEXTBOX.Text = ""
        AddItemStat3TEXTBOX.Text = ""
        AddItemStat4TEXTBOX.Text = ""
        AddItemStat5TEXTBOX.Text = ""
        AddItemStat6TEXTBOX.Text = ""
        AddItemStat7TEXTBOX.Text = ""
        AddItemStat8TEXTBOX.Text = ""
        AddItemStat9TEXTBOX.Text = ""
        AddItemStat10TEXTBOX.Text = ""
        AddItemStat11TEXTBOX.Text = ""
        AddItemStat12TEXTBOX.Text = ""
        AddItemStat13TEXTBOX.Text = ""
        AddItemStat14TEXTBOX.Text = ""
        AddItemStat15TEXTBOX.Text = ""
        AddItemLevelTEXTBOX.Text = "1"
        AddItemMuleNameCOMBOBOX.Text = ""
        AddItemMuleAccountCOMBOBOX.Text = ""
        AddItemPickitBotCOMBOBOX.Text = ""
        AddItemMulePassCOMBOBOX.Text = ""
        AddItemCoreTypeCOMBOBOX.Text = "SoftCore"
        AddItemRealmCOMBOBOX.Text = ""
        AddItemLadderCheckBox.Checked = True
        AddItemImportTimeTEXTBOX.Text = ""
        AddItemImportDateTEXTBOX.Text = ""

        'Wheres the hide pass routine lol?
        'If AppSettings.HideMulePass = True Then AddItemMulePassCOMBOBOX.Text = HidePassRoutine(ItemObjects(ItemIndexNumber).MulePass) Else AddItemMulePassCOMBOBOX.Text =  ""(ItemIndexNumber).MulePass


    End Sub
    Private Sub AddItemSaveChangesBUTTON_Click(sender As Object, e As EventArgs) Handles AddItemSaveChangesBUTTON.Click



        Dim NewItem As New ItemDatabase '                           Define NewItem As A New Object for ItemDatabase Class
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

    End Sub
End Class