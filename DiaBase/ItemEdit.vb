'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'EDIT ITEM FORM 
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Class ItemEdit

    ' Private Sub AllTextBoxes_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles AddItemNameTEXTBOX.KeyDown, AddItemBaseCOMBOBOX.KeyDown, AddItemQualityCOMBOBOX.KeyDown,
    ' AddItemLevelTEXTBOX.KeyDown, AddItemLadderCheckBox.KeyDown, AddItemRunewordCheckBox.KeyDown, AddItemEtherealItemCHECKBOX.KeyDown, AddItemSocketsNumericUpDown.KeyDown, AddItemThrowDamageMinTEXTBOX.KeyDown,
    ' AddItemThrowDamageMaxTEXTBOX.KeyDown, AddItemOneHandDamageMinTEXTBOX.KeyDown, AddItemOneHandDamageMaxTEXTBOX.KeyDown, AddItemTwoHandDamageMinTEXTBOX.KeyDown, AddItemTwoHandDamageMaxTEXTBOX.KeyDown,
    ' AddItemQuantityMinTEXTBOX.KeyDown, AddItemQuantityMaxTEXTBOX.KeyDown, AddItemDurabilityMinTEXTBOX.KeyDown, AddItemDurabilityMaxTEXTBOX.KeyDown, AddItemDefenseTEXTBOX.KeyDown, AddItemChanceToBlockTEXTBOX.KeyDown,
    ' AddItemRequiredStrengthTEXTBOX.KeyDown, AddItemRequiredDexterityTEXTBOX.KeyDown, AddItemRequiredLevelTEXTBOX.KeyDown, AddItemReqCharCOMBOBOX.KeyDown, AddItemAttackClassCOMBOBOX.KeyDown,
    ' AddItemAttackSpeedCOMBOBOX.KeyDown, AddItemStat1TEXTBOX.KeyDown, AddItemStat2TEXTBOX.KeyDown, AddItemStat3TEXTBOX.KeyDown, AddItemStat4TEXTBOX.KeyDown, AddItemStat5TEXTBOX.KeyDown,
    ' AddItemStat6TEXTBOX.KeyDown, AddItemStat7TEXTBOX.KeyDown, AddItemStat8TEXTBOX.KeyDown, AddItemStat9TEXTBOX.KeyDown, AddItemStat10TEXTBOX.KeyDown, AddItemStat11TEXTBOX.KeyDown,
    ' AddItemStat12TEXTBOX.KeyDown, AddItemStat13TEXTBOX.KeyDown, AddItemStat14TEXTBOX.KeyDown, AddItemStat15TEXTBOX.KeyDown, AddItemImageTEXTBOX.KeyDown, AddItemRealmCOMBOBOX.KeyDown,
    ' AddItemMuleAccountCOMBOBOX.KeyDown, AddItemMulePassCOMBOBOX.KeyDown, AddItemMuleNameCOMBOBOX.KeyDown, AddItemCoreTypeCOMBOBOX.KeyDown, AddItemPickitBotCOMBOBOX.KeyDown, AddItemImportDateTEXTBOX.KeyDown,
    ' AddItemImportTimeTEXTBOX.KeyDown, AddItemUserReferenceTEXTBOX.KeyDown, AddItemSaveChangesBUTTON.KeyDown

    'Checks For ENTER Keypress and Switches Focus using tabbing order (carrage returns to next field after each entry)
    'If e.KeyCode = Keys.Enter Then
    'e.SuppressKeyPress = True
    'Me.ProcessTabKey(True)
    'End If
    'End Sub





    'EDIT ITEM SHOW EVENT - First Update the forms text / combo / numeric boxes with the selected items stats 
    Private Sub ItemEdit_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

        'must do's yet to do...

        'INCLUDE COMBOBOX POPULATE ROUTINES
        'ADD ENTER PRESS /PASS TO NEXT FIELD FUNXCTION
        'FIND THE HIDE PASSWORD ROUITINE AND ECORPERATE IT INTO THE FIELD UPATE ROUTINE
        'ADD HANDLER FOR PASSING ONLY EDITED FIELDS TO ALL OTHER SELECTED ITEMS IN MORE THAN ONE ITEM IS SELECTED FOR EDIT
        'ADD THE ADJUST FIELDS HANDLER FOR UPDATING THE DATABASE 


        'Was thinking bout adding a foward and backward button to shuffle through selected items from within the edit form itself. SEE WHAT NED THINKS then maybe do it!


        PopulateEditComboboxes()
        'Branch to form update routine
        UpdateItemFields(Main.AllItemsLISTBOX.SelectedIndex)



    End Sub
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    'UPDATES ALL EDIT ITEM FIELD COMBO /TEXT BOXES WIT SELECTED ITEMS DATA
    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    Sub UpdateItemFields(ItemIndexNumber)

        EditItemNameTEXTBOX.Text = ItemObjects(ItemIndexNumber).ItemName
        EditItemBaseCOMBOBOX.Text = ItemObjects(ItemIndexNumber).ItemBase
        EditItemQualityCOMBOBOX.Text = ItemObjects(ItemIndexNumber).ItemQuality
        EditItemRunewordCheckBox.Checked = ItemObjects(ItemIndexNumber).RuneWord
        EditItemSocketsCOMBOBOX.Text = ItemObjects(ItemIndexNumber).Sockets
        EditItemEtherealItemCHECKBOX.Checked = ItemObjects(ItemIndexNumber).EtherealItem
        EditItemThrowDamageMinTEXTBOX.Text = ItemObjects(ItemIndexNumber).ThrowDamageMin
        EditItemThrowDamageMaxTEXTBOX.Text = ItemObjects(ItemIndexNumber).ThrowDamageMax
        EditItemOneHandDamageMinTEXTBOX.Text = ItemObjects(ItemIndexNumber).OneHandDamageMin
        EditItemOneHandDamageMaxTEXTBOX.Text = ItemObjects(ItemIndexNumber).OneHandDamageMax
        EditItemTwoHandDamageMinTEXTBOX.Text = ItemObjects(ItemIndexNumber).TwoHandDamageMin
        EditItemTwoHandDamageMaxTEXTBOX.Text = ItemObjects(ItemIndexNumber).TwoHandDamageMax
        EditItemQuantityMinTEXTBOX.Text = ItemObjects(ItemIndexNumber).QuantityMin
        EditItemQuantityMaxTEXTBOX.Text = ItemObjects(ItemIndexNumber).QuantityMax
        EditItemDurabilityMinTEXTBOX.Text = ItemObjects(ItemIndexNumber).DurabilityMin
        EditItemDurabilityMaxTEXTBOX.Text = ItemObjects(ItemIndexNumber).DurabilityMax
        EditItemDefenseTEXTBOX.Text = ItemObjects(ItemIndexNumber).Defense
        EditItemChanceToBlockTEXTBOX.Text = ItemObjects(ItemIndexNumber).ChanceToBlock
        EditItemRequiredStrengthTEXTBOX.Text = ItemObjects(ItemIndexNumber).RequiredStrength
        EditItemRequiredDexterityTEXTBOX.Text = ItemObjects(ItemIndexNumber).RequiredDexterity
        EditItemRequiredLevelTEXTBOX.Text = ItemObjects(ItemIndexNumber).RequiredLevel
        EditItemAttackClassCOMBOBOX.Text = ItemObjects(ItemIndexNumber).AttackClass
        EditItemAttackSpeedCOMBOBOX.Text = ItemObjects(ItemIndexNumber).AttackSpeed
        EditItemReqCharCOMBOBOX.Text = ItemObjects(ItemIndexNumber).RequiredCharacter
        EditItemImageTEXTBOX.Text = ItemObjects(ItemIndexNumber).ItemImage
        EditItemUserReferenceTEXTBOX.Text = ItemObjects(ItemIndexNumber).UserField
        EditItemStat1TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat1
        EditItemStat2TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat2
        EditItemStat3TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat3
        EditItemStat4TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat4
        EditItemStat5TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat5
        EditItemStat6TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat6
        EditItemStat7TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat7
        EditItemStat8TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat8
        EditItemStat9TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat9
        EditItemStat10TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat10
        EditItemStat11TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat11
        EditItemStat12TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat12
        EditItemStat13TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat13
        EditItemStat14TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat14
        EditItemStat15TEXTBOX.Text = ItemObjects(ItemIndexNumber).Stat15
        EditItemLevelTEXTBOX.Text = ItemObjects(ItemIndexNumber).Itemlevel
        EditItemMuleNameCOMBOBOX.Text = ItemObjects(ItemIndexNumber).MuleName
        EditItemMuleAccountCOMBOBOX.Text = ItemObjects(ItemIndexNumber).MuleAccount
        EditItemPickitBotCOMBOBOX.Text = ItemObjects(ItemIndexNumber).PickitAccount
        EditItemMulePassCOMBOBOX.Text = ItemObjects(ItemIndexNumber).MulePass
        EditItemImportDateTEXTBOX.Text = ItemObjects(ItemIndexNumber).ImportDate
        If ItemObjects(ItemIndexNumber).HardCore = True Then EditItemCoreTypeCOMBOBOX.Text = "HardCore"
        If ItemObjects(ItemIndexNumber).HardCore = False Then EditItemCoreTypeCOMBOBOX.Text = "SoftCore"
        EditItemRealmCOMBOBOX.Text = ItemObjects(ItemIndexNumber).ItemRealm
        EditItemLadderCheckBox.Checked = ItemObjects(ItemIndexNumber).Ladder
        EditItemImportTimeTEXTBOX.Text = ItemObjects(ItemIndexNumber).ImportTime

        'Wheres the hide pass routine lol?
        'If AppSettings.HideMulePass = True Then EditItemMulePassCOMBOBOX.Text = HidePassRoutine(ItemObjects(ItemIndexNumber).MulePass) Else EditItemMulePassCOMBOBOX.Text = ItemObjects(ItemIndexNumber).MulePass


    End Sub

  
    Sub PopulateEditComboboxes()


        'POPULATE EDIT COMBOBOXES
        EditItemMuleNameCOMBOBOX.Items.Clear() 'Next 5 lines clears out all existing combobox items
        EditItemMuleAccountCOMBOBOX.Items.Clear()
        EditItemMulePassCOMBOBOX.Items.Clear()
        EditItemAttackClassCOMBOBOX.Items.Clear()
        EditItemBaseCOMBOBOX.Items.Clear()
        EditItemQualityCOMBOBOX.Items.Clear()
        EditItemPickitBotCOMBOBOX.Items.Clear()

        For Each ItemObjectItem As Object In ItemObjects
            'All Mule Names included in the current dbase
            If ItemObjectItem.MuleName <> "" Then
                If EditItemMuleNameCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then EditItemMuleNameCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
            End If
            'All Mule Accounts included in the current dbase
            If ItemObjectItem.MuleAccount <> "" Then
                If EditItemMuleAccountCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then EditItemMuleAccountCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
            End If
            'All Mule Passwords included in the current dbase (tho are commonly all the same)
            If ItemObjectItem.MulePass <> "" Then
                If EditItemMulePassCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then EditItemMulePassCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
            End If
            'All Item Attack Classes included in the current dbase
            If ItemObjectItem.AttackClass <> "" Then
                If EditItemAttackClassCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then EditItemAttackClassCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
            End If
            'All Item Bases included in the current dbase
            If ItemObjectItem.ItemBase <> "" Then
                If EditItemBaseCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then EditItemBaseCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
            End If
            'All Item Quality
            If ItemObjectItem.ItemQuality <> "" Then
                If EditItemQualityCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then EditItemQualityCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
            End If

            If ItemObjectItem.PickitAccount <> "" Then
                If EditItemPickitBotCOMBOBOX.Items.Contains(ItemObjectItem.PickitAccount) = False Then EditItemPickitBotCOMBOBOX.Items.Add(ItemObjectItem.PickitAccount)
            End If
        Next


    End Sub

    Private Sub EditItemSaveChangesBUTTON_Click_1(sender As Object, e As EventArgs) Handles EditItemSaveChangesBUTTON.Click
        Dim ItemIndexNumber = Main.AllItemsLISTBOX.SelectedIndex
        ItemObjects(ItemIndexNumber).ItemName = EditItemNameTEXTBOX.Text
        ItemObjects(ItemIndexNumber).ItemBase = EditItemBaseCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).ItemQuality = EditItemQualityCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).RuneWord = EditItemRunewordCheckBox.Checked
        ItemObjects(ItemIndexNumber).Sockets = EditItemSocketsCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).EtherealItem = EditItemEtherealItemCHECKBOX.Checked
        ItemObjects(ItemIndexNumber).ThrowDamageMin = EditItemThrowDamageMinTEXTBOX.Text
        ItemObjects(ItemIndexNumber).ThrowDamageMax = EditItemThrowDamageMaxTEXTBOX.Text
        ItemObjects(ItemIndexNumber).OneHandDamageMin = EditItemOneHandDamageMinTEXTBOX.Text
        ItemObjects(ItemIndexNumber).OneHandDamageMax = EditItemOneHandDamageMaxTEXTBOX.Text
        ItemObjects(ItemIndexNumber).TwoHandDamageMin = EditItemTwoHandDamageMinTEXTBOX.Text
        ItemObjects(ItemIndexNumber).TwoHandDamageMax = EditItemTwoHandDamageMaxTEXTBOX.Text
        ItemObjects(ItemIndexNumber).QuantityMin = EditItemQuantityMinTEXTBOX.Text
        ItemObjects(ItemIndexNumber).QuantityMax = EditItemQuantityMaxTEXTBOX.Text
        ItemObjects(ItemIndexNumber).DurabilityMin = EditItemDurabilityMinTEXTBOX.Text
        ItemObjects(ItemIndexNumber).DurabilityMax = EditItemDurabilityMaxTEXTBOX.Text
        ItemObjects(ItemIndexNumber).Defense = EditItemDefenseTEXTBOX.Text
        ItemObjects(ItemIndexNumber).ChanceToBlock = EditItemChanceToBlockTEXTBOX.Text
        ItemObjects(ItemIndexNumber).RequiredStrength = EditItemRequiredStrengthTEXTBOX.Text
        ItemObjects(ItemIndexNumber).RequiredDexterity = EditItemRequiredDexterityTEXTBOX.Text
        ItemObjects(ItemIndexNumber).RequiredLevel = EditItemRequiredLevelTEXTBOX.Text
        ItemObjects(ItemIndexNumber).AttackClass = EditItemAttackClassCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).AttackSpeed = EditItemAttackSpeedCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).RequiredCharacter = EditItemReqCharCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).ItemImage = EditItemImageTEXTBOX.Text
        ItemObjects(ItemIndexNumber).UserField = EditItemUserReferenceTEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat1 = EditItemStat1TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat2 = EditItemStat2TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat3 = EditItemStat3TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat4 = EditItemStat4TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat5 = EditItemStat5TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat6 = EditItemStat6TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat7 = EditItemStat7TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat8 = EditItemStat8TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat9 = EditItemStat9TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat10 = EditItemStat10TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat11 = EditItemStat11TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat12 = EditItemStat12TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat13 = EditItemStat13TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat14 = EditItemStat14TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Stat15 = EditItemStat15TEXTBOX.Text
        ItemObjects(ItemIndexNumber).Itemlevel = EditItemLevelTEXTBOX.Text
        ItemObjects(ItemIndexNumber).MuleName = EditItemMuleNameCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).MuleAccount = EditItemMuleAccountCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).PickitAccount = EditItemPickitBotCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).MulePass = EditItemMulePassCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).ImportDate = EditItemImportDateTEXTBOX.Text
        If EditItemCoreTypeCOMBOBOX.Text = "HardCore" Then ItemObjects(ItemIndexNumber).HardCore = True
        If EditItemCoreTypeCOMBOBOX.Text = "SoftCore" Then ItemObjects(ItemIndexNumber).HardCore = False
        ItemObjects(ItemIndexNumber).ItemRealm = EditItemRealmCOMBOBOX.Text
        ItemObjects(ItemIndexNumber).Ladder = EditItemLadderCheckBox.Checked
        ItemObjects(ItemIndexNumber).ImportTime = EditItemImportTimeTEXTBOX.Text
        Me.Close()
    End Sub

    Private Sub ItemEdit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'POPULATE EDIT COMBOBOXES
        'UpdatingField = True
        'EditItemMuleNameCOMBOBOX.Items.Clear() 'Next 5 lines clears out all existing combobox items
        'EditItemMuleAccountCOMBOBOX.Items.Clear()
        'EditItemMulePassCOMBOBOX.Items.Clear()
        'EditItemAttackClassCOMBOBOX.Items.Clear()
        'EditItemBaseCOMBOBOX.Items.Clear()
        'EditItemQualityCOMBOBOX.Items.Clear()
        'EditItemPickitBotCOMBOBOX.Items.Clear()

        'For Each ItemObjectItem As ItemObjects In Objects
        'If ItemObjectItem.MuleName <> "" Then
        ' If EditItemMuleNameCOMBOBOX.Items.Contains(ItemObjectItem.MuleName) = False Then EditItemMuleNameCOMBOBOX.Items.Add(ItemObjectItem.MuleName)
        'End If
        'If ItemObjectItem.MuleAccount <> "" Then
        'If EditItemMuleAccountCOMBOBOX.Items.Contains(ItemObjectItem.MuleAccount) = False Then EditItemMuleAccountCOMBOBOX.Items.Add(ItemObjectItem.MuleAccount)
        'End If
        'If ItemObjectItem.MulePass <> "" Then
        'If EditItemMulePassCOMBOBOX.Items.Contains(ItemObjectItem.MulePass) = False Then EditItemMulePassCOMBOBOX.Items.Add(ItemObjectItem.MulePass)
        'End If
        'If ItemObjectItem.AttackClass <> "" Then
        'If EditItemAttackClassCOMBOBOX.Items.Contains(ItemObjectItem.AttackClass) = False Then EditItemAttackClassCOMBOBOX.Items.Add(ItemObjectItem.AttackClass)
        'End If
        'If ItemObjectItem.ItemBase <> "" Then
        'If EditItemBaseCOMBOBOX.Items.Contains(ItemObjectItem.ItemBase) = False Then EditItemBaseCOMBOBOX.Items.Add(ItemObjectItem.ItemBase)
        'End If
        'If ItemObjectItem.ItemQuality <> "" Then
        'If EditItemQualityCOMBOBOX.Items.Contains(ItemObjectItem.ItemQuality) = False Then EditItemQualityCOMBOBOX.Items.Add(ItemObjectItem.ItemQuality)
        'End If
        'If ItemObjectItem.PickitBot <> "" Then
        'If EditItemPickitBotCOMBOBOX.Items.Contains(ItemObjectItem.PickitBot) = False Then EditItemPickitBotCOMBOBOX.Items.Add(ItemObjectItem.PickitBot)
        'End If
        'Next
        'UpdatingField = False
    End Sub
End Class