Imports System.Text.RegularExpressions
Module Tradelist
    Public Sub SendToTradeList(ByVal x)
        Dim temp As String = ""

        '***********************************************
        'Annihilus
        '***********************************************
        If ItemObjects(x).ItemBase = "Small Charm" And ItemObjects(x).ItemQuality = "Unique" Then
            If ItemObjects(x).Stat2 = "" Then Main.TradeListRICHTEXTBOX.AppendText("Anni Unid" & vbCrLf & vbCrLf) : Return
            temp = "Anni, "
            Dim temp1 = Regex.Replace(ItemObjects(x).Stat2, "[^0-9]", "") & " " & Regex.Replace(ItemObjects(x).Stat3, "[^0-9]", "") & " " & Regex.Replace(ItemObjects(x).Stat4, "[^0-9]", "")
            Main.TradeListRICHTEXTBOX.AppendText(temp & temp1 & vbCrLf & vbCrLf) : Return
        End If

        '***********************************************
        ' torch
        '***********************************************
        If ItemObjects(x).ItemName.IndexOf("Hellfire Torch") > -1 Then
            If ItemObjects(x).Stat2.IndexOf("Ama") > -1 Then temp = "Zon "
            If ItemObjects(x).Stat2.IndexOf("Druid") > -1 Then temp = "Druid "
            If ItemObjects(x).Stat2.IndexOf("Barb") > -1 Then temp = "Barb "
            If ItemObjects(x).Stat2.IndexOf("Pal") > -1 Then temp = "Pala "
            If ItemObjects(x).Stat2.IndexOf("Sorceress") > -1 Then temp = "Sorc "
            If ItemObjects(x).Stat2.IndexOf("Necro") > -1 Then temp = "Necro "
            If ItemObjects(x).Stat2.IndexOf("Assa") > -1 Then temp = "Sin "
            temp = temp + "Torch " & ItemObjects(x).Stat3 & " " & ItemObjects(x).Stat4
            temp = temp.Replace(" to all Attributes", " ")
            temp = temp.Replace(" All Resistances", " ")
            temp = temp.Replace("+", "")
            Main.TradeListRICHTEXTBOX.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If
        '***********************************************
        ' Gheeds
        '***********************************************
        If ItemObjects(x).ItemName = "Gheed's Fortune Grand Charm" Then
            temp = "Gheeds " & ItemObjects(x).Stat1 & " " & ItemObjects(x).Stat2 & " " & ItemObjects(x).Stat3
            temp = temp.Replace("% Extra Gold from Monsters", "")
            temp = temp.Replace("% Better Chance of Getting Magic Items", "")
            temp = temp.Replace("Reduces all Vendor Prices", "")
            temp = temp.Replace("%", "")
            Main.TradeListRICHTEXTBOX.AppendText(temp & vbCrLf & vbCrLf) : Return
        End If

        ' if Identified & set item go to set function
        If ItemObjects(x).ItemQuality = "Set" And ItemObjects(x).Stat1.IndexOf("Unid") = -1 Then
            temp = Set_items(x)
            If ItemObjects(x).Sockets <> "" Then temp = temp & ", Socs " & ItemObjects(x).Sockets
            If ItemObjects(x).EtherealItem = True Then temp = temp & ", Eth"
        End If

        ' if Identified & set item go to Unique function
        If ItemObjects(x).ItemQuality = "Unique" And ItemObjects(x).Stat1.IndexOf("Unid") = -1 Then
            temp = Uniq_items(x) : If ItemObjects(x).Sockets <> 0 Then temp = temp & ", Socs " & ItemObjects(x).Sockets
            If ItemObjects(x).EtherealItem = True Then temp = temp & ", Eth"
        End If


        If ItemObjects(x).RuneWord = "True" Then temp = RuneWord_items(x) : If ItemObjects(x).EtherealItem = True Then temp = temp & ", Eth"


        '***********************************************
        'Rare/magic/white items
        '***********************************************
        If temp = "" Then
            temp = ItemObjects(x).ItemName & ", " 'sets the default name

            Select Case (ItemObjects(x).ItemBase)
                Case "Armor", "Helm", "Belt", "Shield", "Boots", "Gloves"
                    If IsNumeric(ItemObjects(x).Defense) Then temp = temp & "Def " & ItemObjects(x).Defense
                Case "Small Charm"
                    temp = "SC"
                Case "Large Charm"
                    temp = "LC"
                Case "Grand Charm"
                    temp = "GC"
            End Select
            temp = temp & Stats_items(x)
        End If

        'filters -> abbreviations 
        temp = temp.Replace("Mana stolen per hit", "Moh")
        temp = temp.Replace("Magic Damage Reduced by", "Mdr")
        temp = temp.Replace("Damage Reduced by", "Dr")
        temp = temp.Replace("Lightning Resistance", "Lr")
        temp = temp.Replace("Lightning Resist", "Lr")
        temp = temp.Replace("Cold Resistance", "Cr")
        temp = temp.Replace("Cold Resist", "Cr")
        temp = temp.Replace("Fire Resistance", "Fr")
        temp = temp.Replace("Fire Resist", "Fr")
        temp = temp.Replace("Poison Resistance", "Pr")
        temp = temp.Replace("Poison Resist", "Pr")
        temp = temp.Replace("to Lightning Skill Damage", "to Lit dmg")
        temp = temp.Replace("Sorceress Skill Levels", "Sorc Skills")
        temp = temp.Replace("Paladin Skill Levels", "Pala Skills")
        temp = temp.Replace("Amazon Skill Levels", "Zon Skills")
        temp = temp.Replace("Necromancer Skill Levels", "Nec Skills")
        temp = temp.Replace("Barbarian Skill Levels", "Barb Skills")
        temp = temp.Replace("Druid Skill Levels", "Druid Skills")
        temp = temp.Replace("to Dexterity", "to Dex")
        temp = temp.Replace("Better Chance of Getting Magic Items", "Mf")
        temp = temp.Replace("Ethereal (Cannot be Repaired)", "Eth")
        temp = temp.Replace("Extra Gold from Monsters", "Gf")
        temp = temp.Replace("All Resistances", "Res All")
        temp = temp.Replace("to Experience Gained", "Xp")
        temp = temp.Replace("Faster Cast Rate", "fcr")
        temp = temp.Replace("Enhanced Defense", "Ed")
        temp = temp.Replace("Increase Maximum Durability", "Edur")
        temp = temp.Replace("Enhanced Damage", "Edmg")
        temp = temp.Replace("Faster Hit Recovery", "Fhr")
        temp = temp.Replace("Faster Run/Walk", "Frw")
        temp = temp.Replace("Attack Rating", "Ar")
        temp = temp.Replace("Faster Block Rate", "Fbr")
        temp = temp.Replace("Increased Attack Speed", "Ias")
        temp = temp.Replace("Strength", "Str")
        temp = temp.Replace("Vitality", "Vit")
        temp = temp.Replace("Life stolen per hit", "Loh")
        temp = temp.Replace("Regenerate Mana", "Mana Regen")
        temp = temp.Replace("Life stolen per hit", "Loh")
        temp = temp.Replace("Meditation Aura When Equipped", "Med")
        temp = temp.Replace("Chains of Honor", "COH")
        temp = temp.Replace("Chance of Crushing Blow", "Cb")


        temp = temp.Replace("Maximum Damage", "Max dmg")
        temp = temp.Replace("Defense", "Def")
        temp = temp.Replace("Fire Skill Damage", "Fire dmg")
        temp = temp.Replace("Increased Chance of Blocking", "Icb")
        temp = temp.Replace("Fire Ball", "Fb")
        temp = temp.Replace("Replenish", "Rep")
        temp = temp.Replace("Damage", "Dmg")
        temp = temp.Replace("Character", "Char")
        temp = temp.Replace("Level", "Lvl")
        temp = temp.Replace("damage", "dmg")
        temp = temp.Replace("Lightning", "Lit")
        temp = temp.Replace("lightning", "Lit")
        temp = temp.Replace("Poison and Bone", "PnB")
        temp = temp.Replace("poison", "Psn")
        temp = temp.Replace("Poison", "Psn")
        temp = temp.Replace("seconds", "Secs")
        temp = temp.Replace("Minimum", "Min")
        temp = temp.Replace("Maximum", "Max")
        temp = temp.Replace("Stamina", "Stam")
        temp = temp.Replace("Defensive", "Def")
        temp = temp.Replace("Skeleton", "Skel")

        temp = temp.Replace("Socketed", "Socs")
        temp = temp.Replace("Unidentified", "Unid")
        temp = temp.Replace("Unique", "Uniq")
        temp = temp.Replace("Ethereal", "Eth")


        temp = temp.Replace("+", "")
        temp = temp.Replace("(", "")
        temp = temp.Replace(")", "")
        temp = temp.Replace("Sorceress", "Sorc")
        temp = temp.Replace("Paladin", "Pala")
        temp = temp.Replace("Necromancer", "Nec")
        temp = temp.Replace("Barbarian", "Barb")
        'temp = temp.Replace("Druid Only", "Druid") ' not worth trimming
        temp = temp.Replace("Assassin", "Sin")
        temp = temp.Replace("Amazon", "Zon")
        temp = temp.Replace("Only", "")
        temp = temp.Replace(", ,", ",") ' couple extra "," slip in for some reason
        temp = temp.Replace(",,", ",") ' couple extra "," slip in for some reason

        temp = temp.Replace("Adds", "+") ' leave at end of filters/abbreviations


        Main.TradeListRICHTEXTBOX.AppendText(temp & vbCrLf & vbCrLf)
    End Sub

    Function Stats_items(ByRef x)
        Dim temp As String = ""
        If ItemObjects(x).Stat1 <> "" Then temp = temp & ", " & ItemObjects(x).Stat1
        If ItemObjects(x).Stat2 <> "" Then temp = temp & ", " & ItemObjects(x).Stat2
        If ItemObjects(x).Stat3 <> "" Then temp = temp & ", " & ItemObjects(x).Stat3
        If ItemObjects(x).Stat4 <> "" Then temp = temp & ", " & ItemObjects(x).Stat4
        If ItemObjects(x).Stat5 <> "" Then temp = temp & ", " & ItemObjects(x).Stat5
        If ItemObjects(x).Stat6 <> "" Then temp = temp & ", " & ItemObjects(x).Stat6
        If ItemObjects(x).Stat7 <> "" Then temp = temp & ", " & ItemObjects(x).Stat7
        If ItemObjects(x).Stat8 <> "" Then temp = temp & ", " & ItemObjects(x).Stat8
        If ItemObjects(x).Stat9 <> "" Then temp = temp & ", " & ItemObjects(x).Stat9
        If ItemObjects(x).Stat10 <> "" Then temp = temp & ", " & ItemObjects(x).Stat10
        If ItemObjects(x).Stat11 <> "" Then temp = temp & ", " & ItemObjects(x).Stat11
        If ItemObjects(x).Stat12 <> "" Then temp = temp & ", " & ItemObjects(x).Stat12
        If ItemObjects(x).Stat13 <> "" Then temp = temp & ", " & ItemObjects(x).Stat13
        If ItemObjects(x).Stat14 <> "" Then temp = temp & ", " & ItemObjects(x).Stat14
        If ItemObjects(x).Stat15 <> "" Then temp = temp & ", " & ItemObjects(x).Stat15

        Return temp
    End Function

    Function RuneWord_items(ByRef x)
        Dim temp As String = ""

        '***********************************************
        ' Runeword Armor
        '***********************************************
        If ItemObjects(x).ItemBase = "Armor" Then
            Return ItemObjects(x).ItemName & " Def " & ItemObjects(x).Defense

        End If

        '***********************************************
        ' Runeword Shields
        '***********************************************
        If ItemObjects(x).ItemBase = "Shield" Or ItemObjects(x).ItemBase = "Auric Shield" Then
            If ItemObjects(x).ItemName.IndexOf("Spirit") > -1 Then Return ItemObjects(x).ItemName & " Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName.IndexOf("Splendor") > -1 Then Return ItemObjects(x).ItemName & " Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Runeword Weapons
        '***********************************************
        If ItemObjects(x).ItemBase = "polearm" Or ItemObjects(x).ItemBase = "Sword" Or ItemObjects(x).ItemBase = "Mace" Or ItemObjects(x).ItemBase = "Axe" Then
            If ItemObjects(x).ItemName.IndexOf("Beast") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat10
            If ItemObjects(x).ItemName.IndexOf("Breath") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat10 & ", " & ItemObjects(x).Stat11
            If ItemObjects(x).ItemName.IndexOf("Insight") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName.IndexOf("Spirit") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName.IndexOf("Oak") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat11
            If ItemObjects(x).ItemName.IndexOf("Arms") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7 & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName.IndexOf("Grief") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName.IndexOf("Infinity") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName.IndexOf("Last Wish") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName.IndexOf("Obedience") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat12
            If ItemObjects(x).ItemName.IndexOf("Oath") > -1 Then
                If ItemObjects(x).Stat2 = "Indestructible" Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat10
                Return ItemObjects(x).ItemName & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat11
            End If
            If ItemObjects(x).ItemName.IndexOf("Justice") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat6

        End If

        If ItemObjects(x).ItemBase = "Staff" Then
            If ItemObjects(x).ItemName.IndexOf("Memory") > -1 Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7

        End If

        Return ItemObjects(x).ItemName & " not found"
    End Function


    Function Uniq_items(ByRef x)
        Dim temp As String = ""
        '***********************************************
        'Unique items
        If ItemObjects(x).ItemName = "Rainbow Facet Jewel" Then Return "Rainbow Facet  " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
        '***********************************************
        'Rings
        '***********************************************
        If ItemObjects(x).ItemBase = "Ring" Then
            If ItemObjects(x).ItemName = "Bul-Kathos' Wedding Band Ring" Then Return "BK Ring " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Nagelring Ring" Then Return "Nagel Ring, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "The Stone of Jordan Ring" Then Return "Soj"
            If ItemObjects(x).ItemName = "Carrion Wind Ring" Then Return "Carrion Ring, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Dwarf Star Ring" Then Return ItemObjects(x).ItemName & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Manald Heal Ring" Then Return "Manald Heal Ring, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Raven Frost Ring" Then Return "Raven Frost, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Wisp Projector Ring" Then Return "Wisp Ring, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Nature's Peace Ring" Then Return "Nature's Peace Ring, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
        End If
        '***********************************************
        'Amulets
        '***********************************************
        If ItemObjects(x).ItemBase = "Amulet" Then
            If ItemObjects(x).ItemName = "Atma's Scarab Amulet" Then Return "Atma's Amulet, "
            If ItemObjects(x).ItemName = "Crescent Moon Amulet" Then Return "Crescent Amulet, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Highlord's Wrath Amulet" Then Return "Highlord's Amulet, "
            If ItemObjects(x).ItemName = "Mara's Kaleidoscope Amulet" Then Return "Mara's, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Metalgrid Amulet" Then Return "Metal Grid Amulet, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Nokozan Relic Amulet" Then Return "Nokozan Relic, "
            If ItemObjects(x).ItemName = "Saracen's Chance Amulet" Then Return "Saracen's Amulet, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Seraph's Hymn Amulet" Then Return "Seraph's Hymn Amulet, " & Stats_items(x)
            If ItemObjects(x).ItemName = "The Cat's Eye Amulet" Then Return "Cat's Eye Amulet, "
            If ItemObjects(x).ItemName = "The Eye of Etlich Amulet" Then Return "Etlich Amulet, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "The Mahim-Oak Curio Amulet" Then Return "Mahim-Oak Amulet, "
            If ItemObjects(x).ItemName = "The Rising Sun Amulet" Then Return "Rising Sun Amulet, " & ItemObjects(x).Stat1
        End If
        '***********************************************
        ' Unique Helms
        '***********************************************

        If ItemObjects(x).ItemBase = "Helm" Or ItemObjects(x).ItemBase = "Circlet" Or ItemObjects(x).ItemBase = "Primal Helm" Or ItemObjects(x).ItemBase = "Pelt" Then
            If ItemObjects(x).ItemName = "Andariel's Visage Demonhead" Then Return "Andies, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Arreat's Face Slayer Guard" Then Return "Arreats, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat6 & " " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Arreat's Face Guardian Crown" Then Return "Arreats, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat6 & " " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Biggin's Bonnet Cap" Then Return "Biggins, Def" & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Blackhorn's Face Death Mask" Then Return "Blackhorns, Def" & ItemObjects(x).Defense & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Cerebus' Bite Blood Spirit" Then Return "Cerebus, Def" & ItemObjects(x).Defense & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Coif of Glory Helm" Then Return "Coif of Glory, Def" & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Crown of Ages Corona" Then Return "COA, " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Crown of Thieves Corona" Then Return "Crown of Thieves, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Darksight Helm Bassinet" Then Return "Darksight, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Demonhorn's Edge Destroyer Helm" Then Return "Demonhorn, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Duskdeep Full Helm" Then Return "Duskdeep, Def" & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Giant Skull Bone Visage" Then Return "Giant Skull, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Griffon's Eye Diadem" Then Return "Griffon's Eye, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Halaberd's Reign Conqueror Crown" Then Return "Halaberd, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Harlequin Crest Shako" Then Return "Shako, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Howltusk Great Helm" Then Return "Howltusk, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Jalal's Mane Totemic Mask" Then Return "Jalal's, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Kira's Guardian Tiara" Then Return "Kira's, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Kira's Guardian Diadem" Then Return "Kira's, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Nightwing's Veil Spired Helm" Then Return "Nightwings, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Peasant Crown War Hat" Then Return "Peasant Hat , Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Peasant Crown Shako" Then Return "Peasant Shako , Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Ravenlore Sky Spirit" Then Return "Ravenlore, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Rockstopper Sallet" Then Return "Rockstopper, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & " " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Spirit Keeper Earth Spirit" Then Return "Spirit Keeper, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Steelskull Casque" Then Return "Steelskull, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Steel Shade Armet" Then Return "Steel Shade, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "The Face of Horror Mask" Then Return "Face of Horror, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tarnhelm Skull Cap" Then Return "Tarnhelm, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Undead Crown Crown" Then Return "Undead Crown, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Valkyrie Wing Winged Helm" Then Return "Valkyrie, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Vampire Gaze Grim Helm" Then Return "Vampire Gaze, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Vampire Gaze Bone Visage" Then Return "Vampire Gaze, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Veil of Steel Spired Helm" Then Return "Veil of Steel, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Wolfhowl Fury Visor" Then Return "Wolfhowl, " & Stats_items(x)
            If ItemObjects(x).ItemName = "Wormskull Bone Helm" Then Return "Wolfhowl, Def " & ItemObjects(x).Defense
        End If
        '***********************************************
        'Belts
        '***********************************************
        If ItemObjects(x).ItemBase = "Belt" Then
            If ItemObjects(x).ItemName = "Arachnid Mesh Spiderweb Sash" Then Return "Arach, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Bladebuckle Plated Belt" Then Return "Arach, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Gloom's Trap Mesh Belt" Then Return "Gloom's Belt, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Goldwrap Heavy Belt" Then Return "Goldwrap, Def " & ItemObjects(x).Defense & ", " & Stats_items(x)
            If ItemObjects(x).ItemName = "Goldwrap Troll Belt" Then Return "Goldwrap Tb, Def " & ItemObjects(x).Defense & ", " & Stats_items(x)
            If ItemObjects(x).ItemName = "Lenymo Spiderweb Sash" Then Return "Lenymo Sash, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Nightsmoke Belt" Then Return "Nightsmoke, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Nosferatu's Coil Vampirefang Belt" Then Return "Nosferatu's, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Razortail Vampirefang Belt" Then Return "Razortail, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Snakecord Light Belt" Then Return "Snakecord, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Snowclash Battle Belt" Then Return "Snowclash, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "String of Ears Demonhide Sash" Then Return "String of Ears, " & Stats_items(x)
            If ItemObjects(x).ItemName = "Thundergod's Vigor War Belt" Then Return "Thundergod's Belt, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Thundergod's Vigor Colossus Girdle" Then Return "Thundergod's Girdle, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Verdungo's Hearty Cord Mithril Coil" Then Return "Dungo's, Def " & ItemObjects(x).Defense & Stats_items(x)
        End If

        '***********************************************
        'Gloves
        '***********************************************
        If ItemObjects(x).ItemBase = "Gloves" Then
            If ItemObjects(x).ItemName = "Bloodfist Heavy Gloves" Then Return "Bloodfist, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Bloodfist Vampirebone Gloves" Then Return "Bloodfist, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Chance Guards Chain Gloves" Then Return "Chancies, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Chance Guards Vambraces" Then Return "Chancies, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Dracul's Grasp Vampirebone Gloves" Then Return "Dracs, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Frostburn Gauntlets" Then Return "Frostburn, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Ghoulhide Heavy Bracers" Then Return "Ghoulhide, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Gravepalm Sharkskin Gloves" Then Return "Gravepalm, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Hellmouth War Gauntlets" Then Return "Hellmouth, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Magefist Light Gauntlets" Then Return "Magefist, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Magefist Crusader Gauntlets" Then Return "Magefist Crusader, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Magefist Battle Gauntlets" Then Return "Magefist Battle, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Soul Drainer Vambraces" Then Return "Soul Drainer, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Steelrend Ogre Gauntlets" Then Return "Steelrend, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "The Hand of Brock" Then Return "Hand of Brock, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Venom Grip Demonhide Gloves" Then Return "Venom Grip, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Lava Gout Crusader Gauntlets" Then Return "Lava Gout Crusader Gauntlets, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Lava Gout Battle Gauntlets" Then Return "Lava Gout Battle Gauntlets, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
        End If

        '***********************************************
        'Boots
        '***********************************************
        If ItemObjects(x).ItemBase = "Boots" Then
            If ItemObjects(x).ItemName = "Goblin Toe Light Plated Boots" Then Return "Goblin Toe, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Goblin Toe Mirrored Boots" Then Return "Goblin Toe, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Gore Rider War Boots" Then Return "Gore Rider's, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Gore Rider Myrmidon Greaves" Then Return "Gore Rider's Upp'd, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Gorefoot Heavy Boots" Then Return "Gorefoot, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Hotspur Boots" Then Return "Hotspur, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Hotspur Wyrmhide Boots" Then Return "Hotspur, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Infernostride Deminhide Boots" Then Return "Infernostride, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName = "Marrowwalk Boneweave Boots" Then Return "Marrowwalk, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Sandstorm Trek Scarabshell Boots" Then Return "Treks, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Shadow Dancer Myrmidon Greaves" Then Return "Shadow Dancer, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Silkweave Mesh Boots" Then Return "Silkweave, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Silkweave Boneweave Boots" Then Return "Silkweave, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Tearhaunch Greaves" Then temp = "Tearhaunch, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Treads of Cthon Chain Boots" Then Return "Treads of Cthon, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "War Traveler Battle Boots" Then Return "War Travs, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName = "Waterwalk Sharkskin Boots" Then Return "Waterwalks, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Waterwalk Scarabshell Boots" Then Return "Waterwalks, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
        End If

        '***********************************************
        ' Unique Armor
        '***********************************************
        If ItemObjects(x).ItemBase = "Armor" Then
            If ItemObjects(x).ItemName = "Arkaine's Valor Balrog Skin" Then Return "Arkaine's Valor, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Atma's Wail Embossed Plate" Then Return "Atma's Wail, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Black Hades Chaos Armor" Then Return "Black Hades, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Blinkbat's Form Leather Armor" Then Return "Blinkbat's, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Boneflesh Plate Mail" Then Return "Boneflesh, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Corpsemourn Ornate Plate" Then Return "Corpsemourn, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Crow Caw Tigulated Mail" Then Return "Crow Caw, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Crow Caw Loricated Mail" Then Return "Crow Caw, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Darkglow Ring Mail" Then Return "Darkglow, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Duriel's Shell Cuirass" Then Return "Duriel's Shell, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Duriel's Shell Great Hauberk" Then Return "Duriel's Shell, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Goldskin Full Plate Mail" Then Return "Goldskin, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Goldskin Shadow Plate" Then Return "Goldskin, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Greyform Quilted Armor" Then Return "Greyform, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Guardian Angel Templar Coat" Then Return "Guardian Angel, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Guardian Angel Hellforge Plate" Then Return "Guardian Angel, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Hawkmail Scale Mail" Then Return "Hawkmail, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Hawkmail Loricated Mail" Then Return "Hawkmail, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Heavenly Garb Light Plate" Then Return "Heavenly Garb, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Iceblink Splint Mail" Then Return "Iceblink, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Iron Pelt Trellised Armor" Then Return "Iron Pelt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Leviathan Kraken Shell" Then Return "Leviathan, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Ormus' Robes Dusk Shroud" Then Return "Ormus Robes, " & Stats_items(x)
            If ItemObjects(x).ItemName = "Que-Hegan's Wisdom Mage Plate" Then Return "Que-Hegan's, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Rattlecage Gothic Plate" Then Return "Rattlecage, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Rockfleece Field Plate" Then Return "Rockfleece, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Shaftstop Mesh Armor" Then Return "Shaftstop Armor, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Shaftstop Boneweave" Then Return "Shaftstop bw Armor, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Silks of the Victor Ancient Armor" Then Return "Silks of the Victor, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Skin of the Flayed One Demonhide Armor" Then Return "Skin of the Flayed One, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Skin of the Flayed One Scarab Husk" Then Return "Skin of the Flayed One, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Skin of the Vipermagi Serpentskin Armor" Then Return "Vipermagi,  " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Skin of the Vipermagi Wyrmhide" Then Return "Vipermagi, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Skullder's Ire Russet Armor" Then Return "Skullders,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Skullder's Ire Balrog Skin" Then Return "Skullders,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Sparking Mail Chain Mail" Then Return "Sparking Mail,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Spirit Forge Diamond Mail" Then Return "Spirit Forge,  " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Steel Carapace Shadow Plate" Then Return "Steel Carapace,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Templar's Might Sacred Armor" Then Return "Templar's Might,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "The Centurion Hard Leather Armor" Then Return "The Centurion,  " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "The Gladiator's Bane Wire Fleece" Then Return "Gladiator's Bane,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "The Spirit Shroud Ghost Armor" Then Return "Spirit Shroud,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Toothrow Sharktooth Armor" Then Return "Toothrow,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Twitchthroe Studded Leather" Then Return "Twitchthroe SL,  " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Twitchthroe Wire Fleece" Then Return "Twitchthroe wf,  " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tyrael's Might Sacred Armor" Then Return "Tyrael's Might,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Venom Ward Breast Plate" Then Return "Venom Ward,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Venom Ward Great Hauberk" Then Return "Venom Ward,  " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat1
        End If

        '***********************************************
        'Shields
        '***********************************************
        If ItemObjects(x).ItemBase = "Shield" Or ItemObjects(x).ItemBase = "Auric Shield" Or ItemObjects(x).ItemBase = "Voodoo Heads" Then
            If ItemObjects(x).ItemName = "Alma Negra Sacred Rondache" Then Return "Alma Negra, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Blackoak Shield Luna" Then Return "Blackoak Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Boneflame Succubus Skull" Then Return "Boneflame, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Bverrit Keep Tower Shield" Then Return "Bverrit Keep, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Darkforce Spawn Bloodlord Skull" Then Return "Darkforce, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Dragonscale Zakarum Shield" Then Return "Dragonscale, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Gerke's Sanctuary Pavise" Then Return "Gerke's Sanctuary, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Head Hunter's Glory Troll Nest" Then Return "Head Hunter's Glory, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Herald of Zakarum Gilded Shield" Then Return "Hoz, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Herald of Zakarum Zakarum Shield" Then Return "Hoz, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Homunculus Hierophant Trophy" Then Return "Homunculus, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Lance Guard Barbed Shield" Then Return "Lance Guard, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Lance Guard Blade Barrier" Then Return "Lance Guard, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Lidless Wall Grim Shield" Then Return "Lidless Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Lidless Wall Troll Nest" Then Return "Lidless TN, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Medusa's Gaze Aegis" Then Return "Medusa's Gaze, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Moser's Blessed Circle Round Shield" Then Return "Moser's Blessed Circle, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Pelta Lunata Buckler" Then Return "Pelta Lunata, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Spike Thorn Blade Barrier" Then Return "Pelta Lunata, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Spirit Ward Ward" Then Return "Spirit Ward, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Steelclash Kite Shield" Then Return "Steelclash, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Stormguild Large Shield" Then Return "Stormguild, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Stormshield Monarch" Then Return "Stormshield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Swordback Hold Spiked Shield" Then Return "Swordback Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Swordback Hold Blade Barrier" Then Return "Swordback Barrier, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "The Ward Gothic Shield" Then Return "The Ward, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Tiamat's Rebuke Dragon Shield" Then Return "Tiamat's Rebuke, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName = "Umbral Disk Small Shield" Then Return "Umbral Disk, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Visceratuant Defender" Then Return "Visceratuant, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Wall of the Eyeless Bone Shield" Then Return "Wall of the Eyeless, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        'Weapons Axes
        '***********************************************
        If ItemObjects(x).ItemBase = "Axe" Then
            If ItemObjects(x).ItemName = "Axe of Fechmar" Then Return "Fechmars Axe, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Bladebone Double Axe" Then Return "Bladebone, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Boneslayer Blade Gothic Axe" Then Return "Boneslayer, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Brainhew Great Axe" Then Return "Brainhew, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Butcher's Pupil Cleaver" Then Return "Butcher's Pupil, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Cranebeak War Spike" Then Return "Cranebeak, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Death Cleaver Berserker Axe" Then Return "Death Cleaver, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Deathspade Axe" Then Return "Deathspade, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Ethereal Edge Silver-edged Axe" Then Return "Ethereal Edge, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Executioner's Justice Glorious Axe" Then Return "Executioner's Justice, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Gimmershred Flying Axe" Then Return "Gimmershred, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Goreshovel Broad Axe" Then Return "Goreshovel, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Hellslayer Decapitator" Then Return "Hellslayer"
            If ItemObjects(x).ItemName = "Homongous Giant Axe" Then Return "Homongous, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Islestrike Twin Axe" Then Return "Islestrike, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Lacerator Winged Axe" Then Return "Lacerator, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Messerschmidt's Reaver Champion Axe" Then Return "Messerschmidt's"
            If ItemObjects(x).ItemName = "Pompeii's Wrath War Spike" Then Return "Pompeii Spike, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Rakescar War Axe" Then Return "Rakescar, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Razor's Edge Tomahawk" Then Return "Razor's Edge, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Rune Master Ettin Axe" Then Return "Rune Master, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Skull Splitter Military Pick" Then Return "Skull Splitter, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Stormrider Tabar" Then Return "Stormrider"
            If ItemObjects(x).ItemName = "The Chieftain Battle Axe" Then Return "The Chieftain, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "The Gnasher Hand Axe" Then Return "The Gnasher, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "The Minotaur Ancient Axe" Then Return "The Minotaur, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "The Scalper Francisca" Then Return "The Scalper, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
        End If

        '***********************************************
        'Weapons Bows
        '***********************************************
        If ItemObjects(x).ItemBase = "Amazon Bow" Or ItemObjects(x).ItemBase = "Bow" Or ItemObjects(x).ItemBase = "Crossbow" Then
            If ItemObjects(x).ItemName = "Blastbark Long War Bow" Then Return "Blastbark, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Blood Raven's Charge Matriarchal Bow" Then Return "Blood Raven's Charge, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Cliffkiller Large Siege Bow" Then Return "Cliffkiller, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Demon Machine Chu-Ko-Nu" Then Return "Chu-Ko-Nu, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Eaglehorn Crusader Bow" Then Return "Eaglehorn, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Goldstrike Arch Gothic Bow" Then Return "Goldstrike, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Hellclap Short War Bow" Then Return "Hellclap, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Kuko Shakaku Cedar Bow" Then Return "Kuko Shakaku, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Lycander's Aim Ceremonial Bow" Then Return "Lycander's, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Magewrath Rune Bow" Then Return "Magewrath, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Pluckeye Short Bow" Then Return "Pluckeye"
            If ItemObjects(x).ItemName = "Raven Claw Long Bow" Then Return "Raven Claw, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Riphook Razor Bow" Then Return "Riphook RB, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Riphook Blade Bow" Then Return "Riphook BB, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Rogue's Bow Composite Bow" Then Return "Rogue's Bow, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Stormstrike Short Battle Bow" Then Return "Stormstrike, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Widowmaker Ward Bow" Then Return "Widowmaker, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Windforce Hydra Bow" Then Return "Windforce, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Witchwild String Short Siege Bow" Then Return "Witchwild, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Witherstring Hunter's Bow" Then Return "Witherstring, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Wizendraw Long Battle Bow" Then Return "Wizendraw, " & ItemObjects(x).Stat3
        End If

        '***********************************************
        'Weapons Claws
        '***********************************************
        If ItemObjects(x).ItemBase = "Assassin Claw" Then
            If ItemObjects(x).ItemName = "Bartuc's Cut-Throat Greater Talons" Then Return "Bartuc's, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Firelizard's Talons Feral Claws" Then Return "Firelizard's, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Jade Talon Wrist Sword" Then Return "Jade Talon, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Shadow Killer Battle Cestus" Then Return "Shadow Killer, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Shadow Killer Battle Cestus" Then Return "Shadow Killer, " & ItemObjects(x).Stat3


        End If

        '***********************************************
        'Weapons Crossbows
        '***********************************************
        If ItemObjects(x).ItemBase = "Crossbow" Then
            If ItemObjects(x).ItemName = "Buriza-Do Kyanon Ballista" Then Return "Buriza-Do Kyanon, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Buriza-Do Kyanon Colossus Crossbow" Then Return "Buriza-Do Kyanon, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Doomslinger Repeating Crossbow" Then Return "Doomslinger, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Hellcast Heavy Crossbow" Then Return "Hellcast, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Hellrack Colossus Crossbow" Then Return "Hellrack, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Ichorsting Crossbow" Then Return "Ichorsting"
            If ItemObjects(x).ItemName = "Leadcrow Light Crossbow" Then Return "Leadcrow"
        End If
        '***********************************************
        'Weapons Daggers
        '***********************************************
        If ItemObjects(x).ItemBase = "Dagger" Or ItemObjects(x).ItemBase = "Throwing Knife" Then
            If ItemObjects(x).ItemName = "Blackbog's Sharp Cinquedeas" Then Return "Blackbog's "
            If ItemObjects(x).ItemName = "Deathbit Battle Dart" Then Return "Deathbit, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Fleshripper Fanged Knife" Then Return "Fleshripper, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Ghostflame Legend Spike" Then Return "Ghostflame, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Gull Dagger" Then Return "Gull Dagger"
            If ItemObjects(x).ItemName = "Heart Carver Rondel" Then Return "Heart Carver, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Spectral Shard Blade" Then Return "Spectral Shard"
            If ItemObjects(x).ItemName = "Spineripper Poignard" Then Return "Spineripper, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "The Diggler Dirk" Then Return "The Diggler"
            If ItemObjects(x).ItemName = "The Jade Tan Do Kris" Then Return "The Jade Tan Do, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Wizardspike" Then Return "Wizardspike"
            If ItemObjects(x).ItemName = "Warshrike Winged Knife" Then Return "Warshrike, " & ItemObjects(x).Stat4
        End If
        '***********************************************
        'Weapons Javelins
        '***********************************************
        If ItemObjects(x).ItemBase = "Javelin" Or ItemObjects(x).ItemBase = "Amazon Javelin" Then
            If ItemObjects(x).ItemName = "Demon's Arch Balrog Spear" Then Return "Demon's Arch, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Gargoyle's Bite Winged Harpoon" Then Return "Gargoyle's Bite, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Thunderstroke Matriarchal Javelin" Then Return "Thunderstroke, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Titan's Revenge Ceremonial Javelin" Then Return "Titan's Revenge CJ, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Titan's Revenge Matriarchal Javelin" Then Return "Titan's Revenge MJ, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Wraith Flight Ghost Glaive" Then Return "Wraith Flight, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
        End If
        '***********************************************
        'Weapons Maces
        '***********************************************
        If ItemObjects(x).ItemBase = "Knife" Then
            If ItemObjects(x).ItemName = "Wizardspike Bone Knife" Then Return "Wizardspike"
        End If
        '***********************************************
        'Weapons Maces
        '***********************************************
        If ItemObjects(x).ItemBase = "Mace" Or ItemObjects(x).ItemBase = "Club" Or ItemObjects(x).ItemBase = "Hammer" Then
            If ItemObjects(x).ItemName = "Baezil's Vortex Knout" Then Return "Baezil's Vortex, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Baranar's Star Devil Star" Then Return "Baranar's Star"
            If ItemObjects(x).ItemName = "Bloodrise Morning Star" Then Return "Bloodrise"
            If ItemObjects(x).ItemName = "Bloodtree Stump War Club" Then Return "Bloodtree Stump, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Bonesnap Maul" Then Return "Bonesnap, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Crushflange Mace" Then Return "Crushflangep, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Demon Limb Tyrant Club" Then Return "Demon Limb, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Earth Shifter Thunder Maul" Then Return "Earth Shifter, " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Earthshaker" Then Return "Earthshaker"
            If ItemObjects(x).ItemName = "Felloak Club" Then Return "Felloak, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Fleshrender Barbed Club" Then Return "Fleshrender, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Horizon's Tornado Scourge" Then Return "Horizon's Tornado, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Ironstone War Hammer" Then Return "Ironstone, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Schaefer's Hammer Legendary Mallet" Then Return "Schaefer's Hammer, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Nord's Tenderizer Truncheon" Then Return "Nord's Tenderizer, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Steeldriver Great Maul" Then Return "Steeldriver, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Stone Crusher Legendary Mallet" Then Return "Stone Crusher, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Stormlash Scourge" Then Return "Stormlash, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Stoutnail Spiked Club" Then Return "Stoutnail"
            If ItemObjects(x).ItemName = "The Cranium Basher Thunder Maul" Then Return "The Cranium Basher, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "The Gavel of Pain Martel de Fer" Then Return "The Gavel of Pain, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "The General's Tan Do Li Ga Flail" Then Return "The General's Flail, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "The General's Tan Do Li Ga Scourge" Then Return "The General's Scourge, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Windhammer Ogre Maul" Then Return "Windhammer, " & ItemObjects(x).Stat3
        End If


        '***********************************************
        'Weapons orbs
        '***********************************************
        If ItemObjects(x).ItemBase = "Orb" Then
            If ItemObjects(x).ItemName = "Death's Fathom Dimensional Shard" Then Return "Death's Fathom, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Eschuta's Temper Eldritch Orb" Then Return "Eschuta's, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "The Oculus Swirling Crystal" Then Return "Oculus"
        End If

        '***********************************************
        'Weapons Polearms
        '***********************************************
        If ItemObjects(x).ItemBase = "polearm" Then
            If ItemObjects(x).ItemName = "Bonehew Ogre Axe" Then Return "Bonehew, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Dimoak's Hew " Then Return "Dimoak's Hew, "
            If ItemObjects(x).ItemName = "Pierre Tombale Couant Partizan" Then Return "Pierre Tombale, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Soul Harvest Scythe" Then Return "Soul Harvest, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Steelgoad Voulge" Then Return "Steelgoad, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Stormspire Giant Thresher" Then Return "Stormspire, " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "The Battlebranch Poleaxe" Then Return "The Battlebranch, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "The Grim Reaper War Scythe" Then Return "The Grim Reaper, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "The Reaper's Toll Thresher" Then Return "Reaper's Toll, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Tomb Reaver Cryptic Axe" Then Return "Tomb Reaver, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName = "Woestave Halberd" Then Return "Woestave, " & ItemObjects(x).Stat1
        End If

        '***********************************************
        'Weapons Sceptors
        '***********************************************
        If ItemObjects(x).ItemBase = "Scepter" Then
            If ItemObjects(x).ItemName = "Astreon's Iron Ward Caduceus" Then Return "Astreon's Iron Ward, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat9
            If ItemObjects(x).ItemName = "Hand of Blessed Light Divine Sceptor" Then Return "Hand of Blessed Light, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Heaven's Light Mighty Scepter" Then Return "Heaven's Light, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Knell Striker Scepter" Then Return "Knell Striker, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Rusthandle Grand Scepter" Then Return "Rusthandle, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Stormeye War Scepter" Then Return "Stormeye, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "The Fetid Sprinkler Holy Water Sprinkler" Then Return "The Fetid Sprinkler, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "The Redeemer Mighty Scepter" Then Return "The Redeemer, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Zakarum Hand Rune Scepter" Then Return "Zakarum Hand, " & ItemObjects(x).Stat3
        End If

        '***********************************************
        'Weapons Spears
        '***********************************************
        If ItemObjects(x).ItemBase = "Spear" Then
            If ItemObjects(x).ItemName = "Arioc's Needle Hyperion Spear" Then Return "Arioc's Needle, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Bloodthief Brandistock" Then Return "Bloodthief, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Hone Sundan Yari" Then Return "Hone Sundan, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Kelpie Snare Fuscina" Then Return "Kelpie Snare, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Lance of Yaggai Spetum" Then Return "Lance of Yaggai"
            If ItemObjects(x).ItemName = "Lycander's Flank Ceremonial Spike" Then Return "Lycander's, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Razortine Trident" Then Return "Razortine, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Soulfeast Tine War Fork" Then Return "Soulfeast, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Spire of Honor Lance" Then Return "Spire of Honor, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Steel Pillar War Pike" Then Return "Steel Pillar, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Stoneraven Matriarchal Spear" Then Return "Stoneraven, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "The Dragon Chang Spear" Then Return "The Dragon Chang"
            If ItemObjects(x).ItemName = "The Tannr Gorerod Pike" Then Return "Tannr Gorerod, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Viperfork Mancatcher" Then Return "Viperfork, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
        End If

        '***********************************************
        'Weapons Staves
        '***********************************************
        If ItemObjects(x).ItemBase = "Staff" Then
            If ItemObjects(x).ItemName = "Bone Ash" Then Return "Bone Ash"
            If ItemObjects(x).ItemName = "Chromatic Ire" Then Return "Chromatic Ire, " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Mang Song's Lesson Archon Staff" Then Return "Mang Song's Lesson, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Ondal's Wisdom Elder Staff" Then Return "Ondal's Wisdom, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Razorswitch Io Staff" Then Return "Razorswitch"
            If ItemObjects(x).ItemName = "Ribcracker Quarterstaff" Then Return "Ribcracker, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Serpent Lord Long Staff" Then Return "Serpent Lord, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Skull Collector Rune Staff" Then Return "Skull Collector"
            If ItemObjects(x).ItemName = "Spire of Lazarus Gnarled Staff" Then Return "Spire of Lazarus"
            If ItemObjects(x).ItemName = "The Iron Jang Bong War Staff" Then Return "The Iron Jang Bong"
            If ItemObjects(x).ItemName = "The Salamander Battle Staff" Then Return "The Salamander"
            If ItemObjects(x).ItemName = "Warpspear Gothic Staff" Then Return "Warpspear"

        End If

        '***********************************************
        'Weapons Swords
        '***********************************************
        If ItemObjects(x).ItemBase = "Sword" Then
            If ItemObjects(x).ItemName = "Azurewrath Phase Blade" Then Return "Azurewrath, " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Blacktongue Bastard Sword" Then Return "Blacktongue, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Blade of Ali Baba Tulwar" Then Return "Blade of Ali Baba, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Blade of Ali Baba Hydra Edge" Then Return "Blade of Ali Baba, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Blood Crescent Scimitar" Then Return "Blood Crescent, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Bloodletter Gladius" Then Return "Bloodletter, " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Coldsteel Eye Cutlass" Then Return "Coldsteel Eye, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Crainte Vomir Espondon" Then Return "Crainte Vomir, " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Culwen's Point War Sword" Then Return "Culwen's Point, " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Djinn Slayer Ataghan" Then Return "Djinn Slayer, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Doombringer Champion Sword" Then Return "Doombringer, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Flamebellow Balrog Blade" Then Return "Flamebellow, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat6 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "Frostwind Cryptic Sword" Then Return "Frostwind, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Ginter's Rift Dimensional Blade" Then Return "Ginter's Rift, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Gleamscythe Falchion" Then Return "Gleamscythe, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Griswold's Edge Broad Sword" Then Return "Griswold's Edge, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Headstriker Battle Sword" Then Return "Headstriker"
            If ItemObjects(x).ItemName = "Hellplague Long Sword" Then Return "Hellplague, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Hexfire Shamshir" Then Return "Hexfire, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Kinemil's Awl Giant Sword" Then Return "Kinemil's Awl, " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Lightsabre Phase Blade" Then Return "Lightsabre, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat8
            If ItemObjects(x).ItemName = "Plague Bearer Rune Sword" Then Return "Plague Bearer"
            If ItemObjects(x).ItemName = "Ripsaw Flamberge" Then Return "Ripsaw, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "Rixot's Keen" Then Return "Rixot's Keen" 'unfinnished
            If ItemObjects(x).ItemName = "Shadowfang Two-Handed Sword" Then Return "Shadowfang"
            If ItemObjects(x).ItemName = "Skewer of Krintiz Sabre" Then Return "Skewer of Krintiz"
            If ItemObjects(x).ItemName = "Soulflay Claymore" Then Return "Soulflay, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Swordguard Executioner Sword" Then Return "Swordguard, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat7
            If ItemObjects(x).ItemName = "The Atlantean Ancient Sword" Then Return "The Atlantean, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "The Grandfather Colossus Blade" Then Return "The Grandfather, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "The Patriarch Great Sword" Then Return "The Patriarch, " & ItemObjects(x).Stat1
            If ItemObjects(x).ItemName = "The Vile Husk Tusk Sword" Then Return "The Vile Husk, " & ItemObjects(x).Stat2
            If ItemObjects(x).ItemName = "Todesfaelle Flamme Zweihander" Then Return "Todesfaelle Flamme, " & ItemObjects(x).Stat2
        End If

        '***********************************************
        'Weapons Wands
        '***********************************************
        If ItemObjects(x).ItemBase = "Wand" Then
            If ItemObjects(x).ItemName = "Arm of King Leoric Tomb Wand" Then Return "Arm of King Leoric"
            If ItemObjects(x).ItemName = "Blackhand Key Grave Wand" Then Return "Blackhand Key"
            If ItemObjects(x).ItemName = "Boneshade Lich Wand" Then Return "Boneshade, " & Stats_items(x)
            If ItemObjects(x).ItemName = "Carin Shard Petrified Wand" Then Return "Carin Shard"
            If ItemObjects(x).ItemName = "Death's Web Unearthed Wand" Then Return "Death's Web, " & ItemObjects(x).Stat2 & ", " & ItemObjects(x).Stat3
            If ItemObjects(x).ItemName = "Gravespine Bone Wand" Then Return "Gravespine, " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Maelstrom Yew Wand" Then Return "Maelstrom, " & ItemObjects(x).Stat3 & ", " & ItemObjects(x).Stat4 & ", " & ItemObjects(x).Stat5 & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Suicide Branch Burnt Wand" Then Return "Suicide Branch"
            If ItemObjects(x).ItemName = "Torch of Iro Wand" Then Return "Torch of Iro"
            If ItemObjects(x).ItemName = "Ume's Lament Grim Wand" Then Return "Ume's Lament"
        End If

        Return ItemObjects(x).ItemName & Stats_items(x) & " [Not found]"
    End Function


    Function Set_items(ByVal x)
        Dim temp As String = ""
        'MsgBox("Set Items")
        '***********************************************
        'Amulets
        '***********************************************
        If ItemObjects(x).ItemBase = "Amulet" Then
            If ItemObjects(x).ItemName = "Tal Rasha's Adjudication Amulet" Then Return "Tal's  Amulet"
            Return ItemObjects(x).ItemName
        End If

        '***********************************************
        'Rings
        '***********************************************
        If ItemObjects(x).ItemBase = "Ring" Then
            Return ItemObjects(x).ItemName
        End If

        '***********************************************
        ' Set Helms
        '***********************************************
        If ItemObjects(x).ItemBase = "Helm" Or ItemObjects(x).ItemBase = "Circlet" Or ItemObjects(x).ItemBase = "Primal Helm" Then
            If ItemObjects(x).ItemName = "Aldur's Stony Gaze Hunter's Guise" Then Return "Aldur's Guise, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Arcanna's Head Skull Cap" Then Return "Arcanna's Cap, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Berserker's Headgear Helm" Then Return "Berserker's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Cathan's Visage Mask" Then Return "Cathan's Mask, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Griswold's Valor Corona" Then Return "Griswold's Corona, Def " & ItemObjects(x).Defense & " " & ItemObjects(x).Stat1 & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Guillaume's Face Winged Helm" Then Return "Guillaume's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Hwanin's Splendor Grand Crown" Then Return "Hwanin's Crown, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Immortal King's Will Avenger Guard" Then Return "IK Helm, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Infernal Cranium Cap" Then Return "Infernal Cap, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Iratha's Coil Crown" Then Return "Iratha's Crown, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Isenhart's Horns Full Helm" Then Return "Isenhart's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "M'avina's True Sight Diadem" Then Return "M'avina's Diadem, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Milabrega's Diadem Crown" Then Return "Milabrega's Crown, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Naj's Circlet Circlet" Then Return "Naj's Circlet, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Natalya's Totem Grim Helm" Then Return "Natalya's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Ondal's Almighty Spired Helm" Then Return "Ondal's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sander's Paragon Cap" Then Return "Sander's Cap, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sazabi's Mental Sheath Basinet" Then Return "Sazabi's Basinet, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sigon's Visor Great Helm" Then Return "Sigon's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tal Rasha's Horadric Crest Death Mask" Then Return "Tals Mask, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tancred's Skull Bone Helm" Then Return "Tancred's Helm, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Trang-Ouls' Guise Bone Visage" Then Return "Trang-Ouls Visage, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Set Armor
        '***********************************************
        If ItemObjects(x).ItemBase = "Armor" Then
            If ItemObjects(x).ItemName = "Aldur's Deception Shadow Plate" Then Return "Aldur's, Armor Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat6
            If ItemObjects(x).ItemName = "Angelic Mantle Ring Mail" Then Return "Angelic Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Arcanna's Flesh Light Plate" Then Return "Arcanna's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Arctic Furs Quilted Armor" Then Return "Arctic Furs Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Berserker's Hauberk Splint Mail" Then Return "Beserker's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Cathan's Mesh Chain Mail" Then Return "Cathan's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Dark Adherent Dusk Shroud" Then Return "Dark Adherent Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Griswold's Heart Ornate Plate" Then Return "Griswold's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Haemosu's Adamant Cuirass" Then Return "Haemosu's Cuirass, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Hwanin's Refuge Tigulated Mail" Then Return "Hwanin's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Immortal King's Soul Cage Sacred Armor" Then Return "IK Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Isenhart's Case Breast Plate" Then Return "Isenhart's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "M'avina's Embrace Kraken Shell" Then Return "M'avina's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Milabrega's Robe Ancient Armor" Then Return "Milabrega's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Naj's Light Plate Hellforge Plate" Then Return "Naj's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Natalya's Shadow Loricated Mail" Then Return "Natalya's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sazabi's Ghost Liberator Balrog Skin" Then Return "Sazabi's Armor, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Sigon's Shelter Gothic Plate" Then Return "Sigon's Armor, Def " & ItemObjects(x).Defense & " " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Tal Rasha's Guardianship Lacquered Plate" Then Return "Tal's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tancred's Spine Full Plate Mail" Then Return "Tancred's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Trang-Oul's Scales Chaos Armor" Then Return "Trang-Oul's Armor, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Vidala's Ambush Leather Armor" Then Return "Vidala's, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Set Gloves
        '***********************************************
        If ItemObjects(x).ItemBase = "Shield" Or ItemObjects(x).ItemBase = "Voodoo Heads" Then
            If ItemObjects(x).ItemName = "Civerb's Ward Large Shield" Then Return "Civerb's Shield, Def " & ItemObjects(x).Defense '******
            If ItemObjects(x).ItemName = "Cleglaw's Claw Small Shield" Then Return "Cleglaw's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Griswold's Honor Vortex Shield" Then Return "Griswold's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Hsarus' Iron Fist Buckler" Then Return "Hsarus' Shield, Def " & ItemObjects(x).Defense '******
            If ItemObjects(x).ItemName = "Isenhart's Parry Gothic Shield" Then Return "Isenhart's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Milabrega's Orb Kite Shield" Then Return "Milabrega's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sigon's Guard Tower Shield" Then Return "Sigon's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Taebaek's Glory Ward" Then Return "Taebaek's Shield, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Trang-Oul's Wing Cantor Trophy" Then Return "Trang-Oul's Trophy, Def " & ItemObjects(x).Defense & " " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Whitstan's Guard Round Shield" Then Return "Whitstan's Shield, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Set Gloves
        '***********************************************
        If ItemObjects(x).ItemBase = "Gloves" Then
            If ItemObjects(x).ItemName = "Arctic Mitts Light Gauntlets" Then Return "Arctic Mitts, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Cleglaw's Pincers Chain Gloves" Then Return "Cleglaw's Gloves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Death's Hand Leather Gloves" Then Return "Death's Hand  Gloves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Immortal King's Forge War Gauntlets" Then Return "IK Gauntlets, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Iratha's Cuff Light Gauntlets" Then Return "Iratha's Gauntlets, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Laying of Hands Bramble Mitts" Then Return "LoH Mitts, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "M'avina's Icy Clutch Battle Gauntlets" Then Return "M'avina's Gauntlets, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Magnus' Skin Sharkskin Gloves" Then Return "Magnus Gloves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sander's Taboo Heavy Gloves" Then Return "Sander's Gloves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sigon's Gage Gauntlets" Then Return "Sigon's Gloves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Trang-Oul's Claws Heavy Bracers" Then Return "Trang-Oul's Bracers, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Set Belts
        '***********************************************
        If ItemObjects(x).ItemBase = "Belt" Then
            If ItemObjects(x).ItemName = "Arctic Binding Light Belt" Then Return "Arctic Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Credendum Mithril Coil" Then Return "Credendum Mithril Coil, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Death's Guard Sash" Then Return "Death's Guard Sash, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Hsarus' Iron Stay Belt" Then Return "Hsarus' Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Hwanin's Blessing Belt" Then Return "Hwanin's Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Immortal King's Detail War Belt" Then Return "IK Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Infernal Sign Heavy Belt" Then Return "Infernal Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Iratha's Cord Heavy Belt" Then Return "Iratha's Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "M'avina's Tenet Sharkskin Belt" Then Return "M'avina's Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sigon's Wrap Plated Belt" Then Return "Sigon's Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tal Rasha's Fine-Spun Cloth Mesh Belt" Then Return "Tal's Belt, Def" & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat4
            If ItemObjects(x).ItemName = "Trang-Oul's Girth Troll Belt" Then Return "Trang-Oul's Belt, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Wilhelm's Pride Battle Belt" Then Return "Wilhelm's Belt, Def " & ItemObjects(x).Defense
        End If

        '***********************************************
        ' Set Boots
        '***********************************************
        If ItemObjects(x).ItemBase = "Boots" Then
            If ItemObjects(x).ItemName = "Aldur's Advance Battle Boots" Then Return "Aldur's Boots, Def " & ItemObjects(x).Defense & ", " & ItemObjects(x).Stat5
            If ItemObjects(x).ItemName = "Hsarus' Iron Heel Chain Boots" Then Return "Hsarus' Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Immortal King's Pillar War Boots" Then Return "IK Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Natalya's Soul Mesh Boots" Then Return "Natalya's Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Rite of Passage Demonhide Boots" Then Return "Rite Of Passage Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sander's Riprap Heavy Boots" Then Return "Sander's Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Sigon's Sabot Greaves" Then Return "Sigon's Greaves, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Tancred's Hobnails Boots" Then Return "Tancred's Boots, Def " & ItemObjects(x).Defense
            If ItemObjects(x).ItemName = "Vidala's Fetlock Light Plated Boots" Then Return "Vidala's Boots, Def " & ItemObjects(x).Defense
        End If

        If ItemObjects(x).ItemBase = "Axe" Then
            If ItemObjects(x).ItemName = "Berserker's Hatchet Double Axe" Then Return "Berserker's Axe"
            If ItemObjects(x).ItemName = "Tancred's Crowbill Military Pick" Then Return "Tancred's Pick"
        End If

        If ItemObjects(x).ItemBase = "Bow" Or ItemObjects(x).ItemBase = "Amazon Bow" Then
            If ItemObjects(x).ItemName = "Arctic Horn Short War Bow" Then Return "Arctic Bow"
            If ItemObjects(x).ItemName = "M'avina's Caster Grand Matron Bow" Then Return "M'avina's Bow"
            If ItemObjects(x).ItemName = "Vidala's Barb Long Battle Bow" Then Return "Vidala's Bow"
        End If


        If ItemObjects(x).ItemBase = "Claw" Then
            If ItemObjects(x).ItemName = "Natalya's Mark Scissors Suwayyah" Then Return "Natalya's Mark"
        End If

        If ItemObjects(x).ItemBase = "Mace" Or ItemObjects(x).ItemBase = "Hammer" Then
            If ItemObjects(x).ItemName = "Aldur's Rhythm Jagged Star" Then Return "Aldur's Star"
            If ItemObjects(x).ItemName = "Dangoon's Teaching Reinforced Mace" Then Return "Dangoon's Mace"
            If ItemObjects(x).ItemName = "Immortal King's Stone Crusher Ogre Maul" Then Return "IK Maul, " & ItemObjects(x).Stat7
        End If

        If ItemObjects(x).ItemBase = "Orb" Then
            If ItemObjects(x).ItemName = "Tal Rasha's Lidless Eye Swirling Crystal" Then Return "Tal's Lidless"
        End If

        If ItemObjects(x).ItemBase = "polearm" Then
            If ItemObjects(x).ItemName = "Hwanin's Justice Bill" Then Return "Hwanin's Justice"
        End If

        If ItemObjects(x).ItemBase = "Staff" Then
            If ItemObjects(x).ItemName = "Arcanna's Deathwand War Staff" Then Return "Arcanna's Staff"
            If ItemObjects(x).ItemName = "Cathan's Rule Battle Staff" Then Return "Cathan's Staff"
            If ItemObjects(x).ItemName = "Naj's Puzzler Elder Staff" Then Return "Naj's Staff"
        End If

        If ItemObjects(x).ItemBase = "Sword" Then
            If ItemObjects(x).ItemName = "Angelic Sickle Sabre" Then Return "Angelic Sabre"
            If ItemObjects(x).ItemName = "Bul-Kathos' Sacred Charge Colossus Blade" Then Return "BK Blade"
            If ItemObjects(x).ItemName = "Bul-Kathos' Tribal Guardian Mythical Sword" Then Return "Bk Sword"
            If ItemObjects(x).ItemName = "Cleglaw's Tooth Long Sword" Then Return "Cleglaw's Sword"
            If ItemObjects(x).ItemName = "Death's Touch War Sword" Then Return "Death's Touch Sword"
            If ItemObjects(x).ItemName = "Isenhart's Lightbrand Broad Sword" Then Return "Isenhart's Sword"
            If ItemObjects(x).ItemName = "Sazabi's Cobalt Redeemer Cryptic Sword" Then Return "Sazabi's Sword"
        End If

        If ItemObjects(x).ItemBase = "Scepter" Then
            If ItemObjects(x).ItemName = "Civerb's Cudgel Grand Scepter" Then Return "Civerb's Scepter"
            If ItemObjects(x).ItemName = "Milabrega's Rod War Scepter" Then Return "Milabrega's Scepter"
        End If

        If ItemObjects(x).ItemName = "Infernal Torch Grim Wand" Then Return "Infernal Wand"
        If ItemObjects(x).ItemName = "Sander's Superstition Bone Wand" Then Return "Sander's Wand"


        Return ItemObjects(x).ItemName & Stats_items(x) & " [Not found]"

    End Function

    Public Sub DupesList()
        Dim arr() As String = Main.TradeListRICHTEXTBOX.Text.Split(Chr(10))
        Dim count(UBound(arr)) As Integer

        For i = 0 To UBound(arr) - 1 'find duplicates and delete
            count(i) = 1
            For x = (i + 1) To UBound(arr)
                If arr(x) = "" Then Continue For
                If arr(i) = arr(x) Then arr(x) = "" : count(i) += 1
            Next
        Next
        Main.TradeListRICHTEXTBOX.Clear() 'clear list
        For x = 0 To UBound(arr)
            If count(x) > 1 Then arr(x) = arr(x) & " (" & count(x) & ")"
        Next
        Array.Sort(arr) 'sort list alphabetically

        For x = 0 To UBound(arr) ' re - list
            If arr(x) <> "" Then
                Main.TradeListRICHTEXTBOX.AppendText(arr(x) & vbCrLf & vbCrLf)
            End If
        Next
    End Sub
End Module
