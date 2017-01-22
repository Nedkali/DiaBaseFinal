Public Class RuneWord

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim RWRunes As String = ""
        Dim RWBase As String = ""
        Select Case ItemTypeCBOX.Text
            Case "Weapon"
                Select Case ComboBox1.Text
                    Case "Beast"
                        RWRunes = "Ber, Tir, Um, Mal, Lum"
                        RWBase = "Axe"
                    Case "Brand"
                        RWRunes = "Jah, Lo, Mal, Gul"
                        RWBase = "CrusaderBow, Grand Matron Bow, Hydra Bow"
                    Case "Breath of The Dying"
                        RWRunes = "Vex, Hel, El, Eld, Zod, Eth"
                        RWBase = "Berserker Axe, Colossus Blade, Crusader Bow, Ghost Spear, Great Poleaxe, Hydra Bow, Ogre Maul, War Pike"
                    Case "Call To Arms"
                        RWRunes = "Amn, Ral, Mal, Ist, Ohm"
                        RWBase = "Colussus Blade, Crystal Sword, Flail, Grand Matron Bow, Matriachal Bow, Phase Blade, War Scepter, War Staff"
                    Case "Chaos"
                        RWRunes = "Fal, Ohm, Um"
                        RWBase = "Runic Talons, Suwayyah"
                    Case "Cresent Moon"
                        RWRunes = "Shael, Um, Tir"
                        RWBase = "Balrog Blade, Berseker Axe, Legend Sword, Phase Blade"
                    Case "Death"
                        RWRunes = "Hel, El, Vex, Ort, Gul"
                        RWBase = "Berserker Axe"
                    Case "Destruction"
                        RWRunes = "Vex, Lo, Ber, Jah, Ko"
                        RWBase = "Polearms, Swords"
                    Case "Doom"
                        RWRunes = "Hel, Ohm, Um, Lo, Cham"
                        RWBase = "Axe, Polearm, Hammers"
                    Case "Edge"
                        RWRunes = "Tir, Tal, Amn"
                        RWBase = "Missile Weapon"
                    Case "Faith"
                        RWRunes = "Ohm, Jah, Lem, Eld"
                        RWBase = ""
                    Case "Famine"
                        RWRunes = "Fal, Um, Pul"
                        RWBase = "Axe, Hammer"
                    Case "Fortitude"
                        RWRunes = "El, Sol, Dol, Lo"
                        RWBase = "Armor"
                    Case "Fury"
                        RWRunes = "Jah, Gul, Eth"
                        RWBase = "Balrog Blade, Runic Talons, Suwayyah"
                    Case "Grief"
                        RWRunes = "Eth, Tir, Lo, Mal, Ral"
                        RWBase = "Berserker Axe, Phase Blade, "
                    Case "Hand of Justice"
                        RWRunes = "Sur, Cham, Amn, Lo"
                        RWBase = "Berserker Axe, Crusader Bow,Grand Matron Bow, Legendary Mallet, Phase Blade"
                    Case "Harmony"
                        RWRunes = "Tir, Ith, Sol, Ko"
                        RWBase = "Crusader Bow, Hydra Bow"
                    Case "Heart of the Oak"
                        RWRunes = "Ko, Vex, Pul, Thul"
                        RWBase = "Flail, Scourge"
                    Case "Ice"
                        RWRunes = "Amn, Shael, Jah, Lo"
                        RWBase = "Crusader Bow, Hydra Bow"
                    Case "Infinity"
                        RWRunes = "Ber, Mal, Ber, Ist"
                        RWBase = "Polearm"
                    Case "Insight"
                        RWRunes = "Ral, Tir, Tal, Sol"
                        RWBase = "Polearm"
                    Case "Kingslayer"
                        RWRunes = "Mal, Um, Gul, Fal"
                        RWBase = "Berserker Axe, Highland Blade, Phase Blade"
                    Case "Last Wish"
                        RWRunes = "Jah, Mal, Jah, Sur, Jah, Ber"
                        RWBase = "Berserker Axe, Phase Blade"
                    Case "Lawbringer"
                        RWRunes = "Amn, Lem, Ko"
                        RWBase = "Crystal Sword, Executioner Sword, Phase Blade"
                    Case "Memory"
                        RWRunes = "Lum, Io, Sol, Eth"
                        RWBase = "Stave"
                    Case "Oath"
                        RWRunes = "Shael, Pul, Mal, Lum"
                        RWBase = "Swords, Axe, Mace"
                    Case "Obedience"
                        RWRunes = "Hel, Ko, Thul, Eth, Fal"
                        RWBase = "Polearm"
                    Case "Passion"
                        RWRunes = "Dol, Ort, Eld, Lem"
                        RWBase = "Berserker Axe, Champion Sword, Grand Matron Bow,Phase Blade,"
                    Case "Phoenix"
                        RWRunes = "Vex, Vex, Lo, Jah"
                        RWBase = "Berserker Axe, Champion Sword,Legendary Mallet, Phase Blade, Scourge"
                    Case "Pride"
                        RWRunes = "Cham, Sur, Io, Lo"
                        RWBase = "Polearm"
                    Case "Rift"
                        RWRunes = "Hel, Ko, Lem, Gul"
                        RWBase = "Polearm"
                    Case "Silence"
                        RWRunes = "Dol, Eld, Hel, Ist, Tir, Vex"
                        RWBase = "Berserker Axe, Colossus Blade, Crystal Sword, Phase Blade"
                    Case "Spirit"
                        RWRunes = "Tal, Thul, Ort, Amn"
                        RWBase = "Crystal Sword"
                    Case "Voice of Reason"
                        RWRunes = "Lem, Ko, El, Eld"
                        RWBase = "Dimensional Blade, Highland Blade,Phase Blade, Scourge, Zweihander"
                    Case "White"
                        RWRunes = "Dol, Io"
                        RWBase = "Wand"
                    Case "Wrath"
                        RWRunes = "Pul, Lum, Ber, Mal"
                        RWBase = "Crusader Bow"
                    Case "Wind"
                        RWRunes = "Sur, El"
                        RWBase = "Melee Weapons"
                    Case Else
                        Return
                End Select
            Case "Armor"
                Select Case ComboBox1.Text
                    Case "Bone"
                        RWRunes = "Sol, Um, Um"
                        RWBase = ""
                    Case "Bramble"
                        RWRunes = "Ral, Ohm, Sur, Eth "
                        RWBase = ""
                    Case "Chains of Honor"
                        RWRunes = "Dol, Um, Ber, Ist"
                        RWBase = ""
                    Case "Dragon"
                        RWRunes = "Sur, Lo, Sol"
                        RWBase = ""
                    Case "Duress"
                        RWRunes = "Shael, Um, Thul"
                        RWBase = ""
                    Case "Enigma"
                        RWRunes = "Jah, Ith, Ber"
                        RWBase = ""
                    Case "Enlightenment"
                        RWRunes = "Pul, Ral, Sol"
                        RWBase = ""
                    Case "Fortitude"
                        RWRunes = "El, Sol, Dol, Lo"
                        RWBase = ""
                    Case "Myth"
                        RWRunes = "Hel, Amn, Nef"
                        RWBase = ""
                    Case "Peace"
                        RWRunes = "Shael, Thul, Amn"
                        RWBase = ""
                    Case "Principle"
                        RWRunes = "Ral, Gul, Eld"
                        RWBase = ""
                    Case "Prudence"
                        RWRunes = "Mal, Tir"
                        RWBase = ""
                    Case "Rain"
                        RWRunes = "Ort, Mal, Ith"
                        RWBase = ""
                    Case "Stone"
                        RWRunes = "Shael, Um, Pul, Um"
                        RWBase = ""
                    Case "Treachery"
                        RWRunes = "Shael, Thul, Lem"
                        RWBase = ""
                    Case "Wealth"
                        RWRunes = "Lem, Ko, Tir"
                        RWBase = ""
                    Case Else
                        Return
                End Select

            Case "Shield"
                Select Case ComboBox1.Text
                    Case "Dragon"
                        RWRunes = "Sur, Lo, Sol"
                        RWBase = ""
                    Case "Dream"
                        RWRunes = "Io, Jah, Pul"
                        RWBase = ""
                    Case "Exile"
                        RWRunes = "Vex, Ohm, Ist, Dol"
                        RWBase = ""
                    Case "Phoenix"
                        RWRunes = "Vex, Vex, Pul, Thul"
                        RWBase = ""
                    Case "Rhyme"
                        RWRunes = "Shael, Eth"
                        RWBase = ""
                    Case "Sanctuary"
                        RWRunes = "Ko, Ko, Mal"
                        RWBase = ""
                    Case "Spirit"
                        RWRunes = "Tal, Thul, Ort, Amn"
                        RWBase = ""
                    Case "Splendor"
                        RWRunes = "Eth, Lum"
                        RWBase = ""
                    Case Else
                        Return

                End Select
            Case Else
                Return
        End Select

        Label3.Text = "Rune Order = " & RWRunes
        Label3.Visible = True
        RWSearch(RWBase, RWRunes)

    End Sub


    Private Sub RuneWord_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ItemTypeCBOX.SelectedIndex = 0
        Label3.Visible = False
        If AppSettings.DefaultRealm = "USEast" Then EastRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "USWest" Then WestRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "Asia" Then AsiaRealmCHECKBOX.Checked = True
        If AppSettings.DefaultRealm = "Europe" Then EuropeRealmCHECKBOX.Checked = True
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ItemTypeCBOX.SelectedIndexChanged
        If ItemTypeCBOX.Text = "Weapon" Then
            ComboBox1.Items.Clear()
            Dim installs() As String = New String() {"Beast", "Brand", "Breath of The Dying", "Call To Arms", "Chaos",
                "Cresent Moon", "Death", "Destruction", "Doom", "Edge", "Faith", "Famine", "Fortitude", "Fury", "Grief",
                "Hand of Justice", "Harmony", "Heart of the Oak", "Ice", "Infinity", "Insight", "Kingslayer", "Last Wish",
                "Lawbringer", "Memory", "Oath", "Obedience", "Passion", "Phoenix", "Pride", "Rift", "Silence", "Spirit",
                "Voice of Reason", "White", "Wrath"}
            ComboBox1.Items.AddRange(installs)
            ComboBox1.SelectedIndex = 0
            Return
        End If
        If ItemTypeCBOX.Text = "Armor" Then
            ComboBox1.Items.Clear()
            Dim installs() As String = New String() {"Bone", "Bramble", "Chains of Honor", "Dragon", "Duress", "Enigma",
                "Enlightenment", "Fortitude", "Myth", "Peace", "Principle", "Prudence", "Rain", "Stone", "Treachery", "Wealth"}
            ComboBox1.Items.AddRange(installs)
            ComboBox1.SelectedIndex = 0
            Return
        End If
        If ItemTypeCBOX.Text = "Shield" Then
            ComboBox1.Items.Clear()
            Dim installs() As String = New String() {"Dragon", "Dream", "Exile", "Phoenix", "Rhyme", "Sanctuary", "Spirit", "Splendor"}
            ComboBox1.Items.AddRange(installs)
            ComboBox1.SelectedIndex = 0
        End If

    End Sub

    Private Sub EastRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles EastRealmCHECKBOX.CheckedChanged
        If EastRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub WestRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles WestRealmCHECKBOX.CheckedChanged
        If WestRealmCHECKBOX.Checked = True Then
            EastRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub AsiaRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles AsiaRealmCHECKBOX.CheckedChanged
        If AsiaRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            EastRealmCHECKBOX.Checked = False
            EuropeRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub EuropeRealmCHECKBOX_CheckedChanged(sender As Object, e As EventArgs) Handles EuropeRealmCHECKBOX.CheckedChanged
        If EuropeRealmCHECKBOX.Checked = True Then
            WestRealmCHECKBOX.Checked = False
            EastRealmCHECKBOX.Checked = False
            AsiaRealmCHECKBOX.Checked = False
        End If
    End Sub

    Private Sub RWSearch(ByVal base As String, ByVal runes As String)
        If ItemObjects.Count = 0 Then Return
        DataGridView1.Rows.Clear()

        Dim sRealm As String = ""
        If WestRealmCHECKBOX.Checked = True Then sRealm = "USWest"
        If EastRealmCHECKBOX.Checked = True Then sRealm = "USEast"
        If AsiaRealmCHECKBOX.Checked = True Then sRealm = "Asia"
        If EuropeRealmCHECKBOX.Checked = True Then sRealm = "Europe"

        Dim itemstr As String = ""
        Dim temp = base.Split(",")

        Dim runestr As String = ""
        Dim temp2 = runes.Split(",")
        Dim soc = temp2.Count

        For index = 0 To temp.Count - 1
            itemstr = itemstr & temp(index)
        Next


        For index = 0 To temp2.Count - 1
            runestr = runestr & temp2(index) & " Rune "
        Next

        For index = 0 To ItemObjects.Count - 1
            If ItemObjects(index).ItemRealm <> sRealm Then Continue For

            If LadderCBox.Checked = True And ItemObjects(index).Ladder = False Then Continue For
            If LadderCBox.Checked = False And ItemObjects(index).Ladder = True Then Continue For

            If CoreCBox.Checked = True And ItemObjects(index).HardCore = False Then Continue For
            If CoreCBox.Checked = False And ItemObjects(index).HardCore = True Then Continue For

            'If itemstr.Contains(ItemObjects(index).ItemBase) = True Then

            ' End If

            If runestr.Contains(ItemObjects(index).ItemName) = True Then
                DataGridView1.Rows.Add(ItemObjects(index).ItemName, ItemObjects(index).MuleAccount, ItemObjects(index).MuleName)
            End If
        Next

        DataGridView1.Sort(DataGridView1.Columns(2), System.ComponentModel.ListSortDirection.Ascending)

    End Sub
End Class