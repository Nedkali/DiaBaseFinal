﻿Imports System.IO
Module AutoLogger
    Public Sub ImportLogFiles(relog)

        'DEFAULT DATABASE LOADED CHECK
        'If AppSettings.CurrentDatabase <> AppSettings.DefaultDatabase Then
        'Main.RichTextBox1.AppendText("Aborting - Default database not loaded")
        'Return
        'End If
        Dim RealmPath = ""
        If AppSettings.EtalVersion = "NED" Then
            'Assign correct directory to each log search using the REalmPath Var
            RealmPath = "\scripts\Configs\USEast\AMS\" : Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf & "East") : GetLogs(RealmPath, relog)
            RealmPath = "\scripts\Configs\USWest\AMS\" : Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf & "West") : GetLogs(RealmPath, relog)
            RealmPath = "\scripts\Configs\Asia\AMS\" : Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf & "Asia") : GetLogs(RealmPath, relog)
            RealmPath = "\scripts\Configs\Europe\AMS\" : Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf & "Europe") : GetLogs(RealmPath, relog)
            Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf)
            Main.ImportLogRICHTEXTBOX.ScrollToCaret()
            Return
        End If
        If AppSettings.EtalVersion = "KOL" Then
            Main.ImportLogRICHTEXTBOX.AppendText(vbCrLf & "Kolbot ")
            RealmPath = "\d2bs\kolbot\\MuleInventory\"
            GetLogs(RealmPath, relog)
            Return
        End If
        RealmPath = "\scripts\AMS"
        GetLogs(RealmPath, relog)

    End Sub

    Sub GetLogs(RealmPath, Relog)

        'DataBasePath = Application.StartupPath + "\DataBase\"
        MuleDataPath = AppSettings.EtalPath + RealmPath + "\MuleLogs\"
        ArchiveFolder = Application.StartupPath + "\Archive\"
        If Relog = False Then MuleLogPath = AppSettings.EtalPath + RealmPath
        If Relog = True Then MuleLogPath = AppSettings.InstallPath + "\Archive\" : Main.AllItemsLISTBOX.Items.Clear()
        'MessageBox.Show(MuleLogPath)

        'Check Log folder for files to process
        GetLogFiles()
        If LogFilesList.Count = 0 Then
            If AppSettings.EtalVersion = "NED" Then Main.ImportLogRICHTEXTBOX.AppendText(" Realm Has No Logs Ready.") : Main.ImportLogRICHTEXTBOX.ScrollToCaret() : Return 'If There Are no Log Files - exit
            Main.ImportLogRICHTEXTBOX.AppendText(MuleLogPath) : Return
        End If

            'Backup the database
            If AppSettings.BackupBeforeImports = True Then
                CreateBackup(AppSettings.CurrentDatabase)
            End If

            Main.ImportLogRICHTEXTBOX.AppendText(" Realm, Logs To Import = " & LogFilesList.Count & vbCrLf)
            Pretotal = ItemObjects.Count
            ProcessLogFiles(Relog)
            WriteToFile(0, AppSettings.DefaultDatabase, False)  'saves entire item objects and overwrites file contents
            For ItemIndex = Pretotal To ItemObjects.Count - 1
                Main.AllItemsLISTBOX.Items.Add(ItemObjects(ItemIndex).ItemName)
            Next
            Main.ImportLogRICHTEXTBOX.AppendText(TimeOfDay & " - Import Complete. Total Items = " & (ItemObjects.Count - Pretotal) & vbCrLf)
        Main.ImportLogRICHTEXTBOX.ScrollToCaret()
    End Sub


    Sub GetLogFiles()

        Dim Tally As Integer = 0
        LogFilesList.Clear()
        Dim LogFiles As String() = (Directory.GetFiles(MuleLogPath, "*")).Select(Function(p) Path.GetFileName(p)).ToArray() ' gets file and crops path

        Do Until Tally = LogFiles.Count
            If LogFiles(Tally).IndexOf("delete me.txt") = -1 Then                   'ignore the delete me file (if its still there)
                If Replace(LogFiles(Tally), ".txt", "").IndexOf("_") > -1 Then
                    LogFilesList.Add(LogFiles(Tally))                               'ADD LOG FILE NAME TO LogFilesList()
                End If
            End If
            Tally = Tally + 1
        Loop
    End Sub

    Sub GetmuleaccountFiles()
        Dim Tally As Integer = 0
        PassFiles.Clear()
        If (My.Computer.FileSystem.DirectoryExists(MuleDataPath)) = True Then
            Dim AllFiles As String() = (Directory.GetFiles(MuleDataPath, "*")).Select(Function(p) Path.GetFileName(p)).ToArray() ' gets file and crops path
            For Each item In AllFiles
                If AllFiles(Tally).IndexOf("_muleaccount.txt") > -1 Then PassFiles.Add(AllFiles(Tally))
                Tally = Tally + 1
            Next
        End If

    End Sub

    Function GetMulePass(ByVal accname)
        GetmuleaccountFiles()
        Dim counter As Integer = 0
        Dim temp As String = ""
        Dim returnstring As String = "Unknown"
        Dim temp1 As Array
        'Main.RichTextBox1.AppendText("Looking up " & accname & " information" & vbCrLf)
        For Each item In PassFiles

            Dim ReadPassFiles = My.Computer.FileSystem.OpenTextFileReader(MuleDataPath & PassFiles(counter))
            While ReadPassFiles.EndOfStream = False
                temp = ReadPassFiles.ReadLine()
                If temp.IndexOf(accname) <> -1 Then
                    returnstring = PassFiles(counter).Replace("_muleaccount.txt", "")
                    temp1 = temp.Split("/")
                    returnstring = returnstring + "," + temp1(1)
                    ReadPassFiles.Close()
                    Return returnstring
                    Exit For
                End If
            End While
            counter = counter + 1
        Next
        Return (",")
    End Function

    Private Sub ProcessLogFiles(relog)
        Dim temp As String = ""
        Dim Tally As Integer = 0
        Dim mycounter As Integer = 0
        Dim found As Boolean = True
        Dim myarray As Array
        Dim itemsremoved = 0

        Dim thispickbot As String = ""
        Dim thislogmuleacc As String = ""
        Dim thislogmulename As String = ""
        Dim thislogpass As String = ""
        Dim ThislogRealm As String = ""
        Dim ThislogCore As Boolean
        Dim ThislogLadder As Boolean
        Dim Thislogexpansion As Boolean
        Dim ThislogTime As String = ""
        Dim ThislogDate As String = ""


        Do Until Tally = LogFilesList.Count
            If My.Computer.FileSystem.FileExists(MuleLogPath & LogFilesList(Tally)) = True Then 'Verify the log Exists
                Dim LogFile = My.Computer.FileSystem.OpenTextFileReader(MuleLogPath & LogFilesList(Tally))
                Try 'Should catch most issues ??
                    thispickbot = LogFile.ReadLine() 'these lines should exist for each log
                    thislogmuleacc = LogFile.ReadLine()
                    thislogmulename = LogFile.ReadLine()
                    thislogpass = LogFile.ReadLine()
                    ThislogRealm = LogFile.ReadLine()
                    ThislogCore = LogFile.ReadLine()
                    ThislogLadder = LogFile.ReadLine()
                    Thislogexpansion = LogFile.ReadLine()
                    ThislogTime = LogFile.ReadLine()
                    ThislogDate = LogFile.ReadLine()
                    LogFile.ReadLine()


                    '------------------------------------------------------------------------------------------------------------------------------------------------
                    ' May need to move this section closer to end of routine to catch other errors ? find end section :)
                    ' Note we dont want to remove duplicates on error - try to retain items instead
                    '------------------------------------------------------------------------------------------------------------------------------------------------
                Catch ex As Exception
                    LogFile.Close()
                    My.Computer.FileSystem.MoveFile(MuleLogPath & LogFilesList(Tally), ArchiveFolder + "\Faulty\" & LogFilesList(Tally), True)
                    ' Need to display this error somewhere
                    Main.ImportLogRICHTEXTBOX.AppendText("Error in file " & LogFilesList(Tally) & vbCrLf)
                    Tally = Tally + 1
                    Continue Do ' need to continue through next files
                End Try

                If AppSettings.RemoveMuleDupes = True Then 'Adding in remove duplicated mule option (Ned)
                    Dim range = ItemObjects.Count - 1
                    For mc = range To 0 Step -1
                        If Pretotal > 0 And ItemObjects(mc).MuleName = thislogmulename And ItemObjects(mc).ItemRealm = ThislogRealm Then ' added realm into this check in case same mule name exists on more than 1 realm
                            ' maybe recover password and pickitbot here? from previous item info before delete
                            If thispickbot = "Unknown" And ItemObjects(mc).PickitAccount <> "Unknown" Then thispickbot = ItemObjects(mc).PickitAccount
                            If thislogpass = "Unknown" And ItemObjects(mc).MulePass <> "Unknown" Then thislogpass = ItemObjects(mc).MulePass
                            ThislogDate = ItemObjects(mc).ImportDate ' retain original date - usefull for sorting ladder/nonladder
                            'Main.ImportLogRICHTEXTBOX.AppendText("Import Date retained" & vbCrLf) 'debug msg
                            If relog = False Then Main.AllItemsLISTBOX.Items.RemoveAt(mc) 'remove from displayed list
                            ItemObjects.RemoveAt(mc)
                            Pretotal = Pretotal - 1
                            itemsremoved = itemsremoved + 1
                        End If
                    Next
                End If

                'If unable to recover pickit bot and password from above - try to recover from etal log files
                If (thispickbot = "Unknown" Or thislogpass = "Unknown") Then 'lets try to find them
                        Dim result As String = GetMulePass(thislogmuleacc)
                        myarray = result.Split(",")
                        If (myarray(0) <> "") Then
                            thispickbot = myarray(0)
                        End If
                        If (myarray(1) <> "") Then
                            thislogpass = myarray(1)
                        End If
                    End If

                'Assign password if all else failed to get from above from app settings
                If thislogpass = "Unknown" And AppSettings.DefaultPassword.Length > 0 Then thislogpass = AppSettings.DefaultPassword

                '------------------------------------------------------------------------------------------------------------------------------------------------
                'end section
                '------------------------------------------------------------------------------------------------------------------------------------------------

                Do
                    Dim NewObject As New ItemDatabase
                    NewObject.PickitAccount = thispickbot
                    NewObject.MuleAccount = thislogmuleacc
                    NewObject.MuleName = thislogmulename
                    NewObject.MulePass = thislogpass
                    NewObject.ItemRealm = ThislogRealm
                    NewObject.HardCore = ThislogCore
                    NewObject.Ladder = ThislogLadder
                    NewObject.Expansion = Thislogexpansion
                    NewObject.ImportTime = ThislogTime
                    NewObject.ImportDate = ThislogDate
                    NewObject.LastLogDate = Date.Now()
                    For x = 0 To 4 '     just in case of extra blank lines
                        If LogFile.EndOfStream = True Then Exit Do
                        temp = LogFile.ReadLine()
                        If temp <> "" Then Exit For
                    Next

                    Dim ItemName = temp
                    NewObject.ItemName = temp 'these 5 lines should exist for each item
                    temp = LogFile.ReadLine()
                    NewObject.Itemlevel = temp.Replace("Item Level ", "")
                    temp = LogFile.ReadLine()
                    NewObject.ItemBase = temp.Replace("Item Base ", "")
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemQuality = myarray(2)
                    temp = LogFile.ReadLine() : myarray = Split(temp, " ")
                    NewObject.ItemImage = myarray(2)
                    If NewObject.ItemImage = 653 Then NewObject.ItemName = "Token of Absolution" : NewObject.Stat1 = "Right-click to reset Stat/Skill Points"
                    NewObject.RuneWord = LogFile.ReadLine()


                    While LogFile.EndOfStream = False   'attempt to read item added information and exit if end of stream/file 
                        temp = LogFile.ReadLine()


                        If temp = Nothing Then Exit While 'breaks while loop if we find blank line then move onto next item  

                        'below probably not needed - added to dll
                        'If LCase(ItemName).Contains(LCase(temp)) = True Then Continue While ' Hack for dll created files

                        myarray = Split(temp, ": ", 0)  ' gathers most basic stats

                        Select Case myarray(0) + ":"
                            Case "Throw Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.ThrowDamageMin = myarray(0)
                                NewObject.ThrowDamageMax = myarray(myarray.Length - 1)

                            Case "One-Hand Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.OneHandDamageMin = myarray(0)
                                NewObject.OneHandDamageMax = myarray(myarray.Length - 1)

                            Case "Two-Hand Damage:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "to ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.TwoHandDamageMin = myarray(0)
                                NewObject.TwoHandDamageMax = myarray(myarray.Length - 1)

                            Case "Quantity:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.QuantityMin = myarray(myarray.Length - 1)

                            Case "Durability:"
                                Dim temp1 = myarray(myarray.Length - 1)
                                temp1 = Replace(temp1, "of ", "")
                                myarray = Split(temp1, " ", 0)
                                NewObject.DurabilityMin = myarray(0)
                                NewObject.DurabilityMax = myarray(myarray.Length - 1)

                            Case "Defense:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.Defense = myarray(myarray.Length - 1)

                            Case "Chance to Block:"
                                myarray = Split(temp, ": ", 0)
                                temp = myarray(myarray.Length - 1)
                                temp = temp.Replace("%", "")
                                NewObject.ChanceToBlock = temp

                            Case "Required Strength:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredStrength = myarray(myarray.Length - 1)

                            Case "Required Level:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredLevel = myarray(myarray.Length - 1)

                            Case "Required Dexterity:"
                                myarray = Split(temp, ": ", 0)
                                NewObject.RequiredDexterity = myarray(myarray.Length - 1)

                            Case Else
                                found = False

                        End Select

                        ' if not basic item stat then we need to set it to stat1 stat2 etc
                        If found = False Then
                            If temp.IndexOf("Class - ") <> -1 Then
                                myarray = Split(temp, "Class - ", 0)
                                Dim temp1 = myarray(myarray.Length - 1)
                                NewObject.AttackSpeed = temp1
                                NewObject.AttackClass = NewObject.ItemBase
                            End If

                            If temp.IndexOf("Socketed (") <> -1 Then
                                Dim sock As Integer = temp.IndexOf("Socketed (")
                                NewObject.Sockets = Integer.Parse(temp.Substring(sock + 10, 1))
                            End If

                            If temp.IndexOf("Ethereal") <> -1 Then
                                NewObject.EtherealItem = True
                            End If

                            If temp = "(Paladin Only)" Then
                                NewObject.RequiredCharacter = "Paladin"
                                found = True
                            End If

                            If temp = "(Sorceress Only)" Then
                                NewObject.RequiredCharacter = "Sorceress"
                                found = True
                            End If

                            If temp = "(Necromancer Only)" Then
                                NewObject.RequiredCharacter = "Necromancer"
                                found = True
                            End If

                            If temp = "(Amazon Only)" Then
                                NewObject.RequiredCharacter = "Amazon"
                                found = True
                            End If

                            If temp = "(Assassin Only)" Then
                                NewObject.RequiredCharacter = "Assassin"
                                found = True
                            End If

                            If temp = "(Druid Only)" Then
                                NewObject.RequiredCharacter = "Druid"
                                found = True
                            End If
                            ' check for fix for item class
                            If temp.IndexOf("Class") > -1 Then
                                myarray = temp.Split(" ")
                                NewObject.AttackClass = myarray(0)
                                myarray = temp.Split("- ")
                                NewObject.AttackSpeed = LTrim(myarray(1))
                                Continue While
                            End If
                            If found = False And NewObject.Stat1 = "" Then NewObject.Stat1 = temp : found = True
                            If found = False And NewObject.Stat2 = "" Then NewObject.Stat2 = temp : found = True
                            If found = False And NewObject.Stat3 = "" Then NewObject.Stat3 = temp : found = True
                            If found = False And NewObject.Stat4 = "" Then NewObject.Stat4 = temp : found = True
                            If found = False And NewObject.Stat5 = "" Then NewObject.Stat5 = temp : found = True
                            If found = False And NewObject.Stat6 = "" Then NewObject.Stat6 = temp : found = True
                            If found = False And NewObject.Stat7 = "" Then NewObject.Stat7 = temp : found = True
                            If found = False And NewObject.Stat8 = "" Then NewObject.Stat8 = temp : found = True
                            If found = False And NewObject.Stat9 = "" Then NewObject.Stat9 = temp : found = True
                            If found = False And NewObject.Stat10 = "" Then NewObject.Stat10 = temp : found = True
                            If found = False And NewObject.Stat11 = "" Then NewObject.Stat11 = temp : found = True
                            If found = False And NewObject.Stat12 = "" Then NewObject.Stat12 = temp : found = True
                            If found = False And NewObject.Stat13 = "" Then NewObject.Stat13 = temp : found = True
                            If found = False And NewObject.Stat14 = "" Then NewObject.Stat14 = temp : found = True
                            If found = False And NewObject.Stat15 = "" Then NewObject.Stat15 = temp : found = True
                        End If
                    End While

                    ' fixes area - correcting item imports
                    If NewObject.ItemName.IndexOf("Large Charm") > -1 Then NewObject.ItemBase = "Large Charm" ' torches misread as medium ????? weird maybe valid ??? 
                    If NewObject.ItemName.IndexOf("Grand Charm") > -1 Then NewObject.ItemBase = "Grand Charm" ' this also as above 

                    If NewObject.ItemBase = "Rune" Then
                        temp = GetRunes(NewObject.ItemName)
                        myarray = Split(temp, ",")
                        NewObject.RequiredLevel = myarray(0)
                        NewObject.Stat1 = myarray(1)
                        NewObject.Stat2 = myarray(2)
                        NewObject.Stat3 = myarray(3)
                        NewObject.Stat4 = myarray(4)
                        NewObject.Stat5 = myarray(5)
                        NewObject.Stat6 = myarray(6)
                        NewObject.Stat7 = myarray(7)
                        NewObject.Stat8 = myarray(8)
                        NewObject.Stat9 = myarray(9)
                        NewObject.Stat10 = myarray(10)
                        NewObject.Stat11 = myarray(11)

                    End If
                    If NewObject.ItemBase = "Gem" Then
                        temp = GetGemsStats(NewObject.ItemName)
                        myarray = Split(temp, ",")
                        NewObject.RequiredLevel = myarray(0)
                        NewObject.Stat1 = myarray(1)
                        NewObject.Stat2 = myarray(2)
                        NewObject.Stat3 = myarray(3)
                        NewObject.Stat4 = myarray(4)
                    End If

                    If NewObject.ItemName.IndexOf("Token of Absolution") <> -1 Then
                        NewObject.ItemName = "Token of Absolution"
                        NewObject.Stat1 = "Right-click to reset Stat/Skill Points"
                        NewObject.Stat2 = ""
                    End If

                    ' Handles set to nonladder if date set before last b.net reset
                    Dim resetdate As Date = Date.ParseExact(AppSettings.ResetDate, "d/M/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        Dim tempdate As Date
                        tempdate = Date.ParseExact(NewObject.ImportDate, "d/M/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo)
                        If Date.Compare(tempdate, resetdate) < 0 Then NewObject.Ladder = False

                    ItemObjects.Add(NewObject)
                    'Dim a = ItemObjects.Count - 1
                    'Main.AllItemsLISTBOX.Items.Add(ItemObjects(a).ItemName)
                Loop Until LogFile.EndOfStream
                    LogFile.Close()
                    If relog = False Then My.Computer.FileSystem.MoveFile(MuleLogPath & LogFilesList(Tally), ArchiveFolder & LogFilesList(Tally), True)

            End If
            Tally = Tally + 1
        Loop

        Main.ImportLogRICHTEXTBOX.AppendText("Duplicated Items removed = " & itemsremoved & vbCrLf)
    End Sub

    Function GetRunes(ByVal runename)
        Dim runestats As String = ""
        Select Case runename

            Case "Eld Rune"
                runestats = "11, Weapons:, +50 To Attack Rate, +175% Damage To Undead, Armour:, +15% Stamina Drain, Helmets:, +15% Stamina Drain, Shields:, Increased Chance Of Blocking,,,,"

            Case "El Rune"
                runestats = "11, Weapons:, +1 To Light Radius, 50 To Attack Rating,  Armour:, +50 To Attack Rate Vs Undead, +175% Damage Vs Undead, Helmets:, +1 To Light Radius, +15 To Defense, Shields:, +1 to Light Radius, +15 To Defense"

            Case "Tir Rune"
                runestats = "13, Weapons:, +2 Mana After Kill, Armour:, +2 Mana After Kill, Helmets:, +2 Mana After Kill, Shields:, +2 Mana After Kill,,,,"

            Case "Nef Rune"
                runestats = "13, Weapons:, Knockback, Armour:, +30 Defense Vs Missiles, Helmets:, +30 Defense Vs Missiles, Shields:, +30 Defense Vs Missiles,,,"

            Case "Eth Rune"
                runestats = "15, Weapons:, -25% To Targets Defense, Armour:, Regenerate Mana 15%, Helmets:, Regenerate Mana 15%, Shields:, Regenerate Mana 15%,,,"

            Case "Ith Rune"
                runestats = "15, Weapons:, -25% To Targets Defense, Armour:, Regenerate Mana 15%, Helmets:, Regenerate Mana 15%, Shields:, Regenerate Mana 15%,,,"

            Case "Tal Rune"
                runestats = "17, Weapons:, +19 Poison Damage for 2 seconds, Armour:, Poison Resistance +30%, Helmets:, Poison Resistance +30%, Shields:, Regenerate Mana 15%,,,,,"

            Case "Ral Rune"
                runestats = "19, Weapons:, Adds 5 To 30 Fire Damage, Armour:, Fire Resistance +30%, Helmets:, Fire Resistance +30%, Shields:, Fire Resistance +30%,,,,,,,"

            Case "Ort Rune"
                runestats = "19, Weapons:, Adds 1 To 50 Light Damage, Armour:, Lightning Resistance +30%, Helmets:, Lightning Resistance +30%, Shields:, Lightning Resistance +30%,,,,,"

            Case "Thul Rune"
                runestats = "21, Weapons:, Adds 3 To 14 Cold Damage, Armour:, Cold Resistance +30%, Helmets:, Cold Resistance +30%, Shields:, Cold Resistance +30%,,,,"

            Case "Amn Rune"
                runestats = "23, Weapons:, +7% Life Stolen Per Hit, Armour:, Attacker Takes Damage Of 14, Helmets:, Attacker Takes Damage Of 14, Shields:, Attacker Takes Damage Of 14,,,"

            Case "Sol Rune"
                runestats = "25, Weapons:, +9 To Minimum Damage, Armour:, Damage Reduced By 7, Helmets:, Damage Reduced By 7, Shields:, Damage Reduced By 7,,,"

            Case "Shael Rune"
                runestats = "27, Weapons:, Increased Attack Speed, Armour:, Faster Hit Recovery, Helmets:, Faster Hit Recovery, Shields:, Faster Block Rate,,,"

            Case "Shae Rune"
                runestats = "29, Weapons:, Increased Attack Speed, Armour:, Faster Hit Recovery, Helmets:, Faster Hit Recovery, Shields:, Faster Block Rate,,,"

            Case "Dol Rune"
                runestats = "31, Weapons:, +32% Monster Flees, Armour:, +7 Replenish Life, Helmets:, +7 Replenish Life, Shields:, +7 Replenish Life,,,"

            Case "Hel Rune"
                runestats = "1, Weapons:, -20 Requirements, Armour:, -15 Requirements, Helmets:, -15 Requirements, Shields:, -15 Requirements,,,"

            Case "Po Rune"
                runestats = "35, Weapons:, +10 To Vitality, Armour:, +10 To Vitality, Helmets:, +10 To Vitality, Shields:, +10 To Vitality,,,"

            Case "Io Rune"
                runestats = "35, Weapons:, +10 To Vitality, Armour:, +10 To Vitality, Helmets:, +10 To Vitality, Shields:, +10 To Vitality,,,"

            Case "Lum Rune"
                runestats = "37, +10 To Energy, Armour:, +10 To Energy, Helmets:, +10 To Energy, Shields:, +10 To Energy,,,,"

            Case "Ko Rune"
                runestats = "39, Weapons:, +10 To Dexterity, Armour:, +10 To Dexterity, Helmets:, +10 To Dexterity, Shields:, +10 To Dexterity,,,"

            Case "Fal Rune"
                runestats = "41, Weapons:, +10 To Strength, Armour:, +10 To Strength, Helmets:, +10 To Strength, Shields:, +10 To Strength,,,"

            Case "Lem Rune"
                runestats = "43, Weapons:, +75% Extra Gold From Slain Monsters, Armour:, +50% Extra Gold From Slain Monsters, Helmets:, +50% Extra Gold From Slain Monsters, Shields:, +50% Extra Gold From Slain Monsters,,,"

            Case "Pul Rune"
                runestats = "45, Weapons:, +100 To Attack Rating Against Demons, +175% Damage Against Demons , Armour:, +30% Enhanced Defense, Helmets:, +30% Enhanced Defense, Shields:, +30% Enhanced Defense,,"

            Case "Um Rune"
                runestats = "47, Weapons:, +25% Chance Of Open Wounds, Armour:, +20 To All Resistances, Helmets:, +10 To All Resistances, Shields:, +20 To All Resistances,,,"

            Case "Mal Rune"
                runestats = "49, Weapons:, Prevent Monster Heal, Armour:, -7 To Magic Damage Received, Helmets:, -7 To Magic Damage Received, Shields:, -7 To Magic Damage Received,,,"

            Case "Ist Rune"
                runestats = "51, Weapons:, +30% To Magic Items, Armour:, +25% To Magic Find, Helmets:, +25% To Magic Items, Shields:, +25% To Magic Items,,,"

            Case "Gul Rune"
                runestats = "53, Weapons:, +20% To Attack Rating, Armour:, +3% To Max Poison Resist, Helmets:, +3% To Max Poison Resist, Shields:, +3% To Max Poison Resist,,,"

            Case "Vex Rune"
                runestats = "55, Weapons:, +7% Mana Stolen Per Hit, Armour:, +3% To Max Fire Resist, Helmets:, +3% To Max Fire Resist, Shields:, +3% To Max Fire Resist,,,"

            Case "Ohm Rune"
                runestats = "57, Weapons:, +50% Enhanced Damege, Armour:, +3% To Max Cold Resist, Helmets:, +3% To Max Cold Resist, Shields:, +3% To Max Cold Resist,,,"

            Case "Lo Rune"
                runestats = "59, Weapons:, +20% Chance Of Deadly Strike, Armour:, +3% To Max Lightning Resist, Helmets:, +3% To Max Lightning Resist, Shields:, +3% To Max Lightning Resist,,,"

            Case "Sur Rune"
                runestats = "61, Weapons:, Hit Blinds Target, Armour:, +5% To Max Mana, Helmets:, +5% To Max Mana, Shields:, +50 To Max Mana,,,"

            Case "Jah Rune"
                runestats = "65, Weapons:, Ignore Target's Defense, Armour:, Increase Maximum Life 5%, Helmets:, Increase Maximum Life 5%, Shields:, +50 To Life,,,"
            Case "Ber Rune"
                runestats = "63, Weapons:, +20% Chance Of Crushing Blow, Armour:, -8% To Damage Received, Helmets:, -8% To Damage Received, Shields:, -8% To Damage Received,,,"

            Case "Jo Rune"
                runestats = "65, Weapons:, Slows Target By 25%, Armour:, +5% To Max Life, Helmets:, +5% To Max Life, Shields:, +50 To Life,,,"

            Case "Cho Rune"
                runestats = "65, Weapons:, Slows Target By 25%, Armour:, +5% To Max Life, Helmets:, +5% To Max Life, Shields:, +50 To Life,,,"

            Case "Cham Rune"
                runestats = "67, Weapons:, Freeze Target, Armour:, Cannot Be Frozen, Helmets:, Cannot Be Frozen, Shields:, Cannot Be Frozen,,,"

            Case "Zod Rune"
                runestats = "69, Weapons:, Indestructable, Armour:, Indestructable, Helmets:, Indestructable, Shields:, Indestructable,,,"

        End Select
        Return runestats
    End Function
    Function GetGemsStats(ByVal gemname)

        Dim gemstats As String = ""
        Select Case gemname

            Case "Chipped Diamond"
                gemstats = "1,Weapons: +28% Damage to Undead, Armor: +20 to Attack Rating, Helms: +20 to Attack Rating, Shields: All Resistances 6"

            Case "Chipped Ruby"
                gemstats = "1,Weapons: Adds 3-4 Fire Damage, Armor: +10 to Life, Helms: +10 to Life, Shields: Fire Resist +12%"

            Case "Chipped Sapphire"
                gemstats = "1,Weapons: Adds 1-3 Cold Damage, Armor: +10 to Mana, Helms: +10 to Mana, Shields: Cold Resist +12%"

            Case "Chipped Emerald"
                gemstats = "1,Weapons: +10 Poison Damage over 3 Seconds, Armor: +3 to Dexterity, Helms: +3 to Dexterity, Shields: Poison Resist +12%"

            Case "Chipped Topaz"
                gemstats = "1,Weapons: Adds 1-8 Lightning Damage, Armor: 9% BetterChance of Finding Magic Items, Helms: 9% BetterChance of Finding Magic Items, Shields: Lightning Resist +12%"

            Case "Chipped Amethyst"
                gemstats = "1,Weapons: +40 To Attack Rating, Armor: +3 To Strength, Helms: +3 To Strength, Shields: +8 Defense"

            Case "Chipped Skull"
                gemstats = "1,Weapons: 2% Life Stolen Per Hit, 1% Mana Stolen Per Hit, Armor: Regenerate Mana 8%, Replenish life +2, Helms: Regenerate Mana 8%, Replenish life +2, Shields: Attacker Takes Damage of 4"

            Case "Flawed Diamond"
                gemstats = "5,Weapons: +34% Damage to Undead, Armor: +40 to Attack Rating, Helms: +40 to Attack Rating, Shields: All Resistances 8"

            Case "Flawed Ruby"
                gemstats = "5,Weapons: Adds 5-8 Fire Damage, Armor: +17 to Life, Helms: +17 to Life, Shields: Fire Resist +16%"

            Case "Flawed Sapphire"
                gemstats = "5,Weapons: Adds 3-5 Cold Damage, Armor: +17 to Mana, Helms: +17 to Mana, Shields: Cold Resist +16%"

            Case "Flawed Emerald"
                gemstats = "5,Weapons: +20 Poison Damage over 4 Seconds, Armor: +4 to Dexterity, Helms: +4 to Dexterity, Shields: Poison Resist +16%"

            Case "Flawed Topaz"
                gemstats = "5,Weapons: Adds 1-14 Lightning Damage, Armor: 13% BetterChance of Finding Magic Items, Helms: 13% BetterChance of Finding Magic Items, Shields: Lightning Resist +16%"

            Case "Flawed Amethyst"
                gemstats = "5,Weapons: +60 To Attack Rating, Armor: +4 To Strength, Helms: +4 To Strength, Shields: +12 Defense"

            Case "Flawed Skull"
                gemstats = "5,Weapons: 2% Life Stolen Per Hit, 2% Mana Stolen Per Hit, Armor: Regenerate Mana 8%, Replenish life +3, Helms: Regenerate Mana 8%, Replenish life +3, Shields: Attacker Takes Damage of 8"

            Case "Diamond"
                gemstats = "12,Weapons: +44% Damage to Undead, Armor: +60 to Attack Rating, Helms: +60 to Attack Rating, Shields: All Resistances 11"

            Case "Ruby"
                gemstats = "12,Weapons: Adds 8-12 Fire Damage, Armor: +24 to Life, Helms: +24 to Life, Shields: Fire Resist +22%"

            Case "Sapphire"
                gemstats = "12,Weapons: Adds 4-7 Cold Damage, Armor: +24 to Mana, Helms: +24 to Mana, Shields: Cold Resist +22%"

            Case "Emerald"
                gemstats = "12,Weapons: +40 Poison Damage over 5 Seconds, Armor: +6 to Dexterity, Helms: +6 to Dexterity, Shields: Poison Resist +22%"

            Case "Topaz"
                gemstats = "12,Weapons: Adds 1-22 Lightning Damage, Armor: 16% BetterChance of Finding Magic Items, Helms: 16% BetterChance of Finding Magic Items, Shields: Lightning Resist +22%"

            Case "Amethyst"
                gemstats = "12,Weapons: +80 To Attack Rating, Armor: +6 To Strength, Helms: +6 To Strength, Shields: +18 Defense"

            Case "Skull"
                gemstats = "12,Weapons: 3% Life Stolen Per Hit, 2% Mana Stolen Per Hit, Armor: Regenerate Mana 12%, Replenish life +3, Helms: Regenerate Mana 12%, Replenish life +3, Shields: Attacker Takes Damage of 12"

            Case "Flawless Diamond"
                gemstats = "15,Weapons: +54% Damage to Undead, Armor: +80 to Attack Rating, Helms: +80 to Attack Rating, Shields: All Resistances 14"

            Case "Flawless Ruby"
                gemstats = "15,Weapons: Adds 10-16 Fire Damage, Armor: +31 to Life, Helms: +31 to Life, Shields: Fire Resist +28%"

            Case "Flawless Sapphire"
                gemstats = "15,Weapons: Adds 6-10 Cold Damage, Armor: +31 to Mana, Helms: +31 to Mana, Shields: Cold Resist +28%"

            Case "Flawless Emerald"
                gemstats = "15,Weapons: +60 Poison Damage over 6 Seconds, Armor: +8 to Dexterity, Helms: +8 to Dexterity, Shields: Poison Resist +28%"

            Case "Flawless Topaz"
                gemstats = "15,Weapons: Adds 1-30 Lightning Damage, Armor: 20% BetterChance of Finding Magic Items, Helms: 20% BetterChance of Finding Magic Items, Shields: Lightning Resist +28%"

            Case "Flawless Amethyst"
                gemstats = "15,Weapons: +100 To Attack Rating, Armor: +8 To Strength, Helms: +8 To Strength, Shields: +24 Defense"

            Case "Flawless Skull"
                gemstats = "15,Weapons: 3% Life Stolen Per Hit, 3% Mana Stolen Per Hit, Armor: Regenerate Mana 12%, Replenish life +4, Helms: Regenerate Mana 12%, Replenish life +4, Shields: Attacker Takes Damage of 16"

            Case "Perfect Diamond"
                gemstats = "18,Weapons: +68% Damage to Undead, Armor: +100 to Attack Rating, Helms: +100 to Attack Rating, Shields: All Resistances 19"

            Case "Perfect Ruby"
                gemstats = "18,Weapons: Adds 15-20 Fire Damage, Armor: +38 to Life, Helms: +38 to Life, Shields: Fire Resist +40%"

            Case "Perfect Sapphire"
                gemstats = "18,Weapons: Adds 10-14 Cold Damage, Armor: +38 to Mana, Helms: +38 to Mana, Shields: Cold Resist +40%"

            Case "Perfect Emerald"
                gemstats = "18,Weapons: +100 Poison Damage over 7 Seconds, Armor: +10 to Dexterity, Helms: +10 to Dexterity, Shields: Poison Resist +40%"

            Case "Perfect Topaz"
                gemstats = "18,Weapons: Adds 1-40 Lightning Damage, Armor: 24% BetterChance of Finding Magic Items, Helms: 24% BetterChance of Finding Magic Items, Shields: Lightning Resist +40%"

            Case "Perfect Amethyst"
                gemstats = "18,Weapons: +150 To Attack Rating, Armor: +10 To Strength, Helms: +10 To Strength, Shields: +30 Defense"

            Case "Perfect Skull"
                gemstats = "18,Weapons: 4% Life Stolen Per Hit, 3% Mana Stolen Per Hit, Armor: Regenerate Mana 19%, Replenish life +5, Helms: Regenerate Mana 19%, Replenish life +5, Shields: Attacker Takes Damage of 20"

        End Select
        If gemstats = "" Then gemstats = "18, Weapons: N/A, Armor: N/A, Helms: N/A, Shields: N/A"
        Return gemstats
    End Function

End Module
