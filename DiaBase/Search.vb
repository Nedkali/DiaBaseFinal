Module Search
    Sub SearchRoutine()

        RefineSearchReferenceList.Clear()                       'Clear out old refine list
        If SearchReferenceList.Count > 0 Then                   'Can only run a refined search if items exist already in the search lisbox
            For Each item In SearchReferenceList
                RefineSearchReferenceList.Add(item)             'Create a reference list of each items object location in database that have 
                '                                               'Already been matched and are right now still in the matched search list....
                '                                               'Refine searches will use this list to look for matches as opposed all items list
            Next
        End If

        SearchReferenceList.Clear()                             'Clear Out old Item Search Reference List (Holds location in database of each matched item)
        Main.SearchLISTBOX.Items.Clear()                        'Clear out old search matches
        Dim searchtype = Main.SearchFieldCOMBOBOX.Text          'added this so it reads better
        Dim a As Integer

        Select Case searchtype
            '--------------------------------------------------------------------------
            'text based searches atributes handled elsewhere
            '--------------------------------------------------------------------------
            Case "Item Name", "Item Quality", "Item Base", "Mule Name", "Mule Account", "Mule Pass", "Realm", "User Reference", "Attack Class", "Attack Speed"
                If Main.RefineSearchCHECKBOX.Checked = False Then
                    If searchtype = "Item Name" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).ItemName) : Next
                    If searchtype = "Item Base" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).ItemBase) : Next
                    If searchtype = "Item Quality" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).ItemQuality) : Next
                    If searchtype = "Mule Name" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).MuleName) : Next
                    If searchtype = "Mule Account" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).MuleAccount) : Next
                    If searchtype = "Realm" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).ItemRealm) : Next
                    If searchtype = "User Reference" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).UserField) : Next
                    If searchtype = "Attack Class" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).AttackClass) : Next
                    If searchtype = "Attack Speed" Then For index = 0 To ItemObjects.Count - 1 : TextDecipher(index, ItemObjects(index).AttackSpeed) : Next
                End If

                If Main.RefineSearchCHECKBOX.Checked = True Then

                    For chck = 0 To RefineSearchReferenceList.Count - 1
                        a = CInt(RefineSearchReferenceList(chck))
                        If searchtype = "Item Name" Then TextDecipher(a, ItemObjects(a).ItemName)
                        If searchtype = "Item Base" Then TextDecipher(a, ItemObjects(a).ItemBase)
                        If searchtype = "Item Quality" Then TextDecipher(a, ItemObjects(a).ItemQuality)
                        If searchtype = "Mule Name" Then TextDecipher(a, ItemObjects(a).MuleName)
                        If searchtype = "Mule Account" Then TextDecipher(a, ItemObjects(a).MuleAccount)
                        If searchtype = "Realm" Then TextDecipher(a, ItemObjects(a).ItemRealm)
                        If searchtype = "User Reference" Then TextDecipher(a, ItemObjects(a).UserField)
                        If searchtype = "Attack Class" Then TextDecipher(a, ItemObjects(a).AttackClass)
                        If searchtype = "Attack Speed" Then TextDecipher(a, ItemObjects(a).AttackSpeed)
                    Next
                End If

                '--------------------------------------------------------------------------
                'Bool based searches
                '--------------------------------------------------------------------------
            Case "RuneWord", "Ladder", "Ethereal", "Hardcore", "Expansion"
                If Main.RefineSearchCHECKBOX.Checked = False Then
                    If searchtype = "RuneWord" Then For index = 0 To ItemObjects.Count - 1 : boolchecker(index, ItemObjects(index).RuneWord) : Next
                    If searchtype = "Ladder" Then For index = 0 To ItemObjects.Count - 1 : boolchecker(index, ItemObjects(index).Ladder) : Next
                    If searchtype = "Ethereal" Then For index = 0 To ItemObjects.Count - 1 : boolchecker(index, ItemObjects(index).EtherealItem) : Next
                    If searchtype = "Hardcore" Then For index = 0 To ItemObjects.Count - 1 : boolchecker(index, ItemObjects(index).HardCore) : Next
                    If searchtype = "Expansion" Then For index = 0 To ItemObjects.Count - 1 : boolchecker(index, ItemObjects(index).Expansion) : Next
                End If

                If Main.RefineSearchCHECKBOX.Checked = True Then
                    For chck = 0 To RefineSearchReferenceList.Count - 1
                        a = CInt(RefineSearchReferenceList(chck))
                        If searchtype = "RuneWord" Then boolchecker(a, ItemObjects(a).RuneWord)
                        If searchtype = "Ladder" Then boolchecker(a, ItemObjects(a).Ladder)
                        If searchtype = "Ethereal" Then boolchecker(a, ItemObjects(a).EtherealItem)
                        If searchtype = "Hardcore" Then boolchecker(a, ItemObjects(a).HardCore)
                        If searchtype = "Expansion" Then boolchecker(a, ItemObjects(a).Expansion)
                    Next
                End If

                '--------------------------------------------------------------------------
                'stats 1 to 15
                '--------------------------------------------------------------------------
            Case "Unique Attributes"
                If Main.RefineSearchCHECKBOX.Checked = False Then
                    For index = 0 To ItemObjects.Count - 1
                        If AppSettings.HideDupes = True And Main.SearchLISTBOX.Items.Contains(ItemObjects(index).ItemName) = True Then Continue For

                        If ItemObjects(index).Stat1 <> "" Then If SearchUnique(index, ItemObjects(index).Stat1) = True Then Continue For
                        If ItemObjects(index).Stat2 <> "" Then If SearchUnique(index, ItemObjects(index).Stat2) = True Then Continue For
                        If ItemObjects(index).Stat3 <> "" Then If SearchUnique(index, ItemObjects(index).Stat3) = True Then Continue For
                        If ItemObjects(index).Stat4 <> "" Then If SearchUnique(index, ItemObjects(index).Stat4) = True Then Continue For
                        If ItemObjects(index).Stat5 <> "" Then If SearchUnique(index, ItemObjects(index).Stat5) = True Then Continue For
                        If ItemObjects(index).Stat6 <> "" Then If SearchUnique(index, ItemObjects(index).Stat6) = True Then Continue For
                        If ItemObjects(index).Stat7 <> "" Then If SearchUnique(index, ItemObjects(index).Stat7) = True Then Continue For
                        If ItemObjects(index).Stat8 <> "" Then If SearchUnique(index, ItemObjects(index).Stat8) = True Then Continue For
                        If ItemObjects(index).Stat9 <> "" Then If SearchUnique(index, ItemObjects(index).Stat9) = True Then Continue For
                        If ItemObjects(index).Stat10 <> "" Then If SearchUnique(index, ItemObjects(index).Stat10) = True Then Continue For
                        If ItemObjects(index).Stat11 <> "" Then If SearchUnique(index, ItemObjects(index).Stat11) = True Then Continue For
                        If ItemObjects(index).Stat12 <> "" Then If SearchUnique(index, ItemObjects(index).Stat12) = True Then Continue For
                        If ItemObjects(index).Stat13 <> "" Then If SearchUnique(index, ItemObjects(index).Stat13) = True Then Continue For
                        If ItemObjects(index).Stat14 <> "" Then If SearchUnique(index, ItemObjects(index).Stat14) = True Then Continue For
                        If ItemObjects(index).Stat15 <> "" Then If SearchUnique(index, ItemObjects(index).Stat15) = True Then Continue For
                    Next
                End If

                If Main.RefineSearchCHECKBOX.Checked = True Then
                    For chck = 0 To RefineSearchReferenceList.Count - 1
                        Dim Index = RefineSearchReferenceList(chck)
                        If ItemObjects(Index).Stat1 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat1) = True Then Continue For
                        If ItemObjects(Index).Stat2 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat2) = True Then Continue For
                        If ItemObjects(Index).Stat3 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat3) = True Then Continue For
                        If ItemObjects(Index).Stat4 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat4) = True Then Continue For
                        If ItemObjects(Index).Stat5 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat5) = True Then Continue For
                        If ItemObjects(Index).Stat6 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat6) = True Then Continue For
                        If ItemObjects(Index).Stat7 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat7) = True Then Continue For
                        If ItemObjects(Index).Stat8 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat8) = True Then Continue For
                        If ItemObjects(Index).Stat9 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat9) = True Then Continue For
                        If ItemObjects(Index).Stat10 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat10) = True Then Continue For
                        If ItemObjects(Index).Stat11 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat11) = True Then Continue For
                        If ItemObjects(Index).Stat12 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat12) = True Then Continue For
                        If ItemObjects(Index).Stat13 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat13) = True Then Continue For
                        If ItemObjects(Index).Stat14 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat14) = True Then Continue For
                        If ItemObjects(Index).Stat15 <> "" Then If SearchUnique(Index, ItemObjects(Index).Stat15) = True Then Continue For
                    Next
                End If

                '--------------------------------------------------------------------------
                'integer based searches
                '--------------------------------------------------------------------------
            Case Else
                If Main.RefineSearchCHECKBOX.Checked = False Then
                    If searchtype = "Sockets" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).Sockets) : Next
                    If searchtype = "Chance To Block" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).ChanceToBlock) : Next
                    If searchtype = "One Hand Damage Max" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).OneHandDamageMax) : Next
                    If searchtype = "One Hand Damage Min" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).OneHandDamageMin) : Next
                    If searchtype = "Two Hand Damage Max" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).TwoHandDamageMax) : Next
                    If searchtype = "Two Hand Damage Min" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).TwoHandDamageMin) : Next
                    If searchtype = "Throw Damage Max" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).ThrowDamageMax) : Next
                    If searchtype = "Throw Damage Min" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).ThrowDamageMin) : Next
                    If searchtype = "Required Level" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).RequiredLevel) : Next
                    If searchtype = "Required Strength" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).RequiredStrength) : Next
                    If searchtype = "Required Dexterity" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).RequiredDexterity) : Next
                    If searchtype = "Item Level" Then For index = 0 To ItemObjects.Count - 1 : Numberchecker(index, ItemObjects(index).Itemlevel) : Next

                End If

                If Main.RefineSearchCHECKBOX.Checked = True Then
                    For chck = 0 To RefineSearchReferenceList.Count - 1
                        a = CInt(RefineSearchReferenceList(chck))
                        If searchtype = "Sockets" Then Numberchecker(a, ItemObjects(a).Sockets)
                        If searchtype = "Chance To Block" Then Numberchecker(a, ItemObjects(a).ChanceToBlock)
                        If searchtype = "One Hand Damage Max" Then Numberchecker(a, ItemObjects(a).OneHandDamageMax)
                        If searchtype = "One Hand Damage Min" Then Numberchecker(a, ItemObjects(a).OneHandDamageMin)
                        If searchtype = "Two Hand Damage Max" Then Numberchecker(a, ItemObjects(a).TwoHandDamageMax)
                        If searchtype = "Two Hand Damage Min" Then Numberchecker(a, ItemObjects(a).TwoHandDamageMin)
                        If searchtype = "Throw Damage Max" Then Numberchecker(a, ItemObjects(a).ThrowDamageMax)
                        If searchtype = "Throw Damage Min" Then Numberchecker(a, ItemObjects(a).ThrowDamageMin)
                        If searchtype = "Required Level" Then Numberchecker(a, ItemObjects(a).RequiredLevel)
                        If searchtype = "Required Strength" Then Numberchecker(a, ItemObjects(a).RequiredStrength)
                        If searchtype = "Required Dexterity" Then Numberchecker(a, ItemObjects(a).RequiredDexterity)
                        If searchtype = "Item Level" Then Numberchecker(a, ItemObjects(a).Itemlevel)
                    Next
                End If

        End Select

        If SearchReferenceList.Count > 0 Then
            Main.SearchLISTBOX.SelectedIndex = 0
            Main.ListboxTABCONTROL.SelectTab(1)
            Main.SearchListControlTabBUTTON.BackColor = Color.DimGray
            Main.ListControlTabBUTTON.BackColor = Color.Black
            Main.TradesListControlTabBUTTON.BackColor = Color.Black

        End If
        Main.ItemTallyTEXTBOX.Text = SearchReferenceList.Count & " - Total Matches"


    End Sub


    Sub TextDecipher(ByVal index, ByVal txtval)
        If txtval = Nothing Then txtval = "" 'To avoid null reference errors
        If AppSettings.HideDupes = True And Main.SearchLISTBOX.Items.Contains(ItemObjects(index).ItemName) = True Then Return


        If Main.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval).Contains(LCase(Main.SearchWordCOMBOBOX.Text)) = True Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval.Contains(Main.SearchWordCOMBOBOX.Text) = True Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
            End If
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) <> LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And Main.SearchWordCOMBOBOX.Text <> txtval Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
            End If
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) > LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval > Main.SearchWordCOMBOBOX.Text Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
            End If
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) < LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval < Main.SearchWordCOMBOBOX.Text Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
            End If
        End If

    End Sub

    Sub boolchecker(ByVal index, ByVal bCheck)
        If AppSettings.HideDupes = True And Main.SearchLISTBOX.Items.Contains(ItemObjects(index).ItemName) = True Then Return
        If bCheck <> True And bCheck <> False Then Return 'prevents errors

        Dim temp = Main.SearchWordCOMBOBOX.Text
        If temp = "" Then temp = True 'if nothing assume search wanted is for true comparison
        If temp <> "True" Then temp = False 'if true not entered assume search wanted is for false comparison


        If Main.SearchOperatorCOMBOBOX.Text = "Equal To" And bCheck = temp Then
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

        If Main.SearchOperatorCOMBOBOX.Text <> "Equal To" And bCheck <> temp Then ' handles greater or less than as well ???
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

    End Sub

    Sub Numberchecker(ByVal index, ByVal iValue)
        If AppSettings.HideDupes = True And Main.SearchLISTBOX.Items.Contains(ItemObjects(index).ItemName) = True Then Return
        If IsNumeric(iValue) = False Then MessageBox.Show("error in index " & index & "Value is NaN " & iValue) : Return

        Dim iCompare As Integer = Main.SearchValueNUMERICUPDWN.Value

        If Main.SearchOperatorCOMBOBOX.Text = "Equal To" And iValue = iCompare Then
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

        If Main.SearchOperatorCOMBOBOX.Text = "Not Equal To" And iValue <> iCompare Then
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

        If Main.SearchOperatorCOMBOBOX.Text = "Greater Than" And iValue > iCompare Then
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

        If Main.SearchOperatorCOMBOBOX.Text = "Less Than" And iValue < iCompare Then
            Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return
        End If

    End Sub

    Function SearchUnique(index, txtval)
        If txtval = Nothing Then txtval = "" 'To avoid null reference errors

        If Main.SearchOperatorCOMBOBOX.Text = "Equal To" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval).Contains(LCase(Main.SearchWordCOMBOBOX.Text)) = True Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval.Contains(Main.SearchWordCOMBOBOX.Text) = True Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
            End If
            Return False
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Not Equal To" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) <> LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And Main.SearchWordCOMBOBOX.Text <> txtval Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
            End If
            Return False
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Greater Than" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) > LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval > Main.SearchWordCOMBOBOX.Text Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
            End If
            Return False
        End If


        If Main.SearchOperatorCOMBOBOX.Text = "Less Than" Then
            If Main.ExactMatchCHECKBOX.Checked = False Then 'not case sensitive
                If LCase(txtval) < LCase(Main.SearchWordCOMBOBOX.Text) Then
                    Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
                End If
            End If
            If Main.ExactMatchCHECKBOX.Checked = True And txtval < Main.SearchWordCOMBOBOX.Text Then 'IS case sensitive
                Main.SearchLISTBOX.Items.Add(ItemObjects(index).ItemName) : SearchReferenceList.Add(index) : Return True
            End If
            Return False
        End If

        Return False
    End Function
End Module
