Imports System.IO

Public Class Export
    Private Sub Export_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GetFileList()
    End Sub

    Private Sub GetFileList()
        Dim temp = AppSettings.CurrentDatabase.Split("\")
        Dim temp1 = temp(temp.Length - 1)
        'Get All Database Files From InstallPath\Databases Directory
        Dim AllSavedDatabaseFileNames As Array = Nothing
        AllSavedDatabaseFileNames = Directory.GetFiles(AppSettings.InstallPath + "\Databases\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        'Populate Listbox items and Combobox Drop Down List items
        ListBox1.Items.Clear() '                                                                             Delete old database file name lists
        For Each DatabaseFileName In AllSavedDatabaseFileNames
            If DatabaseFileName.indexof(".txt") > -1 And DatabaseFileName <> temp1 Then '                        Check for correct .TXT or .txt file extension
                If DatabaseFileName.indexof(".txt") > -1 Then DatabaseFileName = Replace(DatabaseFileName, ".txt", "") '    Remove lower case extension if there is one
                ListBox1.Items.Add(DatabaseFileName)  '                                                      Apply Cropped Database File Name to lists
            End If
        Next
    End Sub

    Private Sub CreateFileBUTTON_Click(sender As Object, e As EventArgs) Handles CreateFileBUTTON.Click

        Dim temp = TextBox1.Text
        If temp.Length = 0 Then Return
        temp = Replace(temp, ".txt", "")
        temp = temp + ".txt"

        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Databases\" + temp) = False Then
            Dim CreateFile As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(AppSettings.InstallPath + "\Databases\" + temp, False)
            CreateFile.Close()
        End If
        TextBox1.Text = ""
        GetFileList()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If ListBox1.SelectedIndex > -1 Then
            ExportButton.Enabled = True
        End If
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.Close()
    End Sub

    Private Sub ExportButton_Click(sender As Object, e As EventArgs) Handles ExportButton.Click
        If Main.SearchLISTBOX.SelectedIndices.Count < 1 Then Return

        Dim x As Integer = 0
        Dim temp = ListBox1.SelectedItem

        temp = AppSettings.InstallPath + "\Databases\" + temp + ".txt"

        If My.Computer.FileSystem.FileExists(temp) = True Then
            'backup before edit TODO

            Try
                Dim LogWriter As System.IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(temp, True)

                For index = 0 To Main.SearchLISTBOX.SelectedIndices.Count - 1
                    x = SearchReferenceList(index)
                    LogWriter.WriteLine("--------------------")
                    LogWriter.WriteLine(ItemObjects(x).ItemName)
                    LogWriter.WriteLine(ItemObjects(x).Itemlevel)
                    LogWriter.WriteLine(ItemObjects(x).ItemRealm)
                    LogWriter.WriteLine(ItemObjects(x).ItemBase)
                    LogWriter.WriteLine(ItemObjects(x).ItemQuality)
                    LogWriter.WriteLine(ItemObjects(x).RequiredCharacter)
                    LogWriter.WriteLine(ItemObjects(x).EtherealItem)
                    LogWriter.WriteLine(ItemObjects(x).Sockets)
                    LogWriter.WriteLine(ItemObjects(x).RuneWord)
                    LogWriter.WriteLine(ItemObjects(x).ThrowDamageMin)
                    LogWriter.WriteLine(ItemObjects(x).ThrowDamageMax)
                    LogWriter.WriteLine(ItemObjects(x).OneHandDamageMin)
                    LogWriter.WriteLine(ItemObjects(x).OneHandDamageMax)
                    LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMin)
                    LogWriter.WriteLine(ItemObjects(x).TwoHandDamageMax)
                    LogWriter.WriteLine(ItemObjects(x).Defense)
                    LogWriter.WriteLine(ItemObjects(x).ChanceToBlock)
                    LogWriter.WriteLine(ItemObjects(x).QuantityMin)
                    LogWriter.WriteLine(ItemObjects(x).QuantityMax)
                    LogWriter.WriteLine(ItemObjects(x).DurabilityMin)
                    LogWriter.WriteLine(ItemObjects(x).DurabilityMax)
                    LogWriter.WriteLine(ItemObjects(x).RequiredStrength)
                    LogWriter.WriteLine(ItemObjects(x).RequiredDexterity)
                    LogWriter.WriteLine(ItemObjects(x).RequiredLevel)
                    LogWriter.WriteLine(ItemObjects(x).AttackClass)
                    LogWriter.WriteLine(ItemObjects(x).AttackSpeed)
                    LogWriter.WriteLine(ItemObjects(x).Stat1)
                    LogWriter.WriteLine(ItemObjects(x).Stat2)
                    LogWriter.WriteLine(ItemObjects(x).Stat3)
                    LogWriter.WriteLine(ItemObjects(x).Stat4)
                    LogWriter.WriteLine(ItemObjects(x).Stat5)
                    LogWriter.WriteLine(ItemObjects(x).Stat6)
                    LogWriter.WriteLine(ItemObjects(x).Stat7)
                    LogWriter.WriteLine(ItemObjects(x).Stat8)
                    LogWriter.WriteLine(ItemObjects(x).Stat9)
                    LogWriter.WriteLine(ItemObjects(x).Stat10)
                    LogWriter.WriteLine(ItemObjects(x).Stat11)
                    LogWriter.WriteLine(ItemObjects(x).Stat12)
                    LogWriter.WriteLine(ItemObjects(x).Stat13)
                    LogWriter.WriteLine(ItemObjects(x).Stat14)
                    LogWriter.WriteLine(ItemObjects(x).Stat15)
                    LogWriter.WriteLine(ItemObjects(x).MuleName)
                    LogWriter.WriteLine(ItemObjects(x).MuleAccount)
                    LogWriter.WriteLine(ItemObjects(x).MulePass)
                    LogWriter.WriteLine(ItemObjects(x).PickitAccount)
                    LogWriter.WriteLine(ItemObjects(x).HardCore)
                    LogWriter.WriteLine(ItemObjects(x).Ladder)
                    LogWriter.WriteLine(ItemObjects(x).Expansion)
                    LogWriter.WriteLine(ItemObjects(x).UserField)
                    LogWriter.WriteLine(ItemObjects(x).ItemImage)
                    LogWriter.WriteLine(ItemObjects(x).ImportTime)
                    LogWriter.WriteLine(ItemObjects(x).ImportDate)
                Next

                LogWriter.Close()

                If CheckBox1.Checked = True Then
                    Dim a As Integer
                    Dim b As Integer
                    For index = Main.SearchLISTBOX.SelectedIndices.Count - 1 To 0 Step -1
                        a = Main.SearchLISTBOX.SelectedIndices(index)
                        b = SearchReferenceList(a)
                        Main.SearchLISTBOX.Items.RemoveAt(a)
                        Main.AllItemsLISTBOX.Items.RemoveAt(b)
                        ItemObjects.RemoveAt(b)
                        SearchReferenceList.RemoveAt(a)
                    Next
                    Main.SearchLISTBOX.SelectedItem = -1
                End If
                Return
            Catch ex As Exception
                'error hadling TODO
                MessageBox.Show("ERROR writing to file")
                Return
            End Try

        End If
        'error hadling TODO
        MessageBox.Show("Export Failed")

    End Sub
End Class