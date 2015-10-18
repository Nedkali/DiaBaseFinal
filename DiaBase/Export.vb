Imports System.IO

Public Class Export
    Private Sub Export_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
End Class