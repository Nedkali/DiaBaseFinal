Imports System.IO

Public Class ExportItemsForm
    'setup export items form - load form event handler
    Private Sub MoveItems_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DatabaseFilenameTEXTBOX.Text = Nothing
        SavedDatabasesLISTBOX.Items.Clear()
        Dim AllSavedDatabasesFileNames As Array = Nothing

        'GETS SAVED DATABASE FILES FROM DATBASE DIRECTORY AND USES THEM TO POPULATE LISTBOX
        AllSavedDatabasesFileNames = Directory.GetFiles(Application.StartupPath + "\Database\", "*").Select(Function(p) Path.GetFileName(p)).ToArray()

        For Each item In AllSavedDatabasesFileNames
            SavedDatabasesLISTBOX.Items.Add(Replace(item, ".txt", ""))
        Next
    End Sub

    'Auto selects the textbox and enters listbox selection into textbox as its selected in the listbox -- keeps textbox always selected ready for filename input
    Private Sub SavedDatabasesLISTBOX_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SavedDatabasesLISTBOX.SelectedIndexChanged
        DatabaseFilenameTEXTBOX.Text = SavedDatabasesLISTBOX.SelectedItem
        DatabaseFilenameTEXTBOX.Select()
        DatabaseFilenameTEXTBOX.SelectionLength = 0
    End Sub

    'Stops list box from selecting when nothing in the textbox is selected - keeps textbox always selected ready for filename input
    Private Sub SavedDatabasesLISTBOX_MouseDown(sender As Object, e As MouseEventArgs) Handles SavedDatabasesLISTBOX.MouseDown
        DatabaseFilenameTEXTBOX.Select()
        DatabaseFilenameTEXTBOX.SelectionLength = 0
    End Sub
End Class