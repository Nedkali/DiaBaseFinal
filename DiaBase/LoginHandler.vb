Public Class LoginHandler
    Private Sub LoginHandler_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim temp = Split(ErrMessage, ",")
        Label2.Text = temp(0)
        Label3.Text = temp(1)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim temp = Split(ErrMessage, ", ")
        Dim account As String = temp(0).ToString
        Dim realm As String = ""
        For index = 0 To temp(1).Length - 1
            If temp(1)(index) <> Nothing Then realm = realm + temp(1)(index)
        Next

        ' MessageBox.Show("bot" & realm.Length & " " & realm)
        ' MessageBox.Show("db" & ItemObjects(0).ItemRealm & " " & ItemObjects(0).ItemRealm.Length)
        For index = ItemObjects.Count - 1 To 0 Step -1
            If ItemObjects(index).MuleAccount = temp(0) And ItemObjects(index).ItemRealm = realm Then
                ' MessageBox.Show("Matched realm")
                ItemObjects.RemoveAt(index)
            End If

        Next
        Main.AllItemsLISTBOX.Items.Clear()
        Main.SearchLISTBOX.Items.Clear()
        PopulateAllItemsLISTBOX()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class