Public Class SkinSelector
    Public a = Main.AllItemsLISTBOX.SelectedIndex



    Private Sub SkinSelectorDownBUTTON_Click(sender As Object, e As EventArgs) Handles SkinSelectorDownBUTTON.Click
        Dim a As Integer = NumericUpDown1.Value
        a = a - 1
        If a < 0 Then Return
        NumericUpDown1.Value = a
        Dim filename = ImageArray(a)
        SkinSelectorPICTUREBOX.Load("Skins\" + filename + ".jpg")

    End Sub

    Private Sub SkinSelectorCancelBUTTON_Click(sender As Object, e As EventArgs) Handles SkinSelectorCancelBUTTON.Click
        Me.Close()
    End Sub

    Private Sub SkinSelectorUpBUTTON_Click(sender As Object, e As EventArgs) Handles SkinSelectorUpBUTTON.Click
        Dim a As Integer = NumericUpDown1.Value
        a = a + 1
        If a > 659 Then Return
        NumericUpDown1.Value = a
        Dim filename = ImageArray(a)

        If My.Computer.FileSystem.FileExists("Skins\" + filename + ".jpg") = True Then SkinSelectorPICTUREBOX.Load("Skins\" + filename + ".jpg")

    End Sub

    Private Sub SkinSelector_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim a = Main.AllItemsLISTBOX.SelectedIndex
        NumericUpDown1.Value = ItemObjects(a).ItemImage
        Dim filename As String = ImageArray(NumericUpDown1.Value)

        If My.Computer.FileSystem.FileExists("Skins\" + filename + ".jpg") = True Then SkinSelectorPICTUREBOX.Load("Skins\" + filename + ".jpg")

    End Sub

    Private Sub SkinSelectorSelectBUTTON_Click(sender As Object, e As EventArgs) Handles SkinSelectorSelectBUTTON.Click
        Dim a = Main.AllItemsLISTBOX.SelectedIndex
        ItemObjects(a).ItemImage = NumericUpDown1.Value
        Me.Close()

    End Sub
End Class