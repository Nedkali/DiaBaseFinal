Public Class UserInput
    
    'YES BUTTON PRESS
    Private Sub UserInputYesBUTTON_Click(sender As Object, e As EventArgs) Handles UserInputYesBUTTON.Click
        If UserInputTEXTBOX.Visible = True And UserInputTEXTBOX.Text <> "" Then DialogResult = Windows.Forms.DialogResult.Yes : Me.Close() Else If UserInputTEXTBOX.Visible = True Then UserInputTEXTBOX.Select()
        If UserInputTEXTBOX.Visible = False Then DialogResult = Windows.Forms.DialogResult.Yes : Me.Close()
    End Sub

    'NO BUTTON PRESS
    Private Sub UserInputNoBUTTON_Click(sender As Object, e As EventArgs) Handles UserInputNoBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No : Me.Close()
    End Sub

    Private Sub UserInput_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Left = Main.Left + 185 : Me.Top = Main.Top + 90         'Set Central window location

    End Sub

    Private Sub UserInput_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'APPLY THE FANCY DIABLO2 FONT TO HEADERS AND BUTTONS ONLY IF THE DiabloFont1.TTF FONT FILE IS PRESENT IN THE EXTRAS DIRECTORY
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then
            pfc.AddFontFile(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF")

            UserInputHeaderLABEL.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            UserInputNoBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            UserInputYesBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

    End Sub
End Class