Public Class ErrorHandlerForm
    'Yesy Dialog Button Press
    Private Sub ErrorTrapYesBUTTON_Click(sender As Object, e As EventArgs) Handles ErrorTrapYesBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.Yes
    End Sub
    'No Dialog Button Press
    Private Sub ErrorTrapNoBUTTON_Click(sender As Object, e As EventArgs) Handles ErrorTrapNoBUTTON.Click
        DialogResult = Windows.Forms.DialogResult.No
    End Sub

    'Setup D2 Fonts on form load
    Private Sub ErrorHandlerForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Computer.FileSystem.FileExists(AppSettings.InstallPath + "\Extras\DiabloFont1.TTF") = True Then

            'Main Form Headers (16 point size)
            'ErrorTrapHeaderLABEL.Font = New Font(pfc.Families(0), 12, FontStyle.Regular)
            'ErrorTrapYesBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
            ' ErrorTrapNoBUTTON.Font = New Font(pfc.Families(0), 9, FontStyle.Regular)
        End If

    End Sub
End Class