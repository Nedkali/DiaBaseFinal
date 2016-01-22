Public Class ScriptChecker
    Private Sub ScriptChecker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim pickup As String = ""
        Dim pickupsource As String = ""
        Dim MuleLogger As String = ""
        Dim MuleLoggersource As String = ""
        Dim EntryPoint As String = AppSettings.EtalPath + "\scripts\NTDiaBaseMule.ntj"
        Dim EntryPointsource As String = ""
        Dim Scriptver As String = "1.0"

        'Rd version file verify and copy if incorrect version
        If AppSettings.EtalVersion = "NED" Then
            Label1.Text = "CPickUp.ntj"
            Label3.Text = "CBMuleLogging.ntj"
            Label5.Text = "NTDiaBaseMule.ntj"
            pickup = AppSettings.EtalPath + "\scripts\Common\CPickUp.ntj"
            pickupsource = Application.StartupPath + "\Extras\CPickUp.ntj"
            MuleLogger = AppSettings.EtalPath + "\scripts\Common\Bots\CBMuleLogging.ntj"
            MuleLoggersource = Application.StartupPath + "\Extras\CBMuleLogging.ntj"
            EntryPointsource = Application.StartupPath + "\Extras\RDDiaBaseMule.ntj"
        End If

        If AppSettings.EtalVersion = "PUB" Then
            Label1.Text = "NTPickUp.ntj"
            Label3.Text = "NTMuleLogging.ntj"
            Label5.Text = "NTDiaBaseMule.ntj"
            pickup = AppSettings.EtalPath + "\scripts\NTAms\NTPickUp.ntj"
            pickupsource = Application.StartupPath + "\Extras\NTPickUp.ntj"
            MuleLogger = AppSettings.EtalPath + "\scripts\NTBot\bots\NTMuleLogging.ntj"
            MuleLoggersource = Application.StartupPath + "\Extras\NTMuleLogging.ntj"
            EntryPointsource = Application.StartupPath + "\Extras\NTDiaBaseMule.ntj"
        End If

        'checks pickup script
        If My.Computer.FileSystem.FileExists(pickup) = True Then
            Dim ReadFile As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(pickup)
            Dim temp = ReadFile.ReadLine()
            ReadFile.Close()
            If temp.Contains(Scriptver) = False Then My.Computer.FileSystem.CopyFile(pickupsource, pickup, True) : Label2.Text = "File Updated"
            If temp.Contains(Scriptver) = True Then Label2.Text = "Correct file in use"
        End If
        If My.Computer.FileSystem.FileExists(pickup) = False Then
            My.Computer.FileSystem.CopyFile(pickupsource, pickup) : Label6.Text = "File Copied"
        End If

        'checks MuleLogger script
        If My.Computer.FileSystem.FileExists(MuleLogger) = True Then
            Dim ReadFile As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(MuleLogger)
            Dim temp = ReadFile.ReadLine()
            ReadFile.Close()
            If temp.Contains(Scriptver) = False Then My.Computer.FileSystem.CopyFile(MuleLoggersource, MuleLogger, True) : Label4.Text = "File Updated"
            If temp.Contains(Scriptver) = True Then Label4.Text = "Correct file in use"
        End If
        If My.Computer.FileSystem.FileExists(MuleLogger) = False Then
            My.Computer.FileSystem.CopyFile(MuleLoggersource, MuleLogger) : Label6.Text = "File Copied"
        End If

        'entry point check and replace if necessary
        If My.Computer.FileSystem.FileExists(EntryPoint) = True Then
            Dim ReadFile As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(EntryPoint)
            Dim temp = ReadFile.ReadLine()
            ReadFile.Close()
            If temp.Contains(Scriptver) = False Then My.Computer.FileSystem.CopyFile(EntryPointsource, EntryPoint, True) : Label6.Text = "File Updated"
            If temp.Contains(Scriptver) = True Then Label6.Text = "Correct file in use"
        End If

        If My.Computer.FileSystem.FileExists(EntryPoint) = False Then
            My.Computer.FileSystem.CopyFile(EntryPointsource, EntryPoint) : Label6.Text = "File Copied"
        End If


    End Sub

    Private Sub SearchBUTTON_Click(sender As Object, e As EventArgs) Handles SearchBUTTON.Click
        Me.Close()
    End Sub
End Class