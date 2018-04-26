Public Class ScriptChecker
    Private Sub ScriptChecker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.Left = Main.Left + 150 : Me.Top = Main.Top + 100
        Dim pickup As String = ""
        Dim pickupsource As String = ""
        Dim MuleLogger As String = ""
        Dim MuleLoggersource As String = ""
        Dim EntryPoint As String = ""
        Dim EntryPointsource As String = ""

        If AppSettings.EtalVersion = "KOL" Then
            Label1.Text = "" : Label2.Text = ""
            Label3.Text = "MuleLogger.js"
            Label5.Text = "DiaBaseMuleLog.dbj"
            MuleLogger = AppSettings.EtalPath + "\d2bs\kolbot\libs\DiaBaseLogger.js"
            MuleLoggersource = Application.StartupPath + "\Extras\DiaBaseLogger.js"
            EntryPoint = AppSettings.EtalPath + "\d2bs\kolbot\DiaBaseMuleLog.dbj"
            EntryPointsource = Application.StartupPath + "\Extras\DiaBaseMuleLog.dbj"
        End If

        'Rd version file verify and copy if incorrect version
        If AppSettings.EtalVersion = "NED" Then
            Label1.Text = "CPickUp.ntj"
            Label3.Text = "CBMuleLogging.ntj"
            Label5.Text = "NTDiaBaseMule.ntj"
            pickup = AppSettings.EtalPath + "\scripts\Common\CPickUp.ntj"
            pickupsource = Application.StartupPath + "\Extras\CPickUp.ntj"
            MuleLogger = AppSettings.EtalPath + "\scripts\Common\Bots\CBMuleLogging.ntj"
            MuleLoggersource = Application.StartupPath + "\Extras\CBMuleLogging.ntj"
            EntryPoint = AppSettings.EtalPath + "\scripts\NTDiaBaseMule.ntj"
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
            EntryPoint = AppSettings.EtalPath + "\scripts\NTDiaBaseMule.ntj"
            EntryPointsource = Application.StartupPath + "\Extras\NTDiaBaseMule.ntj"
        End If

        'checks pickup script
        If pickup.Length > 0 Then
            If My.Computer.FileSystem.FileExists(pickup) = True Then
                Dim ReadFile As IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(pickup)
                Dim temp = ReadFile.ReadLine()
                ReadFile.Close()
                If temp.Contains(VersionAndRevision) = True Then
                    Label4.Text = "Correct file in use"
                Else
                    My.Computer.FileSystem.CopyFile(pickupsource, pickup, True) : Label4.Text = "File Updated"
                End If
            Else
                My.Computer.FileSystem.CopyFile(pickupsource, pickup, True) : Label4.Text = "File Restored"
            End If
        End If


        'checks MuleLogger script
        If My.Computer.FileSystem.FileExists(MuleLogger) = True Then
            Dim ReadFile As IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(MuleLogger)
            Dim temp = ReadFile.ReadLine()
            ReadFile.Close()
            If temp.Contains(VersionAndRevision) = True Then
                Label4.Text = "Correct file in use"
            Else
                My.Computer.FileSystem.CopyFile(MuleLoggersource, MuleLogger, True) : Label4.Text = "File Updated"
            End If
        Else
            My.Computer.FileSystem.CopyFile(MuleLoggersource, MuleLogger, True) : Label4.Text = "File Restored"
        End If

        'entry point check and replace if necessary
        If My.Computer.FileSystem.FileExists(EntryPoint) = True Then
            Dim ReadFile As IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(EntryPoint)
            Dim temp = ReadFile.ReadLine()
            ReadFile.Close()
            If temp.Contains(VersionAndRevision) = True Then
                Label6.Text = "Correct file in use"
            Else
                My.Computer.FileSystem.CopyFile(EntryPointsource, EntryPoint, True) : Label6.Text = "File Updated"
            End If
        Else
            My.Computer.FileSystem.CopyFile(EntryPointsource, EntryPoint, True) : Label6.Text = "File Restored"
        End If

    End Sub

    Private Sub SearchBUTTON_Click(sender As Object, e As EventArgs) Handles SearchBUTTON.Click
        Me.Close()
    End Sub
End Class