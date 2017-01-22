Imports System.IO
Imports System.IO.MemoryMappedFiles
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization.Formatters.Binary

Module D2Loader

    Private TargetProcessHandle As Integer
    Private pfnStartAddr As Integer
    Private pszLibFileRemote As String
    Private TargetBufferSize As Integer

    Public Const PROCESS_VM_READ = &H10
    Public Const TH32CS_SNAPPROCESS = &H2
    Public Const MEM_COMMIT = 4096
    Public Const PAGE_READWRITE = 4
    Public Const PROCESS_CREATE_THREAD = (&H2)
    Public Const PROCESS_VM_OPERATION = (&H8)
    Public Const PROCESS_VM_WRITE = (&H20)
    Dim DLLFileName As String
    Public Declare Function ReadProcessMemory Lib "kernel32" (
    ByVal hProcess As Integer,
    ByVal lpBaseAddress As Integer,
    ByVal lpBuffer As String,
    ByVal nSize As Integer,
    ByRef lpNumberOfBytesWritten As Integer) As Integer

    Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (
    ByVal lpLibFileName As String) As Integer

    Public Declare Function VirtualAllocEx Lib "kernel32" (
    ByVal hProcess As Integer,
    ByVal lpAddress As Integer,
    ByVal dwSize As Integer,
    ByVal flAllocationType As Integer,
    ByVal flProtect As Integer) As Integer

    Public Declare Function WriteProcessMemory Lib "kernel32" (
    ByVal hProcess As Integer,
    ByVal lpBaseAddress As Integer,
    ByVal lpBuffer As String,
    ByVal nSize As Integer,
    ByRef lpNumberOfBytesWritten As Integer) As Integer

    Public Declare Function GetProcAddress Lib "kernel32" (
    ByVal hModule As Integer, ByVal lpProcName As String) As Integer

    Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (
    ByVal lpModuleName As String) As Integer

    Public Declare Function CreateRemoteThread Lib "kernel32" (
    ByVal hProcess As Integer,
    ByVal lpThreadAttributes As Integer,
    ByVal dwStackSize As Integer,
    ByVal lpStartAddress As Integer,
    ByVal lpParameter As Integer,
    ByVal dwCreationFlags As Integer,
    ByRef lpThreadId As Integer) As Integer

    Public Declare Function OpenProcess Lib "kernel32" (
    ByVal dwDesiredAccess As Integer,
    ByVal bInheritHandle As Integer,
    ByVal dwProcessId As Integer) As Integer

    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (
    ByVal lpClassName As String,
    ByVal lpWindowName As String) As Integer

    Declare Function CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Integer) As Integer

    Dim ExeName As String = IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath)

    Public Sub Inject(ByVal x)

        TargetProcessHandle = OpenProcess(PROCESS_CREATE_THREAD Or PROCESS_VM_OPERATION Or PROCESS_VM_WRITE, False, x)
        pszLibFileRemote = "Clobot.dll"
        pfnStartAddr = GetProcAddress(GetModuleHandle("Kernel32"), "LoadLibraryA")
        TargetBufferSize = 1 + Len(pszLibFileRemote)
        Dim Rtn As Integer
        Dim LoadLibParamAdr As Integer
        LoadLibParamAdr = VirtualAllocEx(TargetProcessHandle, 0, TargetBufferSize, MEM_COMMIT, PAGE_READWRITE)
        Rtn = WriteProcessMemory(TargetProcessHandle, LoadLibParamAdr, pszLibFileRemote, TargetBufferSize, 0)
        CreateRemoteThread(TargetProcessHandle, 0, 0, pfnStartAddr, LoadLibParamAdr, 0, 0)
        CloseHandle(TargetProcessHandle)

    End Sub


    Public Sub loadD2(ByVal x)
        If D2pid > 0 Then
            Dim d2app = Process.GetProcessesByName("Game")
            For Each process In d2app
                If process.Id = D2pid Then
                    Main.ImportLogRICHTEXTBOX.AppendText("D2 already Running" & vbCrLf)
                    Return
                End If
            Next
        End If


        Dim key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Blizzard Entertainment\Diablo II")
        If key Is Nothing Then
            Main.ImportLogRICHTEXTBOX.AppendText("Error reading registry" & vbCrLf)
            Return
        End If

        Dim D2Path As String = ""
        If key.GetValue("InstallPath") Is Nothing Then
            Main.ImportLogRICHTEXTBOX.AppendText("Error reading registry" & vbCrLf)
            Return
        Else
            D2Path = key.GetValue("InstallPath").ToString()
        End If

        If D2Path(D2Path.Length - 1) <> "\" Then
            D2Path = D2Path + "\"
        End If

        Dim GamePath = D2Path & "Game.exe"

        If My.Computer.FileSystem.FileExists(GamePath) = False Then
            displayloaderror("Unable to locate Game.exe")
            Return
        End If

        Dim mmf As MemoryMappedFile = MemoryMappedFile.CreateNew("DiaBase", 314)
        If MemFile(mmf, x) = False Then Return


        Dim procstartinfo As ProcessStartInfo = New ProcessStartInfo()
        Dim ApArgs As String = " -w"
        procstartinfo.Arguments = ApArgs
        procstartinfo.FileName = GamePath
        procstartinfo.UseShellExecute = False
        procstartinfo.WorkingDirectory = D2Path

        Dim p As Process = New Process()
        p.EnableRaisingEvents = True
        p.StartInfo = procstartinfo
        Try
            p = PInvoke.Extensions.StartSuspended(p, p.StartInfo) 'loads D2 into memory
        Catch ex As Exception
            'MessageBox.Show(ex.Message)
            mmf.Dispose()
            Main.ImportLogRICHTEXTBOX.AppendText(" failed to Launch Diablo 2" & vbCrLf)
            Main.ImportLogRICHTEXTBOX.AppendText(GamePath & vbCrLf)
            Return
        End Try

        D2pid = p.Id

        'displayloaderror("Game PID = " & p.Id)'Debug print

        'blocks 2nd instance check
        Dim address As New IntPtr(&H400000 + &HF562B)
        Dim oldValue(1) As Byte
        Dim newvalue() As Byte = {&HEB, &H45}

        Try 'a287
            If Not PInvoke.Kernel32.ReadProcessMemory(p, address, oldValue) Then Main.ImportLogRICHTEXTBOX.AppendText(" failed to read window fix" & vbCrLf)
            If PInvoke.Kernel32.WriteProcessMemory(p, address, newvalue) = 0 Then Main.ImportLogRICHTEXTBOX.AppendText(" failed to write window fix" & vbCrLf)
        Catch ex As Exception
            displayloaderror("Error applying patch - Load aborted")
            displayloaderror(ex.Message)
            p.Kill()
            mmf.Dispose()
            Return
        End Try

        'If Not PInvoke.Kernel32.LoadRemoteLibrary(p, Application.StartupPath & "\DBase.dll") Then Main.ImportLogRICHTEXTBOX.AppendText(" Failed to load D2Etal.dll" & vbCrLf)

        PInvoke.Kernel32.ResumeProcess(p)

        Try
            p.WaitForInputIdle(3000)
        Catch ex As Exception
            If (ex.Message.Contains("exited") = True) Then Main.ImportLogRICHTEXTBOX.AppendText("Exited" & vbCrLf)
            'MessageBox.Show(ex.Message)
        End Try

        Try
            GlobalVars.startUp(p.Id)
        Catch ex As Exception
            MessageBox.Show("Unable to load DLL : [Code 102]", "ERROR")
        End Try

        mmf.Dispose()





    End Sub

    Public Sub displayloaderror(ByVal txt)

        Main.ImportLogRICHTEXTBOX.AppendText(txt & vbCrLf)

    End Sub

    Function MemFile(ByVal mmf, ByVal x)
        Dim prof As Profile

        'hope to reserve zero for single player use in dll
        Dim temp = 0
        If ItemObjects(x).ItemRealm = "USWest" Then temp = 1
        If ItemObjects(x).ItemRealm = "USEast" Then temp = 2
        If ItemObjects(x).ItemRealm = "Asia" Then temp = 3
        If ItemObjects(x).ItemRealm = "Europe" Then temp = 4

        prof.Account = ItemObjects(x).MuleAccount
        prof.AccPass = ItemObjects(x).MulePass
        prof.CharName = ItemObjects(x).MuleName
        prof.Difficulty = Chr(0)
        'prof.Realm = Chr(0) ' for testing on singleplayer chars
        prof.Realm = Chr(temp)

        prof.MpqFile = Right(AppSettings.MpqFile, Len(AppSettings.MpqFile) - InStrRev(AppSettings.MpqFile, "\"))
        'MessageBox.Show(prof.MpqFile)
        If AppSettings.EtalVersion = "NED" Then
            prof.FilePath = AppSettings.EtalPath & "\scripts\Configs\" & ItemObjects(x).ItemRealm & "\AMS\MuleInventory\"
        Else
            prof.FilePath = AppSettings.EtalPath & "\scripts\AMS\MuleInventory\"
        End If
        prof.FilePath = prof.FilePath.Replace("\\", "\")
        prof.FilePath = prof.FilePath.Replace("\", "/")

        Dim Ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(prof))
        Dim ByteArray(Marshal.SizeOf(prof) - 1) As Byte
        Marshal.StructureToPtr(prof, Ptr, False)
        Marshal.Copy(Ptr, ByteArray, 0, Marshal.SizeOf(prof))
        Marshal.FreeHGlobal(Ptr)

        Try
            Dim Stream As MemoryMappedViewStream = mmf.CreateViewStream()
            Dim writer As BinaryWriter = New BinaryWriter(Stream)
            For index = 0 To ByteArray.Count - 1
                writer.Write(Chr(ByteArray(index)))
            Next

            writer.Close()
        Catch ex As Exception
            Main.ImportLogRICHTEXTBOX.AppendText(ex.Message & vbCrLf)
            Return False
        End Try

        Return True

    End Function

End Module