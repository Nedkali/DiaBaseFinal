Imports System.IO
Imports System.IO.MemoryMappedFiles

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

        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Blizzard Entertainment\Diablo II")
        If key Is Nothing Then
            Main.ImportLogRICHTEXTBOX.Text = "Error reading registry"
            Return
        End If

        Dim D2Path As String = key.GetValue("InstallPath").ToString()
        If D2Path = Nothing Then Return

        If D2Path(D2Path.Length - 1) <> "\" Then
            D2Path = D2Path + "\"
        End If

        Dim GamePath = D2Path & "Game.exe"
        Dim gfxpath = D2Path & "D2gfx.dll"

        'displayloaderror(GamePath)'debug message
        'displayloaderror(gfxpath)'debug message

        If My.Computer.FileSystem.FileExists(GamePath) = False Then
            displayloaderror("Unable to locate Game.exe")
            Return
        End If

        If My.Computer.FileSystem.FileExists(gfxpath) = False Then
            displayloaderror("Unable to locate D2gfx.dll")
            Return
        End If

        ' load routine
        Dim mmf As MemoryMappedFile = MemoryMappedFile.CreateNew("DiaBaseMule", 82)
        MemFile2(mmf, x)

        Dim myprocess As Process = New Process()
        myprocess.EnableRaisingEvents = True
        Dim ApArgs As String = "-w"

        myprocess.StartInfo.Arguments = ApArgs
        myprocess.StartInfo.FileName = GamePath
        myprocess.StartInfo.UseShellExecute = False

        Dim p As Process = New Process()

        myprocess.StartInfo.WorkingDirectory = D2Path

        p = PInvoke.Extensions.StartSuspended(p, myprocess.StartInfo) 'loads D2 into memory

        displayloaderror("Game PID = " & p.Id)

        'blocks 2nd instance check
        Dim oldValue(1) As Byte
        Dim newvalue() As Byte = {&HEB, &H45}
        Dim address As New IntPtr(&H6FA80000 + &HB6B0)
        Try 'a287

            If Not PInvoke.Kernel32.LoadRemoteLibrary(p, gfxpath) Then Main.ImportLogRICHTEXTBOX.AppendText(" Failed to load d2gfx" & vbCrLf)
            If Not PInvoke.Kernel32.ReadProcessMemory(p, address, oldValue) Then Main.ImportLogRICHTEXTBOX.AppendText(" failed to read window fix" & vbCrLf)
            If PInvoke.Kernel32.WriteProcessMemory(p, address, newvalue) = 0 Then Main.ImportLogRICHTEXTBOX.AppendText(" failed to write window fix" & vbCrLf)
        Catch e As System.Exception
            displayloaderror("Error applying patch - Load aborted")
            displayloaderror(e.Message)
            p.Kill()
            mmf.Dispose()
            Return
        End Try

        If Not PInvoke.Kernel32.LoadRemoteLibrary(p, Application.StartupPath & "\DBase.dll") Then Main.ImportLogRICHTEXTBOX.AppendText(" Failed to load D2Etal.dll")

        PInvoke.Kernel32.ResumeProcess(p)
        p.WaitForInputIdle(5000)
        mmf.Dispose()
        D2pid = p.Id

    End Sub

    Public Sub displayloaderror(ByVal txt)

        Main.ImportLogRICHTEXTBOX.AppendText(txt & vbCrLf)


    End Sub

    Function MemFile2(ByVal mmf, ByVal x)
        Try
            Dim Stream As MemoryMappedViewStream = mmf.CreateViewStream()
            Dim writer As BinaryWriter = New BinaryWriter(Stream)

            For y = 0 To ItemObjects(x).MuleAccount.Length - 1 : writer.Write(ItemObjects(x).MuleAccount(y)) : Next
            For a = writer.BaseStream.Position To 23 : writer.Write(Chr(0)) : Next

            For y = 0 To ItemObjects(x).MulePass.Length - 1 : writer.Write(ItemObjects(x).MulePass(y)) : Next
            For a = writer.BaseStream.Position To 35 : writer.Write(Chr(0)) : Next

            For y = 0 To ItemObjects(x).MuleName.Length - 1 : writer.Write(ItemObjects(x).MuleName(y)) : Next
            For a = writer.BaseStream.Position To 51 : writer.Write(Chr(0)) : Next

            Dim temp = 0
            If ItemObjects(x).ItemRealm = "USWest" Then temp = 1
            If ItemObjects(x).ItemRealm = "USEast" Then temp = 2
            If ItemObjects(x).ItemRealm = "Asia" Then temp = 3
            If ItemObjects(x).ItemRealm = "Europe" Then temp = 4

            writer.Write(Chr(temp))
            writer.Write(Chr(0)) 'difficulty

            Dim temp1 = "Mulelogger.ntj"
            For y = 0 To temp1.Length - 1 : writer.Write(temp1(y)) : Next
            For a = writer.BaseStream.Position To 84 : writer.Write("") : Next

            writer.Close()
        Catch ex As Exception
            MessageBox.Show("Error mmf - ?mmf still open")
            Return False

        End Try

        Return True

    End Function

End Module