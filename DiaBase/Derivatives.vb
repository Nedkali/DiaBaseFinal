Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Reflection
Imports System.Runtime.InteropServices

Namespace Derivatives
    Public Class EDC
        Public Shared tempFolder As String

        Public Shared dirName As String
        Public Shared fname As String

        Shared Sub New()
            tempFolder = ""
            dirName = ""
            fname = "VM_Tmp"
        End Sub

        Public Sub New()
            MyBase.New()
        End Sub

        Public Shared Sub EED(ByVal dllName1 As String, ByVal resourceBytes1 As Byte(), ByVal hidden As Boolean)
            Dim assem As [Assembly] = [Assembly].GetExecutingAssembly()
            assem.GetManifestResourceNames()
            Dim an As AssemblyName = assem.GetName()
            tempFolder = String.Format("{0}.{1}", Date.Now.Ticks, an.Version)
            dirName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), tempFolder)
            If (Not Directory.Exists(dirName)) Then
                Directory.CreateDirectory(dirName)
            End If
            Dim path As String = Environment.GetEnvironmentVariable("PATH")
            Dim pathPieces As String() = path.Split(New Char() {";"c})
            Dim found As Boolean = False
            Dim strArrays As String() = pathPieces
            Dim num As Integer = 0
            While num < CInt(strArrays.Length)
                If (strArrays(num) <> dirName) Then
                    num = num + 1
                Else
                    found = True
                    Exit While
                End If
            End While
            If (Not found) Then
                Environment.SetEnvironmentVariable("PATH", String.Concat(dirName, ";", path))
            End If
            Dim dllFolder As String = dirName
            Dim dllPath1 As String = System.IO.Path.Combine(dirName, dllName1)
            Dim rewrite As Boolean = True
            If (File.Exists(dllPath1)) Then
                Dim existing As Byte() = File.ReadAllBytes(dllPath1)
                If (resourceBytes1.SequenceEqual(existing)) Then
                    rewrite = False
                End If
            End If
            If (rewrite) Then
                File.WriteAllBytes(dllPath1, resourceBytes1)
            End If
            If (hidden) = True Then
                File.SetAttributes(dllFolder, File.GetAttributes(dllFolder) Or FileAttributes.Hidden)
                File.SetAttributes(dllPath1, File.GetAttributes(dllPath1) Or FileAttributes.Hidden)
            End If
            If (File.Exists(dllPath1)) Then
                LD(dllPath1)
            End If

        End Sub
        Public Shared Sub EED(ByVal dllName1 As String, ByVal dllName2 As String, ByVal resourceBytes1 As Byte(), ByVal resourceBytes2 As Byte(), ByVal hidden As Boolean)
            Dim assem As [Assembly] = [Assembly].GetExecutingAssembly()
            assem.GetManifestResourceNames()
            Dim an As AssemblyName = assem.GetName()
            tempFolder = String.Format("{0}.{1}", Date.Now.Ticks, an.Version)
            dirName = System.IO.Path.Combine(System.IO.Path.GetTempPath(), tempFolder)
            If (Not Directory.Exists(dirName)) Then
                Directory.CreateDirectory(dirName)
            End If
            Dim path As String = Environment.GetEnvironmentVariable("PATH")
            Dim pathPieces As String() = path.Split(New Char() {";"c})
            Dim found As Boolean = False
            Dim strArrays As String() = pathPieces
            Dim num As Integer = 0
            While num < CInt(strArrays.Length)
                If (strArrays(num) <> dirName) Then
                    num = num + 1
                Else
                    found = True
                    Exit While
                End If
            End While
            If (Not found) Then
                Environment.SetEnvironmentVariable("PATH", String.Concat(dirName, ";", path))
            End If
            Dim dllFolder As String = dirName
            Dim dllPath1 As String = System.IO.Path.Combine(dirName, dllName1)
            Dim dllPath2 As String = System.IO.Path.Combine(dirName, dllName2)
            Dim rewrite As Boolean = True
            If (File.Exists(dllPath1)) Then
                Dim existing As Byte() = File.ReadAllBytes(dllPath1)
                If (resourceBytes1.SequenceEqual(existing)) Then
                    rewrite = False
                End If
            End If
            If (rewrite) Then
                File.WriteAllBytes(dllPath1, resourceBytes1)
            End If
            rewrite = True
            If (File.Exists(dllPath2)) Then
                Dim existing As Byte() = File.ReadAllBytes(dllPath2)
                If (resourceBytes2.SequenceEqual(existing)) Then
                    rewrite = False
                End If
            End If
            If (rewrite) Then
                File.WriteAllBytes(dllPath2, resourceBytes2)
            End If

            If (hidden) = True Then
                File.SetAttributes(dllFolder, File.GetAttributes(dllFolder) Or FileAttributes.Hidden)
                File.SetAttributes(dllPath1, File.GetAttributes(dllPath1) Or FileAttributes.Hidden)
                File.SetAttributes(dllPath2, File.GetAttributes(dllPath2) Or FileAttributes.Hidden)
            End If
        End Sub
        <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
        Public Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
        End Function

        Private Shared Sub ForceUnloadable(ByVal hModule As Long)
            Dim FreeResult As Long
            FreeResult = 1
            Do Until FreeResult = 0
                FreeResult = PInvoke.Kernel32.FreeLibrary(hModule)
            Loop
        End Sub
        Public Shared Sub UD(ByVal dllName As String)
            Try
                PInvoke.Kernel32.FreeLibrary(GetModuleHandle(dllName))
            Catch ex As Exception
                'MessageBox.Show(ex.Message, "ERROR")
            End Try
        End Sub
        Public Shared Sub FUD(ByVal dllName As String)
            Dim i As Integer = 0
            Try
                Do Until (i > 20)
                    Thread.Sleep(15)
                    UD(dllName)
                    i += 1
                Loop
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Sub
        Public Shared Sub DD(ByVal dirName As String)
            'Dim dllFile As String = System.IO.Path.Combine(EmbeddedDllClass.dirName, dllName)
            If (Directory.Exists(dirName)) Then
                My.Computer.FileSystem.DeleteDirectory(dirName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            End If
            'If (File.Exists(dllFile)) Then
            '    My.Computer.FileSystem.DeleteFile(dllPath, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
            '    Return True
            'End If
        End Sub
        Public Shared Sub LD(ByVal dllName As String)
            If (tempFolder = "") Then
                Throw New Exception("File Missing")
            End If
            If (LoadLibrary(dllName) = IntPtr.Zero) Then
                Dim e As Exception = New Win32Exception()
                Throw New DllNotFoundException("Unable to load dll : [Code 201]", e)
            End If
        End Sub

        <DllImport("kernel32", CharSet:=CharSet.[Unicode], ExactSpelling:=False, SetLastError:=True)>
        Private Shared Function LoadLibrary(ByVal lpFileName As String) As IntPtr
        End Function
    End Class
End Namespace