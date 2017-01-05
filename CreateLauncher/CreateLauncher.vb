<Assembly: System.Reflection.AssemblyTitle("DD's Launcher")> 'FileVersionInfo.FileDescription = AssemblyTitle
<Assembly: System.Reflection.AssemblyDescription("Launches File starting with the same name from current directory or sub folder")> 'FileVersionInfo.Comments = AssemblyDescription
<Assembly: System.Reflection.AssemblyFileVersion("1.0.0.0")> 'FileVersionInfo.FileVersion = AssemblyFileVersion

<Assembly: System.Reflection.AssemblyProduct("DD's Launcher")> 'FileVersionInfo.ProductName = AssemblyProduct
<Assembly: System.Reflection.AssemblyInformationalVersion("1, 0, 0, 0")> 'FileVersionInfo.ProductVersion = AssemblyInformationalVersion

<Assembly: System.Reflection.AssemblyCompany("Devang Bhatt")> 'FileVersionInfo.CompanyName = AssemblyCompany

<Assembly: System.Reflection.AssemblyCopyright("Copyright ©  2016")> 'FileVersionInfo.LegalCopyright = AssemblyCopyright
<Assembly: System.Reflection.AssemblyTrademark("")> 'FileVersionInfo.LegalTrademarks = AssemblyTrademark

<Assembly: System.Runtime.InteropServices.ComVisible(False)>
<Assembly: System.Runtime.InteropServices.Guid("f77d67e2-dbb4-4def-903c-802153b28eb1")>

<Assembly: System.Reflection.AssemblyVersion("1.0.0.0")>

<Assembly: System.Resources.NeutralResourcesLanguage("en")>

<Assembly: System.Reflection.AssemblyCulture("")>
<Assembly: System.Reflection.ObfuscateAssembly(True, StripAfterObfuscation:=True)>
<Assembly: System.Reflection.Obfuscation(ApplyToMembers:=True, StripAfterObfuscation:=True)>

Module CreateLauncher

    'Directory From Where the App is called
    Private CurrentDirectory As String = System.Environment.CurrentDirectory
    'Directory From Where the App is stored
    Private AppDir As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().CodeBase.Replace("file:///", "")) 'System.Reflection.Assembly.GetExecutingAssembly().Location

    Private ExeRelativePath32 As String = ""
    Private ExeRelativePath64 As String = ""

    Private EXE_CheckSum_32 As String = ""
    Private EXE_CheckSum_64 As String = ""

    Private OsSystemBit As Integer = 8 * System.IntPtr.Size

    'Launcher Name without extention - For Launcher it is the self, directory and exe name - For CreateLauncher it is only directory and exe name 
    'Taken from argument
    Private LauncherName As String = "" 'System.IO.Path.GetFileNameWithoutExtension(System.Environment.GetCommandLineArgs(0))

    Private ExePath32 As String = CreateLauncher.CurrentDirectory & ExeRelativePath32
    Private Available32 As Boolean = System.IO.File.Exists(ExePath32)

    Private ExePath64 As String = CreateLauncher.CurrentDirectory & ExeRelativePath64
    Private Available64 As Boolean = System.IO.File.Exists(ExePath64)

    Private CustomCmd32 As String = ""
    Private CustomCmd64 As String = ""

    Private IconFilePath As String = ""

    Private AdminMode As Boolean = False
    Private UseShellExec As Boolean = False
    Private IsSilent As Boolean = False
    Private AppTitle As String = ""
    Private WindowStyle As String = "HIDDEN"
    Private IsConsoleApp As Boolean = False
    Private IsCustomCmd As Boolean = False

    Private BatchFile32 As String = ""
    Private BatchFile64 As String = ""

    Private IsBatchExe As Boolean = False
    Private IsRunningInConsole As Boolean = Not (System.Console.Title = System.Environment.GetCommandLineArgs()(0) AndAlso
            System.Environment.GetCommandLineArgs()(0) = System.Reflection.Assembly.GetEntryAssembly().Location AndAlso
            System.Environment.CurrentDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location))
    Private Sub ShowMessage(Message As String, Optional ExitApp As Boolean = False)
        If IsRunningInConsole Then
            System.Console.WriteLine(Message)
        Else
            System.Windows.Forms.MessageBox.Show(Message)
        End If
        If ExitApp Then
            System.Environment.Exit(1)
        End If
    End Sub


    Public Sub ShowUsage()
        Dim UsageInfo As String = "Usage: CreateLauncher.exe ""Application Name""" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "Custom Options:" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/app: /dir: /exe: /x32: /x64: /title:] They Take Arguments" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/debug /admin /shellexec /silent] Admin is always Shellexec" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/normal /max /min] - Window Styles - Default is HIDDEN - Last One Passed is Used" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/ico: /noicon] Icon File" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/console] ConsoleApp (No ShellExec)" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/cmd: /cmd32: /cmd64:] CustomCmd 32 will run only on 32bit and 64 will run only on 64 bit (No Admin, No ShellExec, Normal Console Only)" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "[/bat: /bat32: /bat64:] Bat File to Embed in Exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "Possible Scenarios" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName.exe {Launcher}" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName\AppName.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName\AppName64.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName.exe {Launcher}" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName\x64\AppName.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName\x86\AppName.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName.exe {Launcher}" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\x64\AppName.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\x86\AppName.exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName.exe {Launcher}" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName (32-Bit).exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "\AppName (64-Bit).exe" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "In-Name Detection Keywords: " & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "x64, x86-64, x86_64, 64Bit, 64-Bit, 64_Bit, 64, Win64, amd64" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "x32, x86, 32Bit, 32-Bit, 32_Bit, 32, Win32, amd32" & Microsoft.VisualBasic.vbCrLf
        UsageInfo += "" & Microsoft.VisualBasic.vbCrLf
        ShowMessage(UsageInfo, True)
    End Sub

    Sub Main(args As String())
        If args.Length = 0 Then
            ShowUsage()
        End If
        Dim IsDebug As Boolean = False, AppVar As String = "", DirVar As String = "", ExeVar As String = "", X32Var As String = "", X64Var As String = ""
        For Each arg As String In args
            If arg.ToLowerInvariant.Trim = "/debug".ToLowerInvariant.Trim Then
                IsDebug = True
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/app:") Then
                AppVar = arg.Substring("/app:".Length)
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/dir:") Then
                DirVar = arg.Substring("/dir:".Length)
                If Not System.IO.Directory.Exists(System.IO.Path.Combine(CurrentDirectory, DirVar)) Then
                    ShowMessage("Directory Not Found! " & arg.Substring("/dir:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/exe:") Then
                ExeVar = System.IO.Path.GetFileNameWithoutExtension(arg.Substring("/exe:".Length))
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/x32:") Then
                X32Var = arg.Substring("/x32:".Length)
                If Not System.IO.Path.IsPathRooted(X32Var) Then
                    X32Var = System.IO.Path.Combine(CurrentDirectory, X32Var)
                End If
                If Not System.IO.File.Exists(X32Var) Then
                    ShowMessage("File Not Found! " & arg.Substring("/x32:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/x64:") Then
                X64Var = arg.Substring("/x64:".Length)
                If Not System.IO.Path.IsPathRooted(X64Var) Then
                    X64Var = System.IO.Path.Combine(CurrentDirectory, X64Var)
                End If
                If Not System.IO.File.Exists(X64Var) Then
                    ShowMessage("File Not Found! " & arg.Substring("/x64:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/admin") Then
                AdminMode = True
                UseShellExec = True
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/shellexec") Then
                UseShellExec = True
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/title:") Then
                AppTitle = arg.Substring("/title:".Length)
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/silent") Then
                IsSilent = True
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/normal") Then
                WindowStyle = "NORMAL"
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/max") Then
                WindowStyle = "MAXIMIZED"
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/min") Then
                WindowStyle = "MINIMIZED"
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/ico:") Then
                IconFilePath = arg.Substring("/ico:".Length)
                If Not System.IO.File.Exists(IconFilePath) Then
                    ShowMessage("File Not Found! " & arg.Substring("/ico:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/noicon") Then
                IconFilePath = " "
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/console") Then
                IsConsoleApp = True
                UseShellExec = False
                WindowStyle = "NORMAL"
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/cmd:") Then
                CustomCmd32 = arg.Substring("/cmd:".Length)
                CustomCmd64 = CustomCmd32
                AdminMode = False
                IsConsoleApp = True
                UseShellExec = False
                WindowStyle = "NORMAL"
                IsCustomCmd = True
                If IconFilePath.Trim.Length = 0 Then IconFilePath = " "
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/cmd32:") Then
                CustomCmd32 = arg.Substring("/cmd32:".Length)
                AdminMode = False
                IsConsoleApp = True
                UseShellExec = False
                WindowStyle = "NORMAL"
                IsCustomCmd = True
                If IconFilePath.Trim.Length = 0 Then IconFilePath = " "
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/cmd64:") Then
                CustomCmd64 = arg.Substring("/cmd64:".Length)
                AdminMode = False
                IsConsoleApp = True
                UseShellExec = False
                WindowStyle = "NORMAL"
                IsCustomCmd = True
                If IconFilePath.Trim.Length = 0 Then IconFilePath = " "
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/bat:") Then
                IsBatchExe = True
                BatchFile32 = arg.Substring("/bat:".Length)
                BatchFile64 = BatchFile32
                If Not System.IO.File.Exists(IconFilePath) Then
                    ShowMessage("File Not Found! " & arg.Substring("/bat:".Length), True)
                End If
                If Not IsBatchFileOK(BatchFile32) Then
                    ShowMessage("Can Not handle %* in Batch File " & arg.Substring("/bat:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/bat32:") Then
                IsBatchExe = True
                BatchFile32 = arg.Substring("/bat32:".Length)
                If Not System.IO.File.Exists(IconFilePath) Then
                    ShowMessage("File Not Found! " & arg.Substring("/bat32:".Length), True)
                End If
                If Not IsBatchFileOK(BatchFile32) Then
                    ShowMessage("Can Not handle %* in Batch File " & arg.Substring("/bat32:".Length), True)
                End If
            ElseIf arg.ToLowerInvariant.Trim.StartsWith("/bat64:") Then
                IsBatchExe = True
                BatchFile64 = arg.Substring("/bat64:".Length)
                If Not System.IO.File.Exists(IconFilePath) Then
                    ShowMessage("File Not Found! " & arg.Substring("/bat64:".Length), True)
                End If
                If Not IsBatchFileOK(BatchFile64) Then
                    ShowMessage("Can Not handle %* in Batch File " & arg.Substring("/bat64:".Length), True)
                End If
            Else
                LauncherName = arg
            End If
        Next
        If AppVar.Trim.Length > 0 Then
            LauncherName = AppVar
        ElseIf DirVar.Trim.Length > 0 Then
            If DirVar.Contains(System.IO.Path.DirectorySeparatorChar) Then
                LauncherName = System.IO.Path.GetFileName(DirVar)
            Else
                LauncherName = DirVar
            End If
        ElseIf ExeVar.Trim.Length > 0 Then
            LauncherName = ExeVar
        ElseIf X64Var.Trim.Length > 0 Then
            LauncherName = System.IO.Path.GetFileNameWithoutExtension(X64Var)
        ElseIf X32Var.Trim.Length > 0 Then
            LauncherName = System.IO.Path.GetFileNameWithoutExtension(X32Var)
        End If


        Dim DetailsFromExe As String = ""
        Dim lvi As LauncherVersionInfo = Nothing

        If IsCustomCmd Then
            lvi = New LauncherVersionInfo(AppTitle, "1.0.0.0")
        ElseIf IsBatchExe Then
            If AppTitle.Trim.Length = 0 Then
                If BatchFile64.Trim.Length > 0 Then
                    AppTitle = System.IO.Path.GetFileNameWithoutExtension(BatchFile64)
                ElseIf BatchFile32.Trim.Length > 0 Then
                    AppTitle = System.IO.Path.GetFileNameWithoutExtension(BatchFile32)
                End If
            End If
            lvi = New LauncherVersionInfo(AppTitle, "1.0.0.0")
        Else

            Dim FolderName As String = CreateLauncher.CurrentDirectory
            If DirVar.Trim.Length > 0 Then
                If System.IO.Directory.Exists(System.IO.Path.Combine(FolderName, DirVar)) Then
                    FolderName = System.IO.Path.Combine(FolderName, DirVar)
                End If
            Else
                If System.IO.Directory.Exists(System.IO.Path.Combine(FolderName, LauncherName)) Then
                    FolderName = System.IO.Path.Combine(FolderName, LauncherName)
                End If
            End If

            If X32Var.Trim.Length > 0 Then
                ExePath32 = X32Var
            Else
                If ExeVar.Trim.Length > 0 Then
                    ExePath32 = FindExe(FolderName, ExeVar, False)
                Else
                    ExePath32 = FindExe(FolderName, LauncherName, False)
                End If

            End If
            ExeRelativePath32 = ExePath32.Replace(CreateLauncher.CurrentDirectory, "")
            If ExePath32.Trim.Length > 0 Then
                Available32 = True
                EXE_CheckSum_32 = ComputeFileHash(ExePath32)
            End If
            If X64Var.Trim.Length > 0 Then
                ExePath64 = X64Var
            Else
                If ExeVar.Trim.Length > 0 Then
                    ExePath64 = FindExe(FolderName, ExeVar, True)
                Else
                    ExePath64 = FindExe(FolderName, LauncherName, True)
                End If
            End If
            ExeRelativePath64 = ExePath64.Replace(CreateLauncher.CurrentDirectory, "")
            If ExePath64.Trim.Length > 0 Then
                Available64 = True
                EXE_CheckSum_64 = ComputeFileHash(ExePath64)
            End If

            If Available64 And OsSystemBit = 64 Then
                'Take Details from x64
                DetailsFromExe = ExePath64
            ElseIf Available32 Then
                'Take Details from x32
                DetailsFromExe = ExePath32
            Else
                'x32 not available only x64 available but system is 32 bit
                If Available64 Then
                    DetailsFromExe = ExePath64
                Else
                    ShowMessage("Can Not continue!", True)
                End If
            End If

            'Get File Informations
            Dim fvi As System.Diagnostics.FileVersionInfo
            fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(DetailsFromExe)
            lvi = New LauncherVersionInfo(fvi)

            If IconFilePath.Length = 0 Then 'No Trim so we can have exe with default icons for exe
                Try
                    'Extract Icon to Temp Directory
                    'iconsext.exe /save "C:\Users\Devang\Cloud\Dropbox\Development\Launcher\Launcher\bin\Debug\CCleaner (64-Bit).exe" "C:\Users\Devang\AppData\Local\Temp\" -icons
                    Dim TsudaKageyu As New TsudaKageyu.IconExtractor(DetailsFromExe)
                    'IconFilePath = System.IO.Path.GetTempFileName & ".ico"
                    'IconFilePath = System.IO.Path.GetTempPath & LauncherName & ".ico"
                    IconFilePath = LauncherName.Replace(" ", "_") & ".ico"
                    Using fs As New System.IO.FileStream(System.IO.Path.Combine(CurrentDirectory, IconFilePath), System.IO.FileMode.Create)
                        TsudaKageyu.Save(0, fs)
                    End Using
                Catch
                    If IconFilePath.Trim.Length > 0 AndAlso System.IO.File.Exists(IconFilePath) Then
                        System.IO.File.Delete(IconFilePath)
                    End If
                    'Continue Without Error with No Icon
                    IconFilePath = ""
                    'System.Windows.Forms.MessageBox.Show("Error Extracting Icon! maybe forgot to use /noicon?")
                    'System.Environment.Exit(1)
                End Try
            End If

        End If
        If IsDebug Then
            Dim DebugText As String = ""
            DebugText += "CurrentDirectory: " & CurrentDirectory & Microsoft.VisualBasic.vbCrLf
            DebugText += "AppDir: " & AppDir & Microsoft.VisualBasic.vbCrLf
            DebugText += "OsSystemBit: " & OsSystemBit & Microsoft.VisualBasic.vbCrLf
            DebugText += "ExePath32: " & ExePath32 & Microsoft.VisualBasic.vbCrLf
            DebugText += "ExePath64: " & ExePath64 & Microsoft.VisualBasic.vbCrLf
            ShowMessage(DebugText, True)
        End If


        'Actual Rebuild Process
        Dim Provider As New Microsoft.VisualBasic.VBCodeProvider
        Dim Parameters As New System.CodeDom.Compiler.CompilerParameters
        Parameters.GenerateExecutable = True
        Parameters.OutputAssembly = LauncherName & ".exe"
        Parameters.IncludeDebugInformation = False
        Parameters.GenerateInMemory = False


        If IsConsoleApp Then
            If IconFilePath.Trim.Length > 0 Then
                Parameters.CompilerOptions = " /reference:System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:exe /platform:anycpu /win32icon:" & IconFilePath
            Else
                Parameters.CompilerOptions = " /reference:System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:exe /platform:anycpu"
            End If
        ElseIf IsBatchExe Then
            Dim ResourceString As String = "/res:'"
            If BatchFile32.Trim.Length > 0 Then ResourceString &= BatchFile32 & ",BatchFile32.bat,Private' "
            If BatchFile64.Trim.Length > 0 Then ResourceString &= BatchFile64 & ",BatchFile64.bat,Private' "
            ResourceString = ResourceString.Trim
            If IconFilePath.Trim.Length > 0 Then
                Parameters.CompilerOptions = " /reference:System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:winexe /platform:anycpu " & ResourceString & " /win32icon:" & IconFilePath
            Else
                Parameters.CompilerOptions = " /reference:System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:winexe /platform:anycpu " & ResourceString
            End If
            ShowMessage(Parameters.CompilerOptions)
            ShowMessage("WIP - Batch Compilation", True)
        Else
            If IconFilePath.Trim.Length > 0 Then
                Parameters.CompilerOptions = " / reference : System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:winexe /platform:anycpu /win32icon:" & IconFilePath
            Else
                Parameters.CompilerOptions = " /reference:System.dll,System.Windows.Forms.dll /optimize /optionstrict+ /target:winexe /platform:anycpu"
            End If
        End If


        Dim SourceCode As String = CustomizeSourceCode(lvi, UseCheckSum:=False, ComVisible:=False, ObfuscateAssembly:=True)
        Dim Results As System.CodeDom.Compiler.CompilerResults = Provider.CompileAssemblyFromSource(Parameters, SourceCode)

        If System.IO.File.Exists(IconFilePath) Then
            System.IO.File.Delete(IconFilePath)
        ElseIf System.IO.File.Exists(System.IO.Path.Combine(CurrentDirectory, IconFilePath)) Then
            System.IO.File.Delete(System.IO.Path.Combine(CurrentDirectory, IconFilePath))
        End If

        If Not IsSilent Then
            If (Results.Errors.HasErrors) Then
                ShowMessage("Can not customize for App: " & LauncherName)
            Else
                ShowMessage("Customized for App: " & LauncherName & Microsoft.VisualBasic.vbCrLf & "Relative Paths: " & Microsoft.VisualBasic.vbCrLf & "64-Bit @ " & ExeRelativePath64 & Microsoft.VisualBasic.vbCrLf & "32-Bit @ " & ExeRelativePath32)
            End If
        End If

    End Sub

    ' Calculates a file's hash value and returns it as a byte array.
    Private Function ComputeFileHash(ByVal fileName As String) As String
        Dim RetVar(0) As Byte
        If System.IO.File.Exists(fileName) Then
            Try
                Dim ourHashAlg As System.Security.Cryptography.HashAlgorithm = System.Security.Cryptography.HashAlgorithm.Create("MD5")
                Dim fileToHash As System.IO.FileStream = New System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                RetVar = ourHashAlg.ComputeHash(fileToHash)
                fileToHash.Close()
            Catch ex As System.IO.IOException
                ShowMessage("Program Open! Cannot get lock to compute hash!")
            End Try
        End If
        Return System.Text.Encoding.Unicode.GetString(RetVar)
    End Function
    Private Function HasChanged(filename As String, hash As String) As Boolean
        Dim newHash As Byte() = System.Text.Encoding.Unicode.GetBytes(ComputeFileHash(filename))
        Dim oldHash As Byte() = System.Text.Encoding.Unicode.GetBytes(hash)
        For i As Integer = 0 To newHash.Length - 1
            If newHash(i) <> oldHash(i) Then
                Return True
            End If
        Next
        Return False
    End Function

    'Check for variables not allowed
    Function IsBatchFileOK(BatchFilePathAndName As String) As Boolean
        Dim BatchContent As String = ""
        Using sr As System.IO.StreamReader = New System.IO.StreamReader(BatchFilePathAndName)
            BatchContent = sr.ReadToEnd
        End Using
        If BatchContent.Contains("%*") Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function FindExe(SearchPath As String, FileStartingName As String, Optional Win64 As Boolean = False) As String
        Dim RetVar As String = ""
        Dim SearchFolder As String = ""

        Dim x64 As String() = {"x64", "x86-64", "x86_64", "64Bit", "64-Bit", "64_Bit", "64", "Win64", "amd64"}
        Dim x32 As String() = {"x32", "x86", "32Bit", "32-Bit", "32_Bit", "32", "Win32", "amd32"}

        Dim AddArray As String() = {}
        Dim RemoveArray As String() = {}

        Dim FolderFound As Boolean = False

        If Win64 Then
            AddArray = x64
            RemoveArray = x32
        Else
            AddArray = x32
            RemoveArray = x64
        End If

        For Each PossibleFolder In AddArray
            SearchFolder = System.IO.Path.Combine(SearchPath, PossibleFolder)
            If System.IO.Directory.Exists(SearchFolder) Then
                FolderFound = True
                Exit For
            End If
        Next


        If Not FolderFound Then
            'No Folders found check for files in the SearchPath
            SearchFolder = SearchPath
        End If

        Dim FilesFound As New System.Collections.Generic.List(Of String)
        For Each FileItem As String In System.IO.Directory.GetFiles(SearchFolder, "*.exe")
            Dim ItemName As String = System.IO.Path.GetFileName(FileItem).ToLowerInvariant.Trim
            If ItemName.StartsWith(FileStartingName.ToLowerInvariant.Trim) Then
                FilesFound.Add(FileItem)
            End If
        Next

        If FilesFound.Count > 1 Then
            For i As Integer = FilesFound.Count - 1 To 0 Step -1
                Dim FileName As String = System.IO.Path.GetFileName(FilesFound.Item(i))
                For Each FileEnding As String In RemoveArray
                    If System.Text.RegularExpressions.Regex.Match(FileName, FileStartingName & ".*" & FileEnding & ".*.exe", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value = FileName Then
                        FilesFound.RemoveAt(i)
                        Exit For
                    End If
                Next
            Next
            If FilesFound.Count > 1 Then
                Dim FilesTemp As New System.Collections.Generic.List(Of String)
                For i As Integer = FilesFound.Count - 1 To 0 Step -1
                    Dim FileName As String = System.IO.Path.GetFileName(FilesFound.Item(i))
                    For Each FileEnding As String In AddArray
                        If System.Text.RegularExpressions.Regex.Match(FileName, FileStartingName & ".*" & FileEnding & ".*.exe", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Value = FileName Then
                            FilesTemp.Add(FilesFound.Item(i))
                            Exit For
                        End If
                    Next
                Next
                FilesFound = FilesTemp
            End If
        End If

        If FilesFound.Count = 1 Then
            RetVar = FilesFound.Item(0)
        ElseIf FilesFound.Count = 0 Then
            RetVar = ""
        Else
            If Win64 Then
                ShowMessage("Found Multiple Exes for 64-Bit", True)
            Else
                ShowMessage("Found Multiple Exes for 32-Bit", True)
            End If
        End If

        Return RetVar
    End Function
    Private Function CustomizeSourceCode(LVI As LauncherVersionInfo, Optional UseCheckSum As Boolean = True, Optional ComVisible As Boolean = False, Optional ObfuscateAssembly As Boolean = False) As String
        Dim SourceCode As String = LVI.ToString

        SourceCode += "Module Launcher" & Microsoft.VisualBasic.vbCrLf
        SourceCode += "    Private CurrentDirectory As String = System.Environment.CurrentDirectory" & Microsoft.VisualBasic.vbCrLf
        SourceCode += "    Private AppDir As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().CodeBase.Replace(""file:///"", """"))" & Microsoft.VisualBasic.vbCrLf

        If IsCustomCmd Then
            If CustomCmd32.Trim.Length > 0 Then SourceCode += "    Private CustomCmd32 As String = """ & CustomCmd32 & """" & Microsoft.VisualBasic.vbCrLf
            If CustomCmd64.Trim.Length > 0 Then SourceCode += "    Private CustomCmd64 As String = """ & CustomCmd64 & """" & Microsoft.VisualBasic.vbCrLf
        ElseIf IsBatchExe Then
            If BatchFile32.Trim.Length > 0 Then SourceCode += "    Private BatchFile32 As String = ""BatchFile32.bat""" & Microsoft.VisualBasic.vbCrLf
            If BatchFile64.Trim.Length > 0 Then SourceCode += "    Private BatchFile64 As String = ""BatchFile64.bat""" & Microsoft.VisualBasic.vbCrLf
        Else
            SourceCode += "    Private ExeRelativePath32 As String = """ & ExeRelativePath32 & """" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    Private ExeRelativePath64 As String = """ & ExeRelativePath64 & """" & Microsoft.VisualBasic.vbCrLf
            If UseCheckSum Then
                SourceCode += "    Private EXE_CheckSum_32 As String = """ & EXE_CheckSum_32 & """" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "    Private EXE_CheckSum_64 As String = """ & EXE_CheckSum_64 & """" & Microsoft.VisualBasic.vbCrLf
            End If

            SourceCode += "    Private ExePath32 As String = Launcher.AppDir & ExeRelativePath32" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    Private Available32 As Boolean = System.IO.File.Exists(ExePath32)" & Microsoft.VisualBasic.vbCrLf

            SourceCode += "    Private ExePath64 As String = Launcher.AppDir & ExeRelativePath64" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    Private Available64 As Boolean = System.IO.File.Exists(ExePath64)" & Microsoft.VisualBasic.vbCrLf
        End If

        SourceCode += "    Private OsSystemBit As Integer = 8 * System.IntPtr.Size" & Microsoft.VisualBasic.vbCrLf
        SourceCode += "    Private LauncherName As String = System.IO.Path.GetFileNameWithoutExtension(System.Environment.GetCommandLineArgs(0))" & Microsoft.VisualBasic.vbCrLf

        SourceCode += "    Sub Main(args As String())" & Microsoft.VisualBasic.vbCrLf
        If IsCustomCmd Then
            SourceCode += "        Dim Command As String = """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If OsSystemBit = 64 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Command = CustomCmd64" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        ElseIf OsSystemBit = 32 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Command = CustomCmd32" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If Command.Trim.Length > 0 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Using cmdProcess As New System.Diagnostics.Process" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                With cmdProcess" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .StartInfo = New System.Diagnostics.ProcessStartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    With .StartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .FileName = ""cmd.exe""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .WorkingDirectory = CurrentDirectory" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        If args.Length > 0 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                            .Arguments = ""/c "" & Command & "" "" & String.Join("" "", args)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        Else" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                            .Arguments = ""/c "" & Command" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .UseShellExecute = False" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .Start()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .WaitForExit()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf
        ElseIf IsBatchExe Then
            SourceCode += "        Dim Command As String = """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        For i As Integer = 0 To args.Length - 1" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            If args(i).Contains("" "") Then args(i) = """""""" & args(i) & """"""""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Next" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim BatchFile As String = """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If OsSystemBit = 64 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            BatchFile = BatchFile64" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        ElseIf OsSystemBit = 32 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            BatchFile = BatchFile32" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim BatchFileContent As String = """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Using Stream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceFullName(BatchFile))" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Using sr As New System.IO.StreamReader(Stream)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                Using tr As System.IO.TextReader = sr" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    BatchFileContent = tr.ReadToEnd" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        'Inject SHIFT for passing current exe path and location as %0" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        BatchFileContent = ""@Echo Off"" & Microsoft.VisualBasic.vbCrLf & ""SHIFT"" & Microsoft.VisualBasic.vbCrLf & BatchFileContent" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim TempFileName As String = System.IO.Path.GetTempPath() + System.Guid.NewGuid().ToString().Replace("" - "", """") + "".bat""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Do Until Not System.IO.File.Exists(TempFileName)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            TempFileName = System.IO.Path.GetTempPath() + System.Guid.NewGuid().ToString().Replace("" - "", """") + "".bat""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Loop" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Using fs As New System.IO.StreamWriter(TempFileName, False)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            fs.WriteLine(BatchFileContent)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If TempFileName.Trim.Length > 0 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Using cmdProcess As New System.Diagnostics.Process" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                With cmdProcess" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .StartInfo = New System.Diagnostics.ProcessStartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    With .StartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .FileName = ""cmd.exe""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .WorkingDirectory = AppDir" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        If args.Length > 0 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                            .Arguments = "" / c "" & TempFileName & "" """""" & System.IO.Path.Combine(AppDir, LauncherName & "".exe"") & """""" "" & String.Join("" "", args)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        Else" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                            .Arguments = "" / c "" & TempFileName & "" """""" & System.IO.Path.Combine(AppDir, LauncherName & "".exe"") & """"""""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .UseShellExecute = False" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .Start()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .WaitForExit()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End Using" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If System.IO.File.Exists(TempFileName) Then System.IO.File.Delete(TempFileName)" & Microsoft.VisualBasic.vbCrLf
        Else
            SourceCode += "        Dim ExePath As String = """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If OsSystemBit = 64 And Available64 Then" & Microsoft.VisualBasic.vbCrLf
            If UseCheckSum Then
                SourceCode += "            If HasChanged(ExePath64, EXE_CheckSum_64) Then" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "                System.Windows.Forms.MessageBox.Show(""64-Bit exe changed? Checksum Do Not match!"")" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "                Exit Sub" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "            Else" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "                ExePath = ExePath64" & Microsoft.VisualBasic.vbCrLf
            If UseCheckSum Then
                SourceCode += "            End If" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "        ElseIf Available32 Then" & Microsoft.VisualBasic.vbCrLf
            If UseCheckSum Then
                SourceCode += "            If HasChanged(ExePath32, EXE_CheckSum_32) Then" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "                System.Windows.Forms.MessageBox.Show(""32-Bit exe changed? Checksum Do Not match!"")" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "                Exit Sub" & Microsoft.VisualBasic.vbCrLf
                SourceCode += "            Else" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "                ExePath = ExePath32" & Microsoft.VisualBasic.vbCrLf
            If UseCheckSum Then
                SourceCode += "            End If" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "        Else" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            System.Windows.Forms.MessageBox.Show(""Exe(s) Not found!"")" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Exit Sub" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf

            SourceCode += "        Using cmdProcess As New System.Diagnostics.Process" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            With cmdProcess" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                .StartInfo = New System.Diagnostics.ProcessStartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                With .StartInfo" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .FileName = ExePath" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .WorkingDirectory = System.IO.Path.GetDirectoryName(ExePath)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    If args.Length > 0 Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                        .Arguments = """""""" & String.Join("" "", args) & """"""""" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .LoadUserProfile = True" & Microsoft.VisualBasic.vbCrLf
            If UseShellExec Then
                SourceCode += "                    .UseShellExecute = True" & Microsoft.VisualBasic.vbCrLf
            Else
                SourceCode += "                    .UseShellExecute = False" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "                    .RedirectStandardError = False" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .RedirectStandardInput = False" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                    .RedirectStandardOutput = False" & Microsoft.VisualBasic.vbCrLf
            If AdminMode Then
                SourceCode += "                    .Verb = ""runas""" & Microsoft.VisualBasic.vbCrLf
            End If
            If WindowStyle = "NORMAL" OrElse WindowStyle = "MAXIMIZED" OrElse WindowStyle = "MINIMIZED" Then
                SourceCode += "                    .CreateNoWindow = False" & Microsoft.VisualBasic.vbCrLf
            Else
                SourceCode += "                    .CreateNoWindow = True" & Microsoft.VisualBasic.vbCrLf
            End If

            If WindowStyle = "NORMAL" Then
                SourceCode += "                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Normal" & Microsoft.VisualBasic.vbCrLf
            ElseIf WindowStyle = "MAXIMIZED" Then
                SourceCode += "                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Maximized" & Microsoft.VisualBasic.vbCrLf
            ElseIf WindowStyle = "MINIMIZED" Then
                SourceCode += "                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Minimized" & Microsoft.VisualBasic.vbCrLf
            Else
                SourceCode += "                    .WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden" & Microsoft.VisualBasic.vbCrLf
            End If
            SourceCode += "                End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Try" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                .Start()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                .WaitForExit()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Catch" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End Try" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End With" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End Using" & Microsoft.VisualBasic.vbCrLf
        End If
        SourceCode += "    End Sub" & Microsoft.VisualBasic.vbCrLf
        If Not IsCustomCmd And UseCheckSum Then
            SourceCode += "    Private Function ComputeFileHash(ByVal fileName As String) As String" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim RetVar(0) As Byte" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        If System.IO.File.Exists(fileName) Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Try" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                Dim ourHashAlg As System.Security.Cryptography.HashAlgorithm = System.Security.Cryptography.HashAlgorithm.Create(""MD5"")" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                Dim fileToHash As System.IO.FileStream = New System.IO.FileStream(fileName, System.IO.FileMode.Open, System.IO.FileAccess.Read)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                RetVar = ourHashAlg.ComputeHash(fileToHash)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                fileToHash.Close()" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            Catch ex As System.IO.IOException" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                System.Windows.Forms.MessageBox.Show(""Program Open! Cannot Get lock To compute hash!"")" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End Try" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Return System.Text.Encoding.Unicode.GetString(RetVar)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    End Function" & Microsoft.VisualBasic.vbCrLf
            'Batch Support Function
            SourceCode += "    Private Function ResourceFullName(BatchName As String) As String" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        For Each ResourceName As String In System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            If ResourceName.ToLowerInvariant.EndsWith(BatchName.ToLowerInvariant) Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                Return ResourceName" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Next" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Return """"" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    End Function" & Microsoft.VisualBasic.vbCrLf

            SourceCode += "    Private Function HasChanged(filename As String, hash As String) As Boolean" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim newHash As Byte() = System.Text.Encoding.Unicode.GetBytes(ComputeFileHash(filename))" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Dim oldHash As Byte() = System.Text.Encoding.Unicode.GetBytes(hash)" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        For i As Integer = 0 To newHash.Length - 1" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            If newHash(i) <> oldHash(i) Then" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "                Return True" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "            End If" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Next" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "        Return False" & Microsoft.VisualBasic.vbCrLf
            SourceCode += "    End Function" & Microsoft.VisualBasic.vbCrLf
        End If
        SourceCode += "End Module" & Microsoft.VisualBasic.vbCrLf

        Return SourceCode
    End Function

    Public Class LauncherVersionInfo
        Public AssemblyTitle As String = ""
        Public AssemblyDescription As String = ""
        Public AssemblyFileVersion As String = ""
        Public AssemblyProduct As String = "" 'Portable Apps Takes this as Menu Entry
        Public AssemblyInformationalVersion As String = ""
        Public AssemblyCompany As String = ""
        Public AssemblyCopyright As String = ""
        Public AssemblyTrademark As String = ""
        Public AssemblyVersion As String = "" 'need 1.0.0.0 Format 1, 0, 0, 0 not accepted as in AssemblyFileVersion
        Public ComVisible As Boolean = False
        Public ObfuscateAssembly As Boolean = False
        Public Sub New()
            MyBase.New
        End Sub
        Public Sub New(Title As String, Version As String)
            Me.New
            Me.AssemblyTitle = Title
            Me.AssemblyDescription = Title
            Me.AssemblyFileVersion = Version
            Me.AssemblyProduct = Title
            Me.AssemblyInformationalVersion = Version
            Me.AssemblyVersion = Version
        End Sub
        Public Sub New(FileVersionInfo As System.Diagnostics.FileVersionInfo)
            Me.New()
            If FileVersionInfo.FileDescription.Trim.Length = 0 Then
                Me.AssemblyTitle = LauncherName
            Else
                Me.AssemblyTitle = FileVersionInfo.FileDescription
            End If
            Me.AssemblyDescription = FileVersionInfo.Comments
            Me.AssemblyFileVersion = FileVersionInfo.FileVersion
            If AppTitle.Trim.Length > 0 Then
                Me.AssemblyProduct = AppTitle
            ElseIf FileVersionInfo.ProductName.Trim.Length = 0 Then
                Me.AssemblyProduct = LauncherName
            Else
                Me.AssemblyProduct = FileVersionInfo.ProductName
            End If
            Me.AssemblyInformationalVersion = FileVersionInfo.ProductVersion
            Me.AssemblyCompany = FileVersionInfo.CompanyName
            Me.AssemblyCopyright = FileVersionInfo.LegalCopyright
            Me.AssemblyTrademark = FileVersionInfo.LegalTrademarks
            Me.AssemblyVersion = FileVersionInfo.FileMajorPart & "." & FileVersionInfo.FileMinorPart & "." & FileVersionInfo.FileBuildPart & "." & FileVersionInfo.FilePrivatePart
            Me.ComVisible = ComVisible
            Me.ObfuscateAssembly = ObfuscateAssembly
        End Sub
        Public Overrides Function ToString() As String
            Dim RetVar As String = ""
            If AssemblyTitle.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyTitle(""" & AssemblyTitle & " (via DD's Launcher)"")>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyDescription.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyDescription(""" & AssemblyDescription & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyFileVersion.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyFileVersion(""" & AssemblyFileVersion & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyProduct.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyProduct(""" & AssemblyProduct & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyInformationalVersion.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyInformationalVersion(""" & AssemblyInformationalVersion & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyCompany.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyCompany(""" & AssemblyCompany & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyCopyright.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyCopyright(""" & AssemblyCopyright & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyTrademark.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyTrademark(""" & AssemblyTrademark & """)>" & Microsoft.VisualBasic.vbCrLf
            If ComVisible Then RetVar += "<Assembly: System.Runtime.InteropServices.ComVisible(True)>" & Microsoft.VisualBasic.vbCrLf
            If ComVisible Then RetVar += "<Assembly: System.Runtime.InteropServices.Guid(""" & System.Guid.NewGuid.ToString() & """)>" & Microsoft.VisualBasic.vbCrLf
            If AssemblyVersion.Trim.Length > 0 Then RetVar += "<Assembly: System.Reflection.AssemblyVersion(""" & AssemblyVersion & """)>" & Microsoft.VisualBasic.vbCrLf
            RetVar += "<Assembly: System.Resources.NeutralResourcesLanguage(""en"")>" & Microsoft.VisualBasic.vbCrLf
            RetVar += "<Assembly: System.Reflection.AssemblyCulture("""")>" & Microsoft.VisualBasic.vbCrLf
            If ObfuscateAssembly Then
                RetVar += "<Assembly: System.Reflection.ObfuscateAssembly(True, StripAfterObfuscation:=True)>" & Microsoft.VisualBasic.vbCrLf
                RetVar += "<Assembly: System.Reflection.Obfuscation(ApplyToMembers:=True, StripAfterObfuscation:=True)>" & Microsoft.VisualBasic.vbCrLf
            End If
            Return RetVar
        End Function
    End Class

End Module
