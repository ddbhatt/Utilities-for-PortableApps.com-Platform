Namespace TsudaKageyu

    Friend Class NativeMethods
        ' Methods
        <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function EnumResourceNames(ByVal hModule As System.IntPtr, ByVal lpszType As System.IntPtr, ByVal lpEnumFunc As ENUMRESNAMEPROC, ByVal lParam As System.IntPtr) As Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function FindResource(ByVal hModule As System.IntPtr, ByVal lpName As System.IntPtr, ByVal lpType As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function FreeLibrary(ByVal hModule As System.IntPtr) As Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function GetCurrentProcess() As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("psapi.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function GetMappedFileName(ByVal hProcess As System.IntPtr, ByVal lpv As System.IntPtr, ByVal lpFilename As System.Text.StringBuilder, ByVal nSize As Integer) As Integer
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function LoadLibraryEx(ByVal lpFileName As String, ByVal hFile As System.IntPtr, ByVal dwFlags As System.UInt32) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function LoadResource(ByVal hModule As System.IntPtr, ByVal hResInfo As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function LockResource(ByVal hResData As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function QueryDosDevice(ByVal lpDeviceName As String, ByVal lpTargetPath As System.Text.StringBuilder, ByVal ucchMax As Integer) As Integer
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)>
        <System.Security.SuppressUnmanagedCodeSecurity>
        Public Shared Function SizeofResource(ByVal hModule As System.IntPtr, ByVal hResInfo As System.IntPtr) As System.UInt32
        End Function

    End Class

    <System.Runtime.InteropServices.UnmanagedFunctionPointer(System.Runtime.InteropServices.CallingConvention.Winapi, SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode)>
    <System.Security.SuppressUnmanagedCodeSecurity>
    Friend Delegate Function ENUMRESNAMEPROC(ByVal hModule As System.IntPtr, ByVal lpszType As System.IntPtr, ByVal lpszName As System.IntPtr, ByVal lParam As System.IntPtr) As Boolean

    Public Class IconUtil

        Private Delegate Function GetIconDataDelegate(ByVal icon As System.Drawing.Icon) As Byte()

        Shared _getIconData As GetIconDataDelegate

        Shared Sub New()
            'Create a dynamic method to access Icon.iconData private field.
            Dim dm As New System.Reflection.Emit.DynamicMethod("GetIconData", GetType(Byte()), New System.Type() {GetType(System.Drawing.Icon)}, GetType(System.Drawing.Icon))
            Dim fi As System.Reflection.FieldInfo = GetType(System.Drawing.Icon).GetField("iconData", (System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance))
            Dim gen As System.Reflection.Emit.ILGenerator = dm.GetILGenerator
            gen.Emit(System.Reflection.Emit.OpCodes.Ldarg_0)
            gen.Emit(System.Reflection.Emit.OpCodes.Ldfld, fi)
            gen.Emit(System.Reflection.Emit.OpCodes.Ret)
            IconUtil._getIconData = DirectCast(dm.CreateDelegate(GetType(GetIconDataDelegate)), GetIconDataDelegate)
        End Sub

        'Split an Icon consists of multiple icons into an array of Icon each consists of single icons.
        Public Shared Function Split(ByVal icon As System.Drawing.Icon) As System.Drawing.Icon()
            If (icon Is Nothing) Then
                Throw New System.ArgumentNullException("icon")
            End If
            '// Get an .ico file in memory, then split it into separate icons.
            Dim src As Byte() = IconUtil.GetIconData(icon)
            Dim splitIcons As New System.Collections.Generic.List(Of System.Drawing.Icon)
            Dim count As Integer = System.BitConverter.ToUInt16(src, 4)
            Dim i As Integer
            For i = 0 To count - 1
                Dim length As Integer = System.BitConverter.ToInt32(src, 6 + 16 * i + 8) '// ICONDIRENTRY.dwBytesInRes
                Dim offset As Integer = System.BitConverter.ToInt32(src, 6 + 16 * i + 12) '// ICONDIRENTRY.dwImageOffset
                Using dst As System.IO.BinaryWriter = New System.IO.BinaryWriter(New System.IO.MemoryStream((6 + 16 + length)))
                    '// Copy ICONDIR and set idCount to 1.
                    dst.Write(src, 0, 4)
                    dst.Write(CShort(1))
                    '// Copy ICONDIRENTRY and set dwImageOffset to 22.
                    dst.Write(src, 6 + 16 * i, 12) '// ICONDIRENTRY except dwImageOffset
                    dst.Write(22)
                    '// Copy a picture.
                    dst.Write(src, offset, length)
                    '// Create an icon from the in-memory file.
                    dst.BaseStream.Seek(0, System.IO.SeekOrigin.Begin)
                    splitIcons.Add(New System.Drawing.Icon(dst.BaseStream))
                End Using
            Next i
            Return splitIcons.ToArray
        End Function

        'Converts an Icon to a GDI+ Bitmap preserving the transparent area.
        Public Shared Function ToBitmap(ByVal icon As System.Drawing.Icon) As System.Drawing.Bitmap
            Dim bitmap As System.Drawing.Bitmap
            If (icon Is Nothing) Then
                Throw New System.ArgumentNullException("icon")
            End If
            'Quick workaround: Create an .ico file in memory, then load it as a Bitmap.
            Using ms As System.IO.MemoryStream = New System.IO.MemoryStream
                icon.Save(ms)
                Using bmp As System.Drawing.Bitmap = DirectCast(System.Drawing.Image.FromStream(ms), System.Drawing.Bitmap)
                    bitmap = New System.Drawing.Bitmap(bmp)
                End Using
            End Using
            Return bitmap
        End Function

        'Gets the bit depth of an Icon.
        'This method takes into account the PNG header. If the icon has multiple variations, this method returns the bit depth of the first variation.
        Public Shared Function GetBitCount(ByVal icon As System.Drawing.Icon) As Integer
            If (icon Is Nothing) Then
                Throw New System.ArgumentNullException("icon")
            End If

            'Get an .ico file in memory, then read the header.
            Dim data As Byte() = GetIconData(icon)
            If (data.Length >= 51 _
                AndAlso data(22) = &H89 AndAlso data(23) = &H50 AndAlso data(24) = &H4E AndAlso data(25) = &H47 _
                AndAlso data(26) = &HD AndAlso data(27) = &HA AndAlso data(28) = &H1A AndAlso data(29) = &HA _
                AndAlso data(30) = &H0 AndAlso data(31) = &H0 AndAlso data(32) = &H0 AndAlso data(33) = &HD _
                AndAlso data(34) = &H49 AndAlso data(35) = &H48 AndAlso data(36) = &H44 AndAlso data(37) = &H52
                ) Then
                'The picture is PNG. Read IHDR chunk.
                Select Case (data(47))
                    Case 0
                        Return data(46)
                    Case 2
                        Return data(46) * 3
                    Case 3
                        Return data(46)
                    Case 4
                        Return data(46) * 2
                    Case 6
                        Return data(46) * 4
                    Case Else
                        '// NOP
                        'break;
                End Select
            ElseIf data.Length >= 22 Then
                'The picture is not PNG. Read ICONDIRENTRY structure.
                Return System.BitConverter.ToUInt16(data, 12)
            End If
            Throw New System.ArgumentException("The icon is corrupt. Couldn't read the header.", "icon")
        End Function

        Private Shared Function GetIconData(ByVal icon As System.Drawing.Icon) As Byte()
            Dim data As Byte() = _getIconData.Invoke(icon)
            If (data.Length > 0) Then
                Return data
            Else
                Using ms As New System.IO.MemoryStream
                    icon.Save(ms)
                    Return ms.ToArray
                End Using
            End If
        End Function

    End Class

    Public Class IconExtractor

        'Flags for LoadLibraryEx().
        Private Const LOAD_LIBRARY_AS_DATAFILE As UInteger = &H2

        'Resource types for EnumResourceNames().
        Private Shared ReadOnly RT_ICON As System.IntPtr = System.IntPtr.op_Explicit(3)
        Private Shared ReadOnly RT_GROUP_ICON As System.IntPtr = System.IntPtr.op_Explicit(14)

        Private Const MAX_PATH As Integer = 260

        Private iconData As Byte()() = Nothing   '// Binary data Of Each icon.

        'Gets the full path of the associated file.
        Public Property FileName As String

        'Gets the count of the icons in the associated file.
        Public ReadOnly Property Count As Integer
            Get
                Return Me.iconData.Length
            End Get
        End Property

        'Initializes a new instance of the IconExtractor class from the specified file name.
        Public Sub New(ByVal fileName As String)
            Me.Initialize(fileName)
        End Sub

        'Extracts an icon from the file.
        Public Function GetIcon(ByVal index As Integer) As System.Drawing.Icon
            If index < 0 OrElse Me.Count <= index Then
                Throw New System.ArgumentOutOfRangeException("index")
            End If
            'Create an Icon from the .ico file in memory.
            Using ms As New System.IO.MemoryStream(Me.iconData(index))
                Return New System.Drawing.Icon(ms)
            End Using
        End Function

        'Extracts all the icons from the file.
        Public Function GetAllIcons() As System.Drawing.Icon()
            Dim icons As New System.Collections.Generic.List(Of System.Drawing.Icon)
            For i As Integer = 0 To Me.Count - 1
                icons.Add(GetIcon(i))
            Next i
            Return icons.ToArray
        End Function

        'Save an icon To the specified output Stream.
        Public Sub Save(ByVal index As Integer, ByVal outputStream As System.IO.Stream)
            If index < 0 OrElse Me.Count <= index Then
                Throw New System.ArgumentOutOfRangeException("index")
            End If
            If (outputStream Is Nothing) Then
                Throw New System.ArgumentNullException("outputStream")
            End If
            Dim data As Byte() = Me.iconData(index)
            outputStream.Write(data, 0, data.Length)
        End Sub

        Private Sub Initialize(ByVal fileName As String)
            If (fileName Is Nothing) Then
                Throw New System.ArgumentNullException("fileName")
            End If
            Dim hModule As System.IntPtr = System.IntPtr.Zero
            Try
                hModule = NativeMethods.LoadLibraryEx(fileName, System.IntPtr.Zero, LOAD_LIBRARY_AS_DATAFILE)
                If (hModule = System.IntPtr.Zero) Then
                    Throw New System.ComponentModel.Win32Exception
                End If
                Me.FileName = Me.GetFileName(hModule)
                Dim tmpData As New System.Collections.Generic.List(Of Byte())
                Dim callback As ENUMRESNAMEPROC = Function(ByVal h As System.IntPtr, ByVal t As System.IntPtr, ByVal name As System.IntPtr, ByVal l As System.IntPtr)
                                                      'RT_GROUP_ICON resource consists of a GRPICONDIR and GRPICONDIRENTRY's.
                                                      Dim dir As Byte() = Me.GetDataFromResource(hModule, RT_GROUP_ICON, name)

                                                      'Calculate the size of an entire .icon file.
                                                      Dim count As Integer = System.BitConverter.ToUInt16(dir, 4) 'GRPICONDIR.idCount
                                                      Dim len As Integer = 6 + 16 * count ' sizeof(ICONDIR) + sizeof(ICONDIRENTRY) * count
                                                      For i As Integer = 0 To count - 1
                                                          len = (len + System.BitConverter.ToInt32(dir, 6 + 14 * i + 8)) 'GRPICONDIRENTRY.dwBytesInRes
                                                      Next i
                                                      Using dst As System.IO.BinaryWriter = New System.IO.BinaryWriter(New System.IO.MemoryStream(len))
                                                          'Copy GRPICONDIR to ICONDIR.
                                                          dst.Write(dir, 0, 6)
                                                          Dim picOffset As Integer = 6 + 16 * count 'sizeof(ICONDIR) + sizeof(ICONDIRENTRY) * count

                                                          For i As Integer = 0 To count - 1
                                                              'Load the picture.
                                                              Dim id As UShort = System.BitConverter.ToUInt16(dir, 6 + 14 * i + 12) 'GRPICONDIRENTRY.nID
                                                              Dim pic As Byte() = GetDataFromResource(hModule, RT_ICON, System.IntPtr.op_Explicit(id))
                                                              'Copy GRPICONDIRENTRY to ICONDIRENTRY.
                                                              dst.Seek(6 + 16 * i, System.IO.SeekOrigin.Begin)

                                                              dst.Write(dir, 6 + 14 * i, 8) 'First 8bytes are identical.
                                                              dst.Write(pic.Length) 'ICONDIRENTRY.dwBytesInRes
                                                              dst.Write(picOffset) 'ICONDIRENTRY.dwImageOffset

                                                              'Copy a picture.
                                                              dst.Seek(picOffset, System.IO.SeekOrigin.Begin)
                                                              dst.Write(pic, 0, pic.Length)

                                                              picOffset = (picOffset + pic.Length)
                                                          Next i
                                                          tmpData.Add(DirectCast(dst.BaseStream, System.IO.MemoryStream).ToArray)
                                                      End Using
                                                      Return True
                                                  End Function

                NativeMethods.EnumResourceNames(hModule, RT_GROUP_ICON, callback, System.IntPtr.Zero)
                Me.iconData = tmpData.ToArray
            Finally
                If (hModule <> System.IntPtr.Zero) Then
                    NativeMethods.FreeLibrary(hModule)
                End If
            End Try
        End Sub

        'Load the binary data from the specified resource.
        Private Function GetDataFromResource(ByVal hModule As System.IntPtr, ByVal type As System.IntPtr, ByVal name As System.IntPtr) As Byte()
            Dim hResInfo As System.IntPtr = NativeMethods.FindResource(hModule, name, type)
            If (hResInfo = System.IntPtr.Zero) Then
                Throw New System.ComponentModel.Win32Exception
            End If
            Dim hResData As System.IntPtr = NativeMethods.LoadResource(hModule, hResInfo)
            If (hResData = System.IntPtr.Zero) Then
                Throw New System.ComponentModel.Win32Exception
            End If
            Dim pResData As System.IntPtr = NativeMethods.LockResource(hResData)
            If (pResData = System.IntPtr.Zero) Then
                Throw New System.ComponentModel.Win32Exception
            End If
            Dim size As UInteger = NativeMethods.SizeofResource(hModule, hResInfo)
            If (size = 0) Then
                Throw New System.ComponentModel.Win32Exception
            End If
            Dim buf As Byte() = New Byte(CInt(size - 1)) {}
            System.Runtime.InteropServices.Marshal.Copy(pResData, buf, 0, buf.Length)
            Return buf
        End Function

        Private Function GetFileName(ByVal hModule As System.IntPtr) As String
            '// Alternative to GetModuleFileName() for the module loaded with
            '// LOAD_LIBRARY_AS_DATAFILE option.

            '// Get the file name in the format Like
            '// "\\Device\\HarddiskVolume2\\Windows\\System32\\shell32.dll"


            Dim fileName As String

            Dim buf As New System.Text.StringBuilder(MAX_PATH)
            Dim len As Integer = NativeMethods.GetMappedFileName(NativeMethods.GetCurrentProcess, hModule, buf, buf.Capacity)
            If len = 0 Then
                Throw New System.ComponentModel.Win32Exception
            End If
            fileName = buf.ToString

            '// Convert the device name to drive name Like
            '// "C:\\Windows\\System32\\shell32.dll"

            For Each c As Char In "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
                Dim drive As String = c & ":"
                buf = New System.Text.StringBuilder(MAX_PATH)
                len = NativeMethods.QueryDosDevice(drive, buf, buf.Capacity)
                If len = 0 Then
                    Continue For
                End If

                Dim devPath As String = buf.ToString
                If fileName.StartsWith(devPath) Then
                    Return (drive & fileName.Substring(devPath.Length))
                End If
            Next

            Return fileName
        End Function

    End Class

End Namespace
