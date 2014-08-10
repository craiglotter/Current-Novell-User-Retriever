Imports Microsoft.Win32
Imports System.IO
Imports System.Reflection

Module Application_Code

    Sub Main()
        Try

        
        'HKEY_CURRENT_USER\Volatile Environment
        Dim str As String
        Dim keyflag1 As Boolean = False
        Dim oReg As RegistryKey = Registry.CurrentUser

        Dim oKey As RegistryKey
        oKey = oReg
        Dim subs() As String = ("Volatile Environment").Split("\")
        For Each stri As String In subs
            oKey = oKey.OpenSubKey(stri, True)
        Next

        If Not oKey Is Nothing Then
            str = oKey.GetValue("NWUSERNAME")
            If Not IsNothing(str) And Not (str = "") Then
                Console.WriteLine(str)
            Else
                Console.WriteLine("NWUSERNAME not found")
            End If

            str = oKey.GetValue("HOME_DIRECTORY")
            If Not IsNothing(str) And Not (str = "") Then
                Console.WriteLine(str)
            Else
                Console.WriteLine("HOME_DIRECTORY not found")
            End If

        End If
        oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Console.WriteLine("Fail. Check Error Log for more details.")
            Error_Handler(ex, "Main Code")
        End Try
    End Sub

    Private Function ApplicationPath() As String
        Return _
        Path.GetDirectoryName([Assembly].GetEntryAssembly().Location)
    End Function

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            Dim dir As DirectoryInfo = New DirectoryInfo((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((ApplicationPath() & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_TFSR_Error_Log.txt", True)

            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ":" & ex.ToString)


            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            Console.WriteLine("An error occurred in Current Novell User Retriever's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

End Module
