Module Module1

    Sub Main()
        Dim procx As New Process()
        With procx
            'change the "/C to whatever command you like..."
            .StartInfo = New ProcessStartInfo("cmd.exe", "/C type c:\the\file\source.txt > c:\the\file\destination.txt")
            With .StartInfo
                .CreateNoWindow = True
                .UseShellExecute = False
                .RedirectStandardOutput = True
            End With
            .Start()
            .WaitForExit()
        End With
    End Sub

End Module
