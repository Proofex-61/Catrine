Set objShell = CreateObject("Shell.Application")
For Each wnd In objShell.Windows
    If InStr(wnd.FullName, "explorer.exe") > 0 Then
        wnd.Quit
    End If
Next