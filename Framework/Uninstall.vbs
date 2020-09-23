Dim sCmd, WshShell, fso, Ans, Installed
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Sub main()
Dim Ans

  Ans = MsgBox("Welcome to the Little# Interpreter Uninstall program" _
  & vbCrLf & "This programs will remove and unregister the compoenets for Little# Interpreter" _
  & vbCrLf & vbCrLf & "Do you want to proceed with the uninstall?", vbYesNo Or vbQuestion, "Alpha v1.1")
  
  If Ans = vbNo Then Exit Sub
  ' uninstall the script file extensions
    WshShell.RegDelete "HKCR\.lsi\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\DefaultIcon\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Edit\Command\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Edit\"

    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Make\Command\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Make\"

    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Open\Command\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Open\"
   
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Print\Command\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\Print\"
 
    WshShell.RegDelete "HKCR\lsi.Interpreter\Shell\"
    WshShell.RegDelete "HKCR\lsi.Interpreter\"

    WshShell.run "regsvr32.exe system\DmScriptLib.dll /u /s"
    WshShell.run "regsvr32.exe system\myDll.dll /u /s"

   'Clear up
   sCmd = ""
   Set WshShell = Nothing
   
End Sub

Call main
