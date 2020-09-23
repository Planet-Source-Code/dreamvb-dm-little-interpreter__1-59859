Dim sCmd, WshShell, fso, Ans, Installed
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = WScript.CreateObject("Scripting.FileSystemObject")

Function GetPath()
    Dim path
    path = WScript.ScriptFullName
    GetPath = Left(path, InStrRev(path, "\"))
End Function

Sub main()
Dim Ans

  Ans = MsgBox("Welcome to the Little# Interpreter" _
  & vbCrLf & "This programs will need to do some things in order to get the program running" _
  & vbCrLf & vbCrLf & "1. Install the script extensions" _
  & vbCrLf & "2. Register Little# Interpreter components" _
  & vbCrLf & vbCrLf & "Do you want to proceed with the install?", vbYesNo Or vbQuestion, "Alpha v1.1")
  
  If Ans = vbNo Then Exit Sub
  ' install the script file extensions
  WshShell.RegWrite "HKCR\.lsi\", "lsi.Interpreter"
  WshShell.RegWrite "HKCR\.lsi\Content Type", "text/plain"
  WshShell.RegWrite "HKCR\lsi.Interpreter\", "Little# Interpreter script"
  WshShell.RegWrite "HKCR\lsi.Interpreter\DefaultIcon\", GetPath & "bscript.exe,1"
  WshShell.RegWrite "HKCR\lsi.Interpreter\Shell\Open\Command\", GetPath & "bscript.exe %1 /Run"
  WshShell.RegWrite "HKCR\lsi.Interpreter\Shell\Make\Command\", GetPath & "dmcomp.exe %1 /make"
  WshShell.RegWrite "HKCR\lsi.Interpreter\Shell\Edit\Command\", "notepad.exe %1"
  WshShell.RegWrite "HKCR\lsi.Interpreter\Shell\Print\Command\", "notepad.exe /p %1"
  'Register the system components
  
   sCmd = GetPath & "system\DmScriptLib.dll"
   If Not fso.FileExists(sCmd) Then
       MsgBox "Unable to register" & vbCrLf & sCmd, vbCritical, "error"
       Exit Sub
   Else
        WshShell.Run "regsvr32.exe " & sCmd & " /s"
   End If
   ' next register the test dll
   sCmd = GetPath & "system\myDll.dll"
   If Not fso.FileExists(sCmd) Then
       MsgBox "Unable to register" & vbCrLf & sCmd, vbCritical, "error"
       Exit Sub
   Else
        WshShell.Run "regsvr32.exe " & sCmd & " /s"
   End If
   'Clear up
   sCmd = ""
   Set WshShell = Nothing
   Set fso = Nothing
   
End Sub


Call main



