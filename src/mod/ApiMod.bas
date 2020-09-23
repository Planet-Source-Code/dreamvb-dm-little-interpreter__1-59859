Attribute VB_Name = "ApiMod"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

' Ok I know there not a lot in here yet I add some stuff next time.
' What I really want is a way to call API from my code so I will
' not have to hardcode it in. I been looking in C++ so I may next time include a
' dll to do this
