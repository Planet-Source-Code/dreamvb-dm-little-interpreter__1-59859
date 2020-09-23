Attribute VB_Name = "ModStartUp"
Sub AddCode(nCode As String)
Dim s_main As String, dm_ScriptSrc As String, IncStr As String
Dim sTemp As String

    If frmgui.Visible Then Unload frmgui
    
    dm_ScriptSrc = nCode ' Get the main script code
    dm_ScriptSrc = Replace(dm_ScriptSrc, " _" & vbCrLf, " ", , , vbTextCompare)
    
    GlobalReset ' See ModPublic GlobalReset
    AddGlobalVars ' Load the global variables for this program
    IncStr = GetIncludeFiles(dm_ScriptSrc) 'Get the include file data
    sTemp = dm_ScriptSrc
    dm_ScriptSrc = IncStr & sTemp
    sTemp = "" 'Clean up
    LoadProcedure dm_ScriptSrc 'Load all the Procedures
    dm_ScriptSrc = "" 'Clean
    
    If ProcIndex("main") = -1 Then ' check that we have the main Procedure in the code
        Abort 9
    Else
        s_main = ProcGetCodeBlock("main") ' get main procedure code
        Execute s_main, "main" ' execute the code found above
        s_main = ""
    End If

    If LinkError Then
        GlobalReset
        Unload frmgui
        Unload Form1
        End
    End If
    
End Sub

Function DecodeAndExecute() As String
Dim OffSet As Long
Dim StrBuff As String, StrHead As String, ThisFile As String
Dim StrTmpBuff As String

    ThisFile = FixPath(App.Path) & App.EXEName & ".exe"
    StrHead = Chr(5) & Chr(255) & "DM#"
    nFile = FreeFile
    If IsFileHere(ThisFile) = False Then End
    
    StrBuff = OpenFile(ThisFile)
    OffSet = InStr(1, StrBuff, StrHead, vbTextCompare)
    
    If OffSet = 0 Then ' can;t locate offset so exit
        MsgBox "Unable to locate vaild offset", vbInformation, "error"
        ThisFile = ""
        StrHead = ""
        StrBuff = ""
        End
    Else
        StrTmpBuff = Mid(StrBuff, OffSet + Len(StrHead), Len(StrBuff))
        StrBuff = "": OffSet = 0: StrHead = ""
        DecodeAndExecute = Encrypt(StrTmpBuff)
        ' Run the code
        StrTmpBuff = ""
    End If
End Function

Sub Main()
Dim i_pos As Integer
Const ErrMsg As String = "Incorrect command line parameters."
Dim lpCommandLine As String
    lpCommandLine = Trim(Command) ' get the command line arags
    
    If Len(lpCommandLine) = 0 Then
        ' ok this is a script file been assigned
        ' check for the switch
        AddCode DecodeAndExecute 'Add the code to be executed
    Else
        i_pos = InStrRev(lpCommandLine, "/", Len(lpCommandLine), vbBinaryCompare)
        If i_pos = 0 Then
            MsgBox ErrMsg, vbCritical, "error"
            End
        ElseIf Not LCase(Mid(lpCommandLine, i_pos, Len(lpCommandLine))) = "/run" Then
            MsgBox ErrMsg, vbCritical, "error"
            End
        Else
            ' fix the filename by removeing the switch
            lpCommandLine = RTrim(Mid(lpCommandLine, 1, i_pos - 1))
            ' Ok from here we can now execute the script
            
            AddCode OpenFile(lpCommandLine) 'Add the code to be executed
            lpCommandLine = ""
            i_pos = 0
            End If
    End If

End Sub
