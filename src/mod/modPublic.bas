Attribute VB_Name = "modPublic"
' Public error vairables
Public LinkError As Boolean, GuiError As Boolean
Public Link_Error_Str As String

Public m_FilePtr As Long ' hold  the file number
Public sCommandLine As String ' Hold the command line args
Public m_GotoLabel As String, m_SwitchLabel_A As Variant ' used for goto and switchs
Public ForLoopVarName As String, ForLoopStart As Long, _
ForLoopExpr As String, IfThenPart As String 'Loop variables

Public Const KeyWords As String = "echo,inputbox,if,then,else,endif" & _
",goto,for,next,exit,return,call,beep" ' keywords used in this program

Public sReturnFormFunc As Variant

Function RemoveSingleLineComment(StartPos As Long, StrLine As String)
Dim e_pos As Long
    ' used to strip away comments from a single line
    e_pos = InStr(StartPos + 2, StrLine, "//", vbTextCompare)
    If e_pos <> 0 Then ' Yes we found it so strip away comments and return
        RemoveSingleLineComment = Mid(StrLine, 1, e_pos - 2)
    Else ' Ok we found nothing so just return the line as it came in
        RemoveSingleLineComment = StrLine
    End If
    e_pos = 0 ' reset
End Function

Public Sub Abort(lMsg As Integer, Optional Info As String)
    LinkError = True
    ' This sub is used to show errors the user may have incountered.
    Select Case lMsg
        Case 0
            Link_Error_Str = "'" & Info & "' : undeclared identifier"
        Case 1
            Link_Error_Str = Info & " : Expected"
        Case 2
            Link_Error_Str = "cannot convert from '" & Info & "' to 'int'"
        Case 3
            Link_Error_Str = Info
        Case 4
            Link_Error_Str = Info & " overflow"
        Case 5
            Link_Error_Str = "';' End of line expected " & Info
        Case 6
            Link_Error_Str = "Ambiguos name found: '" & Info & "'"
        Case 7
            Link_Error_Str = "Expected: variable identifier"
        Case 8
            Link_Error_Str = "Keywords cannot be used as variable names"
        Case 9
            Link_Error_Str = "Expected startup main()"
        Case 10
            Link_Error_Str = "Cannot assign to readonly consts"
        Case 11
            Link_Error_Str = "Enum syntax error " & Info
        Case "12"
            Link_Error_Str = "Argument not optional " & Info
        Case "13"
            Link_Error_Str = "sub or function not defined " & Info
    End Select
    
    If LinkError And Len(Link_Error_Str) > 0 Then
        MsgBox "Compile error " & vbCrLf & Link_Error_Str, vbExclamation
        GlobalReset
        Unload frmgui
        Unload Form1
        DoEvents
    End If
    
End Sub

Public Function ReturnData(LookFor As String) As Variant
Dim iTmp As Variant, lTmp As Integer
Dim StrTmp As String

    lTmp = -1
    iTmp = "": iTmp = Trim(LookFor)
    StrTmp = ""
    
    If Left(iTmp, 1) = "$" Then ' we found a hex number
        iTmp = StrRemoveLeft(CVar(iTmp), 1)
        LookFor = Val("&H" & iTmp)
    End If
    
    lTmp = ArrayIndex(GetArrayInfo(LookFor, 1))
    
    If HasSqrBrackets(LookFor) And lTmp <> -1 Then
        ' an array has been found so return the data
        StrTmp = Eval(GetArrayData(lTmp, Eval(GetArrayInfo(LookFor, 2))))
        
        If IsNumeric(StrTmp) Then
            LookFor = StrTmp
        Else
            LookFor = StrTmp
        End If
    End If
    
    If VariableIndex(LookFor) <> -1 Then
        'variable found get the variables data
        ReturnData = GetVariableData(LookFor)
        Exit Function
    Else
        ReturnData = LookFor
        Exit Function
    End If

End Function


Public Sub AddGlobalVars()
    ' string consts
    AddVariable "true", "bool", True, -1
    AddVariable "false", "bool", True, 0
    AddVariable "dmcrlf", "char", True, vbCrLf
    AddVariable "dmtab", "char", True, vbTab
    AddVariable "dmlf", "char", True, vbLf
    AddVariable "dmcr", "char", True, vbCr
    AddVariable "dmspace", "char", True, Chr(32)
    AddVariable "nill", "int", True, 0
    'date and time
    AddVariable "date", "char", False, Date
    AddVariable "time", "char", False, Time
    AddVariable "timer", "float", True, ""
    ' file read and write consts
    AddVariable "fread", "int", True, 1
    AddVariable "fwrite", "int", True, 2
    AddVariable "fappend", "int", True, 3
    AddVariable "foutput", "int", True, 4
    
    ' other consts
    AddVariable "cmd$", "char", True, sCommandLine
    AddVariable "pi", "float", True, 27 / 7
    
    'color consts
    AddVariable "dmblack", "int", True, vbBlack
    AddVariable "dmblue", "int", True, vbBlue
    AddVariable "dmCyan", "int", True, vbCyan
    AddVariable "dmgreen", "int", True, vbGreen
    AddVariable "dmmagenta", "int", True, vbMagenta
    AddVariable "dmred", "int", True, vbRed
    AddVariable "dmwhite", "int", True, vbWhite
    AddVariable "dmyellow", "int", True, vbYellow
    'Messagebox consts
    AddVariable "mb_information", "int", True, vbInformation
    AddVariable "mb_Critical", "int", True, vbCritical
    AddVariable "mb_Exclamation", "int", True, vbExclamation
    AddVariable "mb_OKCancel", "int", True, vbOKCancel
    AddVariable "mb_OKOnly", "int", True, vbOKOnly
    AddVariable "mb_Question", "int", True, vbQuestion
    AddVariable "mb_YesNo", "int", True, vbYesNo
    AddVariable "mb_yes", "int", True, 6
    AddVariable "mb_no", "int", True, 7
    ' COM Call Type methods
    AddVariable "dmLet", "int", True, 4
    AddVariable "dmGet", "int", True, 2
    AddVariable "dmSet", "int", True, 8
    AddVariable "dmMethod", "int", True, 1

End Sub

Public Sub GlobalReset()
    frmgui.UnloadControls ' Unload Controls
    ResetProcStack ' Erase proc Stack
    ResetVarStack ' Erase all the variables contents
    ResetObjStack ' Erase all the object collections
    
    sReturnFormFunc = vbNullChar
    Link_Error_Str = "" ' Clear error message
    m_GotoLabel = ""
    m_SwitchLabel_A = ""
    IfThenPart = ""
    ForLoopStart = 0
    ForLoopVarName = ""
    ForLoopExpr = ""
    ' Turn off error flags
    GuiError = False
    LinkError = False
End Sub
