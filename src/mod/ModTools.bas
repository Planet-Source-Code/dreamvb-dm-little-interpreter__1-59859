Attribute VB_Name = "ModTools"
'This module is used to keep any tools we use

Public Sub LoadProcedure(s_ScriptCode As String)
Dim v As Variant, ProcVars As Variant
Dim cLine As String, sTemp As String, sTempB As String, _
ProcNameType As String, sProcName As String, sBuffer As String, _
lVariableName As String, lVariableType As String

Dim e_pos As Integer, OpenB As Integer, CloseB As Integer, bOK As Integer, _
X As Integer, I As Long
Dim Start As Boolean
    ' This function is used to get all the subs and functions,
    ' set's up the variables and get's the code
    v = Split(s_ScriptCode, vbCrLf)
    For I = LBound(v) To UBound(v)
        cLine = Trim(v(I))
        e_pos = InStr(1, cLine, " ", vbTextCompare) ' Find position of write space eg function <whitespace> Main()
        ProcNameType = LCase(Trim(Mid(cLine, 1, e_pos))) ' Get the name of the function or sub
        If ProcNameType = "type" Then cLine = cLine & "()"
        
        If (ProcNameType = "function") Or (ProcNameType = "sub" Or ProcNameType = "type") Then
        If GetProcInfo(cLine, 1) = 0 Then Abort 3, "'(' Expected": Exit For
        If GetProcInfo(cLine, 2) = 0 Then Abort 3, "')' Expected": Exit For
        
        sProcName = GetProcInfo(cLine, 5, e_pos, GetProcInfo(cLine, 1))
            
        If Len(sProcName) = 0 Then
            ' below check that the function or sub has a name
            Abort 3, "Expected name but found '" & ProcNameType & "'"
            Exit For
        End If
            
        If GetProcInfo(cLine, 4) <> 0 Then 'How many parms has the sub or function got
            ProcVars = Split(GetProcInfo(cLine, 3), ",") ' This gets the parms eg "int x, char b"
            For X = LBound(ProcVars) To UBound(ProcVars) ' Start at 0 move up to UBound
                sTempB = Trim(ProcVars(X))
                e_pos = InStr(1, sTempB, " ", vbBinaryCompare) ' check for white space eg int <whitespace> x
                If e_pos <= 0 Then Abort 7: Exit For ' Error
                lVariableName = LTrim(Mid(sTempB, e_pos + 1, Len(sTempB))) ' Get variable name
                lVariableType = LCase(Mid(sTempB, 1, e_pos - 1)) ' Get variable type
                ' Line below adds the new variable and sets it's default data based on it's datatype
                AddVariable lVariableName, lVariableType, False, DefaultVarData(GetVariableTypeFromStr(lVariableType))
            Next
        End If
        
        AddProc sProcName, GetProcInfo(cLine, 4, ","), GetProcInfo("", 6, ProcNameType), GetProcInfo(cLine, 3) ' Add the proc to the stack. I like using that word :)
        Start = True
        End If
        
        If Start Then OpenB = Left(cLine, 1) = "{"
        
        If OpenB Then bOK = 1
        If CloseB Then bOK = 0
        
        If bOK Then sBuffer = sBuffer & LTrim(cLine) & vbCrLf
        CloseB = Left(cLine, 1) = "}"
        
        If CloseB Then
            sBuffer = RemoveVbCrlf(sBuffer) ' Remove vbcrlf
            If Not Left(sBuffer, 1) = "{" Then Abort 1, "{": Exit For
            sBuffer = Right(sBuffer, Len(sBuffer) - 1) ' remove {
            sBuffer = Left(sBuffer, Len(sBuffer) - 1) ' remove }
            sBuffer = RemoveVbCrlf(sBuffer) 'Remove vbcrlf agian
            AddProcCode CLng(ProcCounter), sBuffer ' add proc code
            sBuffer = "" ' clear it
        End If
    Next
    ' Clean up
    Erase v
    If Not IsEmpty(ProcVars) Then Erase ProcVars
    cLine = "": sTemp = "": sTempB = "": ProcNameType = ""
    sProcName = "": lVariableName = "": lVariableType = ""
    e_pos = 0: OpenB = 0: CloseB = 0: bOK = 0: X = 0: I = 0

End Sub

Public Function GetIncludeFiles(lzCode As String) As String
Dim vCode As Variant, I As Long
Dim e_pos As Integer, n_pos As Integer, s_pos As Integer
Dim sline As String, incFile As String
Dim AbsFile As String, sBuff As String, StrA As String
Dim IncludeFound As Boolean
    ' this function is used to get the data from an include file
    incFile = ""
    If Len(lzCode) = 0 Then Exit Function
    vCode = Split(lzCode, vbCrLf)
    For I = 0 To UBound(vCode)
        sline = Trim(vCode(I))
        e_pos = InStr(1, sline, "#", vbBinaryCompare)
        n_pos = InStr(e_pos + 1, sline, " ", vbBinaryCompare)
        
        If (e_pos > 0 And n_pos > 0) Then
            IncludeFound = UCase(Mid(sline, e_pos + 1, n_pos - e_pos - 1)) = "INCLUDE"
        End If
        
        If IncludeFound Then
            n_pos = InStr(1, sline, "<", vbBinaryCompare)
            s_pos = InStr(1, sline, ">", vbBinaryCompare)
            If (n_pos > 0 And s_pos > 0) Then
                incFile = Trim(Mid(sline, n_pos + 1, s_pos - n_pos - 1))
                If InStr(1, incFile, "\") Then
                    AbsFile = incFile
                Else
                    AbsFile = FixPath(App.Path) & "include\" & incFile
                End If
                
                If Not IsFileHere(AbsFile) Then Abort 3, "Include file not found:" & vbCrLf & AbsFile
                sBuff = OpenFile(AbsFile)
                StrA = StrA & sBuff & vbCrLf
                AbsFile = "": e_pos = 0: n_pos = 0: s_pos = 0: incFile = ""
                lzCode = Replace(lzCode, sline, "")
            End If
        End If
    Next
    GetIncludeFiles = StrA
End Function

Public Function OpenFile(lzFile As String) As String
Dim s_Data As String
Dim nFile As Long
    nFile = FreeFile
    ' Open a File
    Open lzFile For Binary As #nFile
        s_Data = Space(LOF(1))
        Get #nFile, , s_Data
    Close #nFile
    
    OpenFile = s_Data
    s_Data = ""
    
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function FixStr(lzBuff As String) As String
    FixStr = Trim(Mid(lzBuff, GetCharPos(lzBuff, "(") + 1, Len(lzBuff) - GetCharPos(lzBuff, "(") - 1))
End Function

Function GetArrayInfo(lzStr As String, ArrayOption As Integer) As Variant
Dim i_pos As Integer, n_pos As Integer, ArrySize As Variant
    
    i_pos = GetCharPos(lzStr, "[")
    n_pos = GetCharPos(lzStr, "]")
    
    If i_pos = 0 Then GetArrayInfo = "": Exit Function
    If n_pos = 0 Then GetArrayInfo = "": Exit Function
    
    If ArrayOption = 1 Then
        GetArrayInfo = LCase(Trim(Mid(lzStr, 1, i_pos - 1))) 'get array name
    ElseIf ArrayOption = 2 Then
        ArrySize = Trim(Mid(lzStr, i_pos + 1, n_pos - i_pos - 1)) ' get array index
        If ArrySize = "" Then GetArrayInfo = -1 Else GetArrayInfo = ArrySize
    End If
    
    i_pos = 0: n_pos = 0: ArrySize = 0
    
End Function

Function HasSqrBrackets(lzStr As String) As Boolean
Dim i_pos As Integer, n_pos As Integer
    i_pos = InStr(1, lzStr, "[", vbBinaryCompare)
    n_pos = InStr(i_pos + 1, lzStr, "]", vbBinaryCompare)
    HasSqrBrackets = (i_pos > 0 And n_pos > 0)
End Function

Public Function GetFromAssign(lzStr As String, AssignPos As Integer, nPostion As Integer) As String
    If AssignPos = 0 Then Exit Function
    If nPostion = 0 Then ' Get the left side of the string
        GetFromAssign = LCase(Trim(Mid(lzStr, 1, AssignPos - 1)))
    Else
        GetFromAssign = Trim(Mid(lzStr, AssignPos + 1, Len(lzStr)))
    End If
End Function

Function GetIndexFromStr(lzStr As String) As String
Dim C As String, StrNum As String
    
    ' This is a small function I made that returns only the number from the left side of a string
    For I = Len(lzStr) To 1 Step -1
        C = Mid(lzStr, I, 1)
        If Not IsNumeric(C) Then
            Exit For
        Else
            StrNum = C & StrNum
        End If
    Next
    
    GetIndexFromStr = StrNum
    StrNum = ""
    C = ""
    
End Function

Function GetCtrlName(lzStr As String) As String
Dim e_pos As Integer
Dim C As String, StrName As String
    ' This function is used to extract a controls name
    For I = Len(lzStr) To 1 Step -1
        C = Mid(lzStr, I, 1)
        If Not IsNumeric(C) Then
            StrName = Mid(lzStr, 1, I)
            Exit For
        End If
    Next
    
    GetCtrlName = StrName
    StrName = ""
    C = ""
End Function

Function RemoveEol(lzLine As String) As String
    If Not Right(lzLine, 1) = ";" Then RemoveEol = lzLine: Exit Function Else RemoveEol = Left(lzLine, Len(lzLine) - 1)
End Function

Public Function TidyLine(lpStrLine As String) As String
Dim StrA As String
    StrA = lpStrLine
    ' strip out any comment that maybe here
    StrA = RemoveSingleLineComment(EolPos(StrA), StrA)
    ' now we can remove the end of line ;
    StrA = RemoveEol(StrA)
    TidyLine = StrA
    StrA = ""
End Function

Function AddQuotes(lzStr As String) As String
Dim a As Boolean, b As Boolean
    a = Left(lzStr, 1) = Chr(34)
    b = Right(lzStr, 1) = Chr(34)
    
    If a = False And b = False Then AddQuotes = Chr(34) & lzStr & Chr(34): Exit Function
    If a And b Then AddQuotes = lzStr: Exit Function
    If a = False Then AddQuotes = Chr(34) & lzStr: Exit Function
    If b = False Then AddQuotes = lzStr & Chr(34): Exit Function
    
End Function

Function DoIF(lpLine As String) As String
Dim a As String, b As String
    IfThenPart = ""
    If GetProcInfo(lpLine, 1) = 0 Then Abort 1, "if '('": Exit Function
    If GetProcInfo(lpLine, 1) = 0 Then Abort 1, "if '('": Exit Function
    a = Trim(GetProcInfo(lpLine, 3))
    b = LCase(Trim(Mid(lpLine, GetProcInfo(lpLine, 2) + 1, Len(lpLine))))
    If Left(b, 4) = "then" Then b = Trim(StrRemoveLeft(b, 4))
    If Left(b, 4) = "goto" Then IfThenPart = b
    DoIF = a
    a = "": b = ""
End Function

Function DoCallProc(sProcline As String) As String
Dim pName As String, vLst As Variant, I As Integer
Dim TheData As String, pParmLst As Variant, lVarName As String
Dim ipos As Integer, nVarType As String
On Error Resume Next
    If LCase(Left(sProcline, 4)) = "call" Then sProcline = Trim(StrRemoveLeft(sProcline, 4))
    
    DoCallProc = vbNullChar

    If GetProcInfo(sProcline, 1) = 0 Then
        Abort 1, "("
        Exit Function
    ElseIf GetProcInfo(sProcline, 2) = 0 Then
        Abort 1, ")"
        Exit Function
    Else
        pName = GetProcInfo(sProcline, 5, 0, GetProcInfo(sProcline, 1))
    End If
    
    If Len(pName) = 0 Then Abort 1, "sub or function name": Exit Function
    If GetProcInfo(sProcline, 4) <> GetProcParmCount(pName) Then Abort 12, "'" & pName & "'" & vbCrLf & "sub or function not defined": Exit Function
    
    vLst = Split(GetProcInfo(sProcline, 3), ",")
    pParmLst = Split(GetProcParmList(pName), ",")
    
    If UBound(vLst) = -1 Then DoCallProc = ProcGetCodeBlock(pName): Exit Function
    If UBound(pParmLst) = -1 Then DoCallProc = ProcGetCodeBlock(pName): Exit Function

    For I = LBound(pParmLst) To UBound(pParmLst)
        TheData = Trim(pParmLst(I))
        ipos = InStr(1, TheData, " ", vbTextCompare)
        nVarType = Mid(TheData, 1, ipos)
        If ipos <> 0 Then lVarName = Trim(Mid(TheData, ipos, Len(TheData)))

        Select Case LCase(Trim(nVarType))
            Case "int":  SetVariable lVarName, CInt(Eval(vLst(I)))
            Case "long": SetVariable lVarName, CLng(Eval(vLst(I)))
            Case "char": SetVariable lVarName, CStr(Eval(vLst(I)))
            Case "bool": SetVariable lVarName, Int(Eval(vLst(I)))
            Case "float": SetVariable lVarName, CSng(Eval(vLst(I)))
        End Select
        
        If Err Then Abort 3, Err.Description: Exit Function
    Next
    
    DoCallProc = ProcGetCodeBlock(pName)
    pName = "": lVarName = "": I = 0
    Erase vLst: Erase pParmLst
    
End Function

Function DoEnum(mLine As String) As Boolean
Dim vEnumList As Variant, sTemp As String
Dim I As Integer, ipos As Integer
    
    sTemp = Trim(mLine)
    ipos = InStr(1, sTemp, "{", vbTextCompare)
    If ipos = 0 Then Abort 11, "expected '{'": Exit Function
    If Not Right(sTemp, 1) = "}" Then Abort 11, "expected '}'": Exit Function
    sTemp = Mid(sTemp, ipos + 1, Len(sTemp) - ipos - 1)
    sTemp = Replace(sTemp, " ", "")
    sTemp = LCase(sTemp)
    ipos = 0
    
    vEnumList = Split(sTemp, ",")
    If UBound(vEnumList) = 0 Then DoEnum = True: Exit Function
    
    For I = LBound(vEnumList) To UBound(vEnumList)
        sTemp = vEnumList(I)
        If VarNotkeyword(sTemp) Then
            Erase vEnumList: sTemp = "": ipos = 0
            Exit For
            Exit Function
        End If
        
        If I = 0 Then
            AddVariable sTemp, "int", True, 0
        Else
            AddVariable sTemp, "int", True, ipos + 1
            ipos = Val(GetVariableData(sTemp))
            sTemp = ""
        End If
    Next
    
    OldIndex = 0: I = 0
    Erase vEnumList
    DoEnum = True
    
End Function

Function GetForLoopData(mLine As String) As Boolean
Dim StrA As String, sTempA As String
Dim ipos As Integer

    GetForLoopData = False
    
    If GetProcInfo(mLine, 1) = 0 Then Abort 1, "(": Exit Function
    If GetProcInfo(mLine, 2) = 0 Then Abort 1, ")": Exit Function
    StrA = GetProcInfo(mLine, 3)
    
    If Len(StrA) = 0 Then Abort 7: Exit Function
    StrA = Replace(StrA, " ", "")
    ipos = GetCharPos(StrA, ";")
    If ipos = 0 Then Abort 1, ";": Exit Function
    sTempA = Mid(StrA, 1, ipos - 1)
    ForLoopExpr = Mid(StrA, ipos + 1, Len(StrA))
    If Len(ForLoopExpr) = 0 Then Abort 1, "To": Exit Function
    ipos = GetCharPos(StrA, "=")
    If ipos = 0 Then Abort 1, "=": Exit Function
    ForLoopVarName = Mid(sTempA, 1, ipos - 1)
    If VariableIndex(ForLoopVarName) = -1 Then Abort 7: Exit Function
    sTempA = ReturnData(Mid(sTempA, ipos + 1, Len(sTempA)))
    If Len(sTempA) = 0 Then Abort 3, "Type mistake": Exit Function
    ForLoopStart = CLng(sTempA)
    SetVariable ForLoopVarName, Val(sTempA)
    sTempA = "": StrA = "": ipos = 0
    GetForLoopData = True
    
End Function

Public Function EolPos(lzLine As String) As Integer
' Returns the position of the end of line ;
    EolPos = InStrRev(lzLine, ";", Len(lzLine), vbBinaryCompare)
End Function

Function CountIF(lzExpr As String, nChar As String)
Dim X As Integer, iCount As Integer
Dim sByte() As Byte
    sByte = lzExpr
    ' Just used to count how many nChar are in lzExpr
    For X = LBound(sByte) To UBound(sByte)
        If sByte(X) = Asc(nChar) Then iCount = iCount + 1
    Next
    X = 0: Erase sByte
    CountIF = iCount
    iCount = 0
End Function

Function RemoveVbCrlf(StrA As String) As String
Dim StrB As String
    ' Removes vbCrLf from left and right side of a string.
    StrB = StrA
    If Right(StrB, 2) = vbCrLf Then StrB = Left(StrB, Len(StrB) - 2)
    If Left(StrB, 2) = vbCrLf Then StrB = Right(StrB, Len(StrB) - 2)
    RemoveVbCrlf = StrB
    StrB = ""
End Function

Public Function GetCharPos(lpStr As String, lpChar As String) As Integer
    GetCharPos = InStr(1, lpStr, lpChar, vbBinaryCompare)
End Function

Function StrRemoveLeft(lzStr As String, iPlaces As Integer) As String
    StrRemoveLeft = Right(lzStr, Len(lzStr) - iPlaces)
    ' removes the left side of a string a number of places
End Function

Function StrRemoveRight(lzStr As String, iPlaces As Integer) As String
    StrRemoveRight = Left(lzStr, Len(lzStr) - iPlaces)
    ' removes the right side of a string a number of places
End Function

Public Function GetProcInfo(lzStr As String, mOption As Integer, Optional Data1 As Variant, Optional Data2 As Variant) As Variant
' used for subs and function preforms number of different actions see below:

    If mOption = 1 Then ' see if open barcket ( is found
        GetProcInfo = InStr(1, lzStr, "(", vbTextCompare)
        Exit Function
    End If
    
    If mOption = 2 Then ' see if close barcket ) is found
        GetProcInfo = InStr(1, lzStr, ")", vbTextCompare)
        Exit Function
    End If
    
    If mOption = 3 Then ' strip the data between the backets eg (int x,int y) returns int x, int y
        If CBool(GetProcInfo(lzStr, 1) And CBool(GetProcInfo(lzStr, 2))) Then
            GetProcInfo = Trim(Mid(lzStr, GetProcInfo(lzStr, 1) + 1, GetProcInfo(lzStr, 2) - GetProcInfo(lzStr, 1) - 1))
        End If
    End If
    
    If mOption = 4 Then ' Counts the number of parms for the proc
        If Len(GetProcInfo(lzStr, 3)) = 0 Then GetProcInfo = 0: Exit Function
        GetProcInfo = CountIF(lzStr, ",") + 1
        Exit Function
    End If
    
    If mOption = 5 Then ' returns the proc name
      GetProcInfo = Trim(Mid(lzStr, Data1 + 1, Data2 - Data1 - 1))
      Exit Function
    End If
    
    If mOption = 6 Then GetProcInfo = LCase(Data1) = "function" ' returns a bool value if function is found true is returned else false is retured
    
End Function

Public Function Encrypt(lzStr As String) As String
Dim sByte() As Byte, I As Integer
On Error Resume Next
    ' Function to encrypt / decrypt a string
    sByte() = StrConv(lzStr, vbFromUnicode)
    
    For I = LBound(sByte) To UBound(sByte)
        sByte(I) = sByte(I) Xor 94
    Next
    
    Encrypt = StrConv(sByte, vbUnicode)
    I = 0
    Erase sByte()
    
End Function
