Attribute VB_Name = "Modexecute"
Function FixLine(lzStrLineBuff As String, RemovePlaces As Integer) As String
' ok all this does is clean up a line. I was finding I was using the same
' code over and over agian so I left it in here to make things a little easyer
Dim StrAB As String
    StrAB = lzStrLineBuff
    
    StrAB = Trim(StrRemoveLeft(StrAB, RemovePlaces))
    FixLine = Trim(TidyLine(StrAB))
    StrAB = ""
End Function

Public Function Execute(sExecuteCode As String, Optional t_ProcName As String)
Dim e_pos As Integer, n_pos As Integer, h_pos As Integer, i_pos As Integer, I As Long, vLine As Variant
Dim lVariableName As String, lVariableType As String, lVariableData As String _
, sTemp1 As String, sTemp2 As String, TheToken As String, sline As String, _
nForln As Long, nLoop As Long, nLoopExpr As String, nArrIndex As Integer
Dim Temp3 As String, IfSkip As Boolean, skip As Boolean
Dim lVarNameLst As Variant, TempI As Integer, TempJ As Integer
On Error Resume Next

    ' main function for executeing the script. NOTE this code is quite messy
    
    vLine = Split(sExecuteCode, vbCrLf)
    For I = LBound(vLine) To UBound(vLine)
        If LinkError And GuiError Then MsgBox Link_Error_Str: Exit For ' We must stop if LinkError = True
        
        sline = Trim(vLine(I))
        If Err Then sline = ""
        
        If LCase(sline) = "doevents;" Then DoEvents
        
        If m_GotoLabel <> "" Then
            ' check for a goto label
            If m_GotoLabel = LCase(sline) Then m_GotoLabel = ""
            GoTo ThisBlock
        End If
        
        If m_SwitchLabel_A <> "" Then
            'check for a switch label
            TempI = LCase(Mid(sline, 1, 5)) = "case:"
            If TempI <> 0 Then
                Temp3 = sline
                Temp3 = Trim(StrRemoveLeft(Temp3, 5))
                Temp3 = Eval(TidyLine(Temp3))
            End If
            
            If LCase(Temp3) = LCase(m_SwitchLabel_A) Then
                m_SwitchLabel_A = ""
                Temp3 = ""
            End If
            GoTo ThisBlock
        End If
        
        If LCase(sline) = "next;" And nForln > 0 Then
            If Eval(ForLoopExpr) = 0 Then
                I = Val(nForln)
            End If
        End If
        
        If LCase(sline) = "loop;" Then
            If Len(nLoopExpr) = 0 Then Abort 3, "Loop; found but expected While"
            If Eval(nLoopExpr) = 0 Then
                I = nLoop
            Else
                nLoopExpr = ""
                sline = ""
            End If
        End If
        
        If skip = True Then
            If LCase(sline) = "end if" Or LCase(sline) = "endif" Then
                skip = False
            End If
            GoTo ThisBlock
        End If
        
        If IfSkip = True Then
            Select Case LCase(sline)
                Case "end if", "endif"
                    IfSkip = False
                Case "else"
                    IfSkip = False
            End Select
            GoTo ThisBlock
        Else
            If LCase(sline) = "else" Then
                skip = True
                GoTo ThisBlock
            End If
        End If
        
        e_pos = InStr(1, sline, " ", vbTextCompare)
        h_pos = InStr(1, sline, ".", vbTextCompare)
 
        If e_pos <> 0 Then
   
            TheToken = LCase(Trim(Left(sline, e_pos)))
  
            Select Case TheToken
                Case "redim"
                    sline = FixLine(sline, 5)
                    If HasSqrBrackets(sline) Then
                        sTemp1 = Eval(GetArrayInfo(sline, 2)) 'get ressize value
                        sline = GetArrayInfo(sline, 1) 'get array name
                        If Not isArrayEx(sline) Then Abort 7: Exit Function
                        'resize the array
                        ResizeArray ArrayIndex(sline), CInt(sTemp1)
                    End If
                Case "addcontrol"
                    sline = FixLine(sline, 10)
                    If Len(sline) = 0 Then Abort 1, "Control Name": Exit For
                    frmgui.GUIAddControl sline
                Case "return"
                    sline = FixLine(sline, 6)
                    sReturnFormFunc = Eval(sline)
                    sline = ""
                Case "copyfile"
                    sline = FixLine(sline, 8)
                    dmFunction1 sline, dmFileCopy, 1, "copyfile"
                    sline = ""
                Case "const"
                    If EolPos(sline) = 0 Then Abort 5: Exit For
                    sline = FixLine(sline, 5)
                    TempI = GetCharPos(sline, "=")
                    If TempI = 0 Then Abort 1, "Const assignment '='": Exit For
                    lVariableName = Trim(Mid(sline, 1, TempI - 1))
                    lVariableData = Eval(Trim(Mid(sline, TempI + 1, Len(sline))))
                    If VarNotkeyword(lVariableName) Then Abort 8: Exit For
                    AddVariable lVariableName, "char", True, lVariableData
                    lVariableName = "": lVariableData = "": TempI = 0: sline = ""
                Case "call" ' call sub or function
                    If EolPos(sline) = 0 Then Abort 5: Exit For
                    Temp3 = DoCallProc(sline)
                    If (Temp3 <> vbNullChar) Then Execute Temp3: Temp3 = ""
                Case "destroy"
                    sline = FixLine(sline, 7)
                    DestroyArray ArrayIndex(sline)
                Case "echo"
                    If EolPos(sline) = 0 Then Abort 5: Exit For
                    sline = FixLine(sline, 4)
                    Echo sline
                Case "deletefile"
                    sline = FixLine(sline, 10)
                    DelFile sline
                    sline = ""
                 Case "mkdir"
                    sline = FixLine(sline, 5)
                    
                Case "int", "char", "bool", "float", "long" ' we found a variable
                    If EolPos(sline) = 0 Then Abort 5: Exit For
                    sline = TidyLine(sline)
                    lVariableType = Trim(Mid(sline, 1, e_pos - 1)) ' Extract and store the variables type will be char,int or bool
                    lVariableName = LCase(Trim(Mid(sline, e_pos + 1, Len(sline)))) ' extract and store the varibales name
                    lVarNameLst = Split(lVariableName, ",")
                    
                    ' used to see if we got more than one var on a line eg int x,y,z
                    For TempI = LBound(lVarNameLst) To UBound(lVarNameLst)
                         'checks if the variable name is vaild
                        lVariableName = Trim(lVarNameLst(TempI))
                        
                        TempJ = GetCharPos(lVariableName, "=")
                        
                        If TempJ Then
                            lVariableData = Trim(Mid(lVariableName, TempJ + 1, Len(lVariableName)))
                            lVariableName = RTrim(Mid(lVariableName, 1, TempJ - 1)) ' get variable name
                            If VarNotkeyword(lVariableName) Then Abort 8: I = UBound(vLine)
                            AddVariable lVariableName, lVariableType, False, SetVarDataX(GetVariableTypeFromStr(lVariableType), lVariableData)
                            sline = "" ' Clear after we finsihed or you will get an error
                        ElseIf HasSqrBrackets(lVariableName) Then ' check for an array
                           AddArray GetArrayInfo(lVariableName, 1), lVariableType, Val(GetArrayInfo(lVariableName, 2))
                        Else
                            If VarNotkeyword(lVariableName) Then Abort 8: Exit For
                            ' ok so we got the variables type and name ,now we add the variable and set the default data for it
                            AddVariable lVariableName, lVariableType, False, DefaultVarData(GetVariableTypeFromStr(lVariableType))
                        End If
                        
                    Next
                    lVariableData = "": lVariableName = "": lVariableType = ""
                    TempI = 0: TempJ = 0
                    
                Case "enum"
                    If EolPos(sline) = 0 Then Abort 5: Exit For
                    sline = FixLine(sline, 4)
                    If Not DoEnum(sline) Then Exit For
                Case "goto"
                    m_GotoLabel = LCase(Trim(StrRemoveLeft(sline, 4)))
                Case "switch"
                    m_SwitchLabel_A = Eval(LCase(Trim(StrRemoveLeft(sline, 6)))) ' Switch
                Case "if"
                    sline = StrRemoveLeft(sline, 3)
                    DoIF sline
                    If Left(IfThenPart, 4) <> "goto" Then
                        If Not Eval(DoIF(sline)) = 1 Then IfSkip = True
                    Else
                        Execute IfThenPart
                        sline = ""
                    End If
                    
                Case "for"
                    sline = StrRemoveLeft(sline, 3)
                    If Not GetForLoopData(sline) Then Exit For
                    nForln = I
                    sline = ""
                Case "while"
                    nLoopExpr = Trim(StrRemoveLeft(sline, 5))
                    If GetProcInfo(nLoopExpr, 1) = 0 Then Abort 1, "while loop '('": Exit For
                    If GetProcInfo(nLoopExpr, 2) = 0 Then Abort 1, "')' while loop": Exit For
                    nLoopExpr = GetProcInfo(nLoopExpr, 3)
                    nLoop = I
                Case "set"
                    If CreateComObject(StrRemoveLeft(TidyLine(sline), 4)) = 0 Then Exit For
                    sline = ""
                Case "let"
                    sline = FixLine(sline, 3)
            End Select
            
            n_pos = InStr(1, sline, "=", vbBinaryCompare) ' Find Assign posistion
            
            If n_pos <> 0 Then ' We found an assign
                 sline = TidyLine(sline) ' Tidy the line
                lVariableName = LCase(Trim(Left(sline, n_pos - 1)))
                
                If HasSqrBrackets(lVariableName) Then
                    lVariableName = GetArrayInfo(lVariableName, 1)
                End If
                
                If isArrayEx(lVariableName) Then
                    On Error Resume Next
                    lVariableName = LCase(lVariableName)
                    lVariableName = Trim(lVariableName)
                    sTemp1 = Eval(Trim(Mid(sline, n_pos + 1, Len(sline)))) 'Array assigned data
                    nArrIndex = Eval(GetArrayInfo(sline, 2)) ' Get the array index
                    SetArrayData ArrayIndex(lVariableName), nArrIndex, sTemp1
                    sTemp1 = "": lVariableName = "": nArrIndex = 0
                Else
                    If Not VariableIndex(lVariableName) = -1 Then
          
                        If GetCharPos(sline, ".") <> 0 Then
                            ' ok it looks like it may be an assign to a control
                            sTemp1 = frmgui.ControlPropRead(GetControlPropInfo(sline, 1), GetControlPropInfo(sline, 2))
                            If IsNumeric(sTemp1) Then
                                sline = lVariableName & " = " & AddQuotes(sTemp1)
                            Else
                                sline = lVariableName & " = " & sTemp1
                            End If
                        End If
                        
                        sTemp1 = Eval(Trim(Mid(sline, n_pos + 1, Len(sline)))) ' Get the data been assigned and eval it
                        SetVariable lVariableName, SetVarDataX(GetVariableType(lVariableName), sTemp1)  ' Add the variable and it's assigned data
                        sTemp1 = "5"
                    End If
                End If
            End If
        End If
        
        If h_pos <> 0 Then
            sTemp1 = LCase(Trim(Mid(sline, 1, h_pos - 1)))
            
            If Len(GetIndexFromStr(sTemp1)) > 0 Then
                ' Found a control property
                sline = StrRemoveLeft(sline, Len(sTemp1) + 1)
                
                If GetCharPos(sTemp1, ".") = 0 Then
                    frmgui.DoControlProp sTemp1, sline
                    sline = "": sTemp1 = ""
                End If
            End If
    
            Select Case LCase(Trim(Mid(sline, 1, h_pos - 1)))
                Case "dialog"
                    sline = Trim(StrRemoveLeft(sline, Len(sTemp1) + 1))
                    sline = TidyLine(sline)
                    
                    If GetCharPos(sline, "(") <> 0 Then
                        Select Case UCase(Mid(sline, 1, GetCharPos(sline, "(") - 1))
                            Case "PRINT":
                                If Not Right(sline, 1) = ")" Then Abort 1, ")": Exit For
                                frmgui.Print Eval(FixStr(sline))
                                sline = ""
                            Case "PSET"
                                frmgui.SetPixel GetProcInfo(sline, 3)
                            Case "LINE"
                                sline = TidyLine(sline)
                                frmgui.DrawLine StrRemoveLeft(sline, 4)
                            Case "CLOSE": GlobalReset: Unload frmgui
                        End Select
                    Else
                        ' Call the forms property
                        frmgui.DoControlProp "dialog", sline
                        sline = ""
                    End If
            End Select
            End If
            
       Dim StrLineA As String

        If GetCharPos(sline, ";") <> 0 Then
            StrLineA = LCase(LTrim(Mid(sline, 1, EolPos(sline) - 1)))
        End If
       
        Select Case StrLineA
        
            Case "break": Exit For
            Case "beep": Beep
            Case "close": End
            Case Else
                If Right(StrLineA, 2) = "++" Then ' increment counter
                    ' extract the variable name
                    lVariableName = Mid(StrLineA, 1, GetCharPos(StrLineA, "+") - 1)
                    If VariableIndex(lVariableName) = -1 Then Abort 7: Exit For
                    ' Get the current variables data end increment by 1
                    SetVariable lVariableName, GetVariableData(lVariableName) + 1
                    StrLineA = ""
                    ' store the new data back to the variable
                ElseIf Right(StrLineA, 2) = "--" Then
                    ' extract the variable name
                    lVariableName = Mid(StrLineA, 1, GetCharPos(StrLineA, "-") - 1)
                    If VariableIndex(lVariableName) = -1 Then Abort 7: Exit For
                    ' Get the current variables data end increment by 1
                    SetVariable lVariableName, GetVariableData(lVariableName) - 1
                    StrLineA = ""
                    ' store the new data back to the variable
                End If
        End Select
ThisBlock:
    Next I
    
'Clean up
lVariableName = "": lVariableName = "": lVariableType = "": lVariableData = ""
sTemp1 = "": sTemp2 = "": TheToken = "": sline = ""
nLoopExpr = ""

e_pos = 0: n_pos = 0: h_pos = 0: i_pos = 0: I = 0: nForln = 0: nLoop = 0
Erase vLine: TempI = 0: TempJ = 0
'If Err Then Err.Clear

End Function


