Attribute VB_Name = "modEval"
' Some of this code was found on the net can't remmber were
' so thanks to the person that did submit it
' I had to chnage some of the code to work work with my program
' also if anyones knows were I can find a good example of RPN please let me know.

Function isOperator(StrExp As String) As Boolean
    isOp = False
    If StrExp = "+" Or StrExp = "-" Or StrExp = "*" Or StrExp = "\" Or StrExp = "/" _
    Or StrExp = "&" Or StrExp = "^" Or StrExp = "=" Or StrExp = "<" Or StrExp = ">" _
    Or StrExp = "%" Then isOperator = True
End Function

Public Function Eval(Expression)
Dim iCounter As Integer, sOperator As String, Value As Variant
Dim iTmp As Variant, StrCh As String

    iCounter = 1
    
    Do While iCounter <= Len(Expression)
        StrCh = Mid(Expression, iCounter, 1)
        
        If isOperator(StrCh) Then
            sOperator = StrCh
            iCounter = iCounter + 1
        End If

        Select Case sOperator
            Case "": Value = Token(Expression, iCounter)
            Case "^": Value = Value ^ Token(Expression, iCounter)
            Case "%": Value = Value Mod Token(Expression, iCounter)
            Case "<"
                Value = Value < Token(Expression, iCounter)
                Value = Abs(Value)
                Value = Abs(Value)
            Case ">"
                Value = Value > Token(Expression, iCounter)
                Value = Abs(Value)
            Case "="
                iTmp = Token(Expression, iCounter)
                Value = Value = iTmp
                Value = Abs(Value)
            Case "-"
                iTmp = Token(Expression, iCounter)
                Value = Value - iTmp
            Case "+": Value = Value + Token(Expression, iCounter)
            Case "*": Value = Value * Token(Expression, iCounter)
            Case "/": Value = Value / Token(Expression, iCounter)
            Case "\": Value = Value \ Token(Expression, iCounter)
            Case "&": Value = Value & Token(Expression, iCounter)
        End Select
    Loop
    
    StrCh = ""
    If IsNumeric(Value) Then
        Eval = Val(Value)
    Else
        Eval = CStr(Value)
    End If
    
End Function

Function Token(Expression, Position)
Dim s_FuncName As String, sTempScipt As String

    Dim StrCh As String
    Dim pl As Integer, es As Integer
    
    Do Until Position > Len(Expression)
        StrCh = Mid(Expression, Position, 1)
        
        Select Case StrCh
            Case "+", "-", "*", "^", "/", "\", "&", "<", "=", ">", "%"
                Exit Do
            Case "("
                Position = Position + 1
                pl = 1
                es = Position
                
                Do Until pl = 0 Or Position > Len(Expression)
                    StrCh = Mid(Expression, Position, 1)
                    If StrCh = "(" Then pl = pl + 1
                    If StrCh = ")" Then pl = pl - 1
                    Position = Position + 1
                Loop
                
                Value = Mid(Expression, es, Position - es - 1)
                
                s_FuncName = LCase(Trim(Token))
                
                On Error GoTo FlagErr:
                
                Select Case s_FuncName
                    Case "cstr": Token = CStr(Eval(Value))
                    Case "arrytostr": Token = ArrayToStr(CStr(Value))
                    Case "sendmessage": Token = SendMsg(CStr(Value))
                    Case "rgb": Token = dmRgb(Value)
                    ' files and folders
                    Case "fopen":  Token = fopen(Value)
                    Case "copyfile": Token = dmFunction1(Value, dmFileCopy, 1, s_FuncName)
                    Case "deletefile": Token = DelFile(CStr(Value))
                    ' String functions
                    Case "left": Token = dmFunction1(Value, dmLeft, 1, s_FuncName)
                    Case "right": Token = dmFunction1(Value, dmRight, 1, s_FuncName)
                    Case "asc": Token = Asc(Eval(Value))
                    Case "fillstr": Token = dmFunction1(Value, dmString, 1, s_FuncName)
                    Case "chr": Token = Chr(Eval((Value)))
                    Case "lcase": Token = LCase(Eval(Value))
                    Case "ucase": Token = UCase(Eval(Value))
                    Case "space": Token = Space(Eval(Value))
                    Case "ltrim": Token = LTrim(Eval(Value))
                    Case "rtrim": Token = RTrim(Eval(Value))
                    Case "trim": Token = Trim(Eval(Value))
                    Case "strlen": Token = Len(Eval(Value))
                    Case "strcpy": Token = dmFunction1(Value, dmStrCpy, 1, s_FuncName, True)
                    'arrays
                    Case "lbound": Token = ArrayBounds(Value, 1, "LBound")
                    Case "ubound": Token = ArrayBounds(Value, 2, "UBound")
                    Case "isarray": Token = ArrayIndex(CStr(Value)) <> -1
                    ' math functions
                    Case "abs": Token = Abs(Eval(Value))
                    Case "xor":  Token = dmFunction1(Value, dmXOR, 1, s_FuncName)
                    Case "sin": Token = Sin(Eval(Value))
                    Case "cos": Token = Cos(Eval(Value))
                    Case "rnd": Token = Rnd(Eval(Value))
                    Case "rndex": Token = Rand(Eval(Value))
                    Case "round": Token = nRound(Value)
                    Case "sqr": Token = Sqr(Eval(Value))
                    Case "log": Token = Log(Eval(Value))
                    Case "eval": Token = Eval(Value)
                    Case "fix": Token = Fix(Eval(Value))
                    Case "hex": Token = dmHex(Value)
                    Case "val": Token = Eval(Value)
                    Case "itoa": Token = Eval(Value)
                    Case "countif": Token = dmFunction1(Value, dmCountIf, 1, s_FuncName)
                    
                    ' System
                    Case "env": Token = Environ(Eval(Value))
                    ' COM Support
                    Case "mycomcall": Token = MyComCall(Value) ' Used for calling ActiveX COM DLLS
                    ' GUI
                    Case "echo": Token = Echo(Value)
                    Case "inputbox": Token = Eval(InputBoxA(Value))
                    ' check if it's a user dinfined function
                    Case Else
                        Token = "-1"
                        sTempScipt = DoCallProc(s_FuncName & "(" & Value & ")") ' we need to rebuid the string
                        Execute sTempScipt, s_FuncName
                        If sReturnFormFunc = vbNullChar Then Abort 3, "'" & s_FuncName & "' must return a value": Exit Function
                        If Not IsNumeric(sReturnFormFunc) Then sReturnFormFunc = CStr(sReturnFormFunc)
                        Token = sReturnFormFunc: sReturnFormFunc = vbNullChar
                End Select
                
                Case Chr(34)
                    pl = 1
                    Position = Position + 1
                    es = pos
                    
                    Do Until pl = 0 Or Position > Len(Expression)
                        StrCh = Mid(Expression, Position, 1)
                        Position = Position + 1
                        If StrCh = Chr(34) Then
                            If Mid(Expression, Position, 1) = Chr(34) Then
                                Value = Value & Chr(34)
                                Position = Position + 1
                            Else
                                Exit Do
                            End If
                        Else
                            Value = Value & StrCh
                        End If
                    Loop
                Token = Value
                
            Case Else
                Token = Token & StrCh
                Position = Position + 1
        End Select
   Loop
   
   Token = ReturnData(CStr(Token))
   
   If IsNumeric(Token) Then
        Token = Val(Token)
    Else
        Token = CStr(Token)
    End If
    
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Function



