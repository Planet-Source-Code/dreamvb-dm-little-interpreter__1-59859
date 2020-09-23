Attribute VB_Name = "modFunc"
' This mod has any inbuild functions we have
Enum FuncLst
    dmString = 0
    dmLeft
    dmRight
    dmCountIf
    dmStrCpy
    dmXOR
    dmFileCopy
End Enum

Function fopen(lParm As Variant)
' I not finsihed this function yet

Dim vInfo As Variant, StrA As String, StrB As Integer
On Error GoTo tErr:
    If CountIF(CStr(lParm), ",") < 1 Then Abort 12, "fopen": Exit Function
    vInfo = Split(lParm, ",")
    StrA = Eval(vInfo(0)) ' Get the filename
    StrB = Eval(vInfo(1)) ' get the filemode see file consts modpublic
    
    If m_FilePtr <> 0 Then Abort 3, "File already open": Exit Function
    
    m_FilePtr = FreeFile
    
    Select Case StrB
        Case 1: Open StrA For Binary Access Read As m_FilePtr 'read
        Case 2: Open StrA For Binary Access Write As m_FilePtr 'write
        Case 3: Open StrA For Append As m_FilePtr 'append
        Case 4: Open StrA For Output As m_FilePtr 'output
    End Select
    fopen = 1
    
    StrA = ""
    Erase vInfo
    
    Exit Function
tErr:
    If Err Then Abort 3, Err.Description
    
    
End Function

Function GetControlPropInfo(lzStr As String, Position As Integer) As String
Dim e_pos As Integer, n_pos As Integer

    If Position = 1 Then ' get left side
        e_pos = GetCharPos(lzStr, "=")
        n_pos = InStr(e_pos + 1, lzStr, ".", vbBinaryCompare)
        GetControlPropInfo = Trim(Mid(lzStr, e_pos + 1, n_pos - e_pos - 1))
    Else
        e_pos = GetCharPos(lzStr, ".")
        GetControlPropInfo = Trim(Mid(lzStr, e_pos + 1, Len(lzStr)))
    End If
    
End Function

Function ArrayBounds(lParm As Variant, op As Integer, StrOp As String)
Dim Idx As Integer
    Idx = -1
    
    If Len(Trim(lParm)) = 0 Then Abort 3, StrOp
    Idx = ArrayIndex(CStr(lParm))
    
    If op = 1 Then  'get lower bound
        If Idx = -1 Then ArrayBounds = -1: Exit Function
        ArrayBounds = 0
        Exit Function
    End If
    
    If op = 2 Then ' get upper bound
        If Idx = -1 Then ArrayBounds = -1: Exit Function
        ArrayBounds = GetArrayUBound(Idx)
    End If
    
End Function

Function nRound(lParm As Variant) As Variant
On Error GoTo tErr:
Dim k As Double

Dim vInfo As Variant
    If Len(Trim(lParm)) = 0 Then Abort 12, "Round": Exit Function
    vInfo = Split(lParm, ",")
    
    If UBound(vInfo) = 0 Then
        nRound = Round(Eval(vInfo(0)))
    Else
        nRound = Round(Eval(vInfo(0)), CLng(Eval(vInfo(1))))
    End If
    
    Erase vInfo

    Exit Function
tErr:
    If Err Then Abort 3, Err.Description & vbCrLf & "Round"
    
End Function

Public Function Rand(num) As Integer
    Randomize
    Rand = Int(num * Rnd)
End Function

Function DelFile(lParm As String) As Long
Dim StrA As String
    StrA = Eval(lParm)
    DelFile = DeleteFile(StrA)
    StrA = ""
End Function

Function ArrayToStr(lzStr As String)
Dim StrA As String, StrB As String, StrC As String, iCount As Integer
Dim sTmp As String
Dim e_pos As Integer, n_pos As Integer

    iCount = CountIF(lzStr, ",")
    If iCount < 1 Then Abort 12, "ArrayToStr": Exit Function
    
    If iCount = 1 Then
        e_pos = InStr(1, lzStr, ",", vbBinaryCompare)
        If e_pos <> 0 Then
            StrA = Trim(Mid(lzStr, 1, e_pos - 1))
            StrB = Trim(Mid(lzStr, e_pos + 1, Len(lzStr)))
        End If
    End If
    
    If iCount > 1 Then
        e_pos = InStr(1, lzStr, ",", vbBinaryCompare)
        n_pos = InStr(e_pos + 1, lzStr, ",", vbBinaryCompare)
        If (e_pos > 0 And n_pos > 0) Then
            StrA = Trim(Mid(lzStr, 1, e_pos - 1))
            StrB = Trim(Mid(lzStr, e_pos + 1, n_pos - e_pos - 1))
            StrC = Eval(Trim(Mid(lzStr, n_pos + 1, Len(lzStr))))
        End If
    End If
    
    If Not isArrayEx(StrA) Then Abort 1, "Array variable name": Exit Function
    If VariableIndex(StrB) = -1 Then Abort 7: Exit Function
    ' Loop thought all the array items and make a string from them
    
    e_pos = ArrayIndex(StrA)
    For I = 0 To GetArrayUBound(e_pos)
        sTmp = sTmp & StrC & Arrays(e_pos).ArrayData(I)
    Next
    
    If Left(sTmp, 1) = StrC Then sTmp = Right(sTmp, Len(sTmp) - 1)
    
    ArrayToStr = sTmp
    
    sTmp = ""
    StrA = ""
    StrB = ""
    StrC = ""
    iCount = 0
    e_pos = 0
    n_pos = 0
    
End Function

Function dmFunction1(lParm As Variant, FuncOp As FuncLst, FuncParmCnt As Integer, Optional StrFuncName As String, Optional isOptional As Boolean) As Variant
Dim vInfo As Variant, ParmCount As Integer, iTmp As Integer

On Error GoTo tErr:
    ParmCount = CountIF(CStr(lParm), ",")
    iTmp = 0
    
    If Not isOptional Then
        If ParmCount <> FuncParmCnt Then Abort 12, StrConv(StrFuncName, vbProperCase): Exit Function
    End If
    
    vInfo = Split(lParm, ",")
    
    Select Case FuncOp
        Case dmString
            dmFunction1 = String(Eval(vInfo(0)), Eval(vInfo(1)))
        Case dmLeft
            dmFunction1 = Left(Eval(vInfo(0)), Eval(vInfo(1)))
        Case dmRight
            dmFunction1 = Right(Eval(vInfo(0)), Eval(vInfo(1)))
        Case dmCountIf
            dmFunction1 = CountIF(CStr(Eval(vInfo(0))), CStr(Eval(vInfo(1))))
        Case dmStrCpy
            On Error Resume Next
            If UBound(vInfo) < 1 Then Abort 12, StrFuncName: Exit Function
            If ParmCount = FuncParmCnt Then dmFunction1 = Mid(Eval(vInfo(0)), Eval(vInfo(1))): Exit Function
            If ParmCount > FuncParmCnt Then dmFunction1 = Mid(Eval(vInfo(0)), Eval(vInfo(1)), Eval(vInfo(2)))
            If Err Then Abort 3, Err.Description: Exit Function
        Case dmXOR
            dmFunction1 = Eval(vInfo(0)) Xor Eval(vInfo(1))
        Case dmFileCopy
            dmFunction1 = CopyFile(CStr(Eval(vInfo(0))), CStr(Eval(vInfo(1))), 0)
    End Select
    
    Erase vInfo
    
tErr:
    Exit Function
    If Err Then Abort 3, Err.Description
    
End Function

Function SendMsg(lParm As String) As Long
On Error Resume Next
Dim nHwnd As Long
Dim isStr As Boolean
Dim sTmp1 As String

    Dim vLst As Variant
    
    If CountIF(lParm, ",") < 3 Then Abort 12, "SendMessage": Exit Function
    vLst = Split(lParm, ",")
    
    nHwnd = CLng(Eval(vLst(0)))
    sTmp1 = Eval(vLst(3))
    
    If IsNumeric(sTmp1) Then
        SendMsg = SendMessage(nHwnd, 384, 0, ByVal CLng(sTmp1))
    Else
        SendMsg = SendMessage(nHwnd, 384, 0, ByVal sTmp1)
    End If

    Erase vLst
    If Err Then SendMsg = 0
    
End Function

Function dmRgb(lParm As Variant) As Long
Dim vHold As Variant
On Error GoTo FlagErr:
    If CountIF(CStr(lParm), ",") <> 2 Then Abort 12, "RGB"
    vHold = Split(lParm, ",")
    dmRgb = RGB(vHold(0), vHold(1), vHold(2))
    Erase vHold
    
    Exit Function
FlagErr:
    Erase vHold
    If Err Then Abort 3, Err.Description & " RGB"
End Function

Function Echo(lpParms As Variant) As VbMsgBoxResult
Dim vParms As Variant, ParmCnt As Integer
Dim msg_txt As String, msg_buttons As Integer, msg_Caption As String
' Display a message box on the screen
On Error GoTo FlagErr:

    If Len(Trim(lpParms)) <= 0 Then Abort 12, "Echo": Exit Function
    
    vParms = Split(lpParms, ",")

    ParmCnt = UBound(vParms)
    
    If ParmCnt = 0 Then
        msg_txt = Eval(CStr(vParms(0)))
    ElseIf ParmCnt = 1 Then
        msg_txt = Eval(CStr(vParms(0)))
        msg_buttons = CInt(Eval(CStr(vParms(1))))
    Else
        msg_txt = Eval(CStr(vParms(0)))
        msg_buttons = CInt(Eval(CStr(vParms(1))))
        msg_Caption = Eval(CStr(vParms(2)))
    End If
   
    ' check if any values are found if not use deafult
    If msg_Caption = "" Then msg_Caption = "Little# Interpreter"
    If LinkError Then Exit Function
    Echo = MsgBox(msg_txt, msg_buttons, msg_Caption)
    
    
    'Clean up
    Erase vParms
    msg_txt = ""
    msg_Caption = ""
    msg_txt = ""
    
    Exit Function
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Function

Function InputBoxA(lpParms As Variant) As Variant
Dim vParms As Variant, ParmCnt As Integer
Dim Input_Disp As String, Input_Caption As String
Dim lp_Return As Variant
On Error GoTo FlagErr:
    ' Standred inputbox function
    If Len(Trim(lpParms)) <= 2 Then Abort 12, "InputBox": Exit Function
    vParms = Split(lpParms, ",")
    
    ParmCnt = UBound(vParms)
    
    If ParmCnt = 0 Then
        Input_Disp = Eval(CStr(vParms(0)))
        Input_Caption = "SharpScript"
    Else
        Input_Disp = Eval(CStr(vParms(0)))
        Input_Caption = Eval(CStr(vParms(1)))
    End If
    
    lp_Return = InputBox(Input_Disp, Input_Caption)
    
    If IsNumeric(lp_Return) Then
        InputBoxA = Val(lp_Return)
    Else
        InputBoxA = AddQuotes(CStr(lp_Return))
    End If
    
    'Clean up
    Erase vParms
    Input_Disp = ""
    Input_Disp = ""

    Exit Function
FlagErr:
    If Err Then Abort 3, Err.Description
End Function

Function CreateComObject(lpParm As String) As Integer
Dim StrObjServ As String, StrTemp As String, StrObjName As String
Dim TmpObj As Object
Dim AssignPos As Integer
    ' This function is used for the calling to ActiveX DLLs and createing the objects
    CreateComObject = 0
    AssignPos = GetCharPos(lpParm, "=")
    If AssignPos = 0 Then Abort 1, "=": Exit Function ' Check for assign sign =
    
    StrObjName = Trim(Mid(lpParm, 1, AssignPos - 1))
    If StrObjName = "" Then Abort 7: Exit Function ' check for the object name
    
    StrTemp = Trim(Mid(lpParm, AssignPos + 1, Len(lpParm) - AssignPos))
    
    If (GetProcInfo(StrTemp, 1) = 0 Or GetProcInfo(StrTemp, 2) = 0) Then
        ' No brackets found so it must be an assign to nothing or anougther variable
        ' I sort this out in the next versions
        Exit Function
    End If
    
    StrObjServ = Eval(Mid(StrTemp, GetProcInfo(StrTemp, 1) + 1, _
    GetProcInfo(StrTemp, 2) - GetProcInfo(StrTemp, 1) - 1))

    If UCase(GetProcInfo(StrTemp, 5, 0, GetProcInfo(StrTemp, 1))) = "CREATEMYCOM" Then
        On Error GoTo ObjCreateErr:
        Set TmpObj = CreateObject(StrObjServ) ' Create the com object
        
        If GetObjIndex(StrObjName) <> -1 Then
            DestroyObject GetObjIndex(StrObjName)
            SetObject GetObjIndex(StrObjName), TmpObj
            CreateComObject = 1
            Set TmpObj = Nothing
            StrObjServ = ""
            StrTemp = ""
            AssignPos = 0
            Exit Function
        End If
        
        AddObject StrObjName, TmpObj
        
        CreateComObject = 1
        'Clean up time
        Set TmpObj = Nothing
        StrObjServ = ""
        StrTemp = ""
        AssignPos = 0
        
        Exit Function
    End If
    
ObjCreateErr:
    If Err Then Abort 1, Err.Description
    CreateComObject = 0
    
End Function

Function MyComCall(lpParm As Variant) As Variant
Dim vParm As Variant
Dim lp_Return As Variant
Dim ObjName As String, ProcName As String, ProcCallType As Integer
On Error GoTo FlagErr:
Dim ObjIndex As Long
Dim TmpObj As Object
Dim UserParms() As Variant
Dim I As Integer, j As Integer


    Erase UserParms
    j = -1
    
    vParm = Split(lpParm, ",")
    If UBound(vParm) < 2 Then Abort 12, "MyComCall": Exit Function
    
    For I = 3 To UBound(vParm)
        ' Add the parms to the UserParms list
        j = j + 1
        ReDim Preserve UserParms(j)
        UserParms(j) = Eval(vParm(I))
    Next
    
    ObjName = vParm(0) ' Get the object name
    ObjIndex = GetObjIndex(ObjName) ' Get the objects index
    If ObjIndex = -1 Then Abort 1, "Invalid procedure call or argument": Exit Function
    ' Above checks for a vaild index of the object if not found we return an error
    ProcName = Trim(Eval(vParm(1)))  'next we get the procedure name of the object to call
    If Len(ProcName) = 0 Then Abort 1, "Procedure name ": Exit Function
    ProcCallType = CInt(Eval(vParm(2)))
    'Ok we can now make a call to the object
    
    Set TmpObj = GetObjectX(ObjIndex) 'Create the temp object
    ' Now we can use the vbCallByName function to make a call to the object

    If (j > -1) Then
        On Error Resume Next
        lp_Return = CallByName(TmpObj, ProcName, ProcCallType, UserParms)
        If Err Then Abort 3, Err.Description
    Else
        On Error Resume Next
        lp_Return = CallByName(TmpObj, ProcName, ProcCallType)
        If Err Then Abort 3, Err.Description
    End If
    
    
    If IsNumeric(lp_Return) Then
        MyComCall = Val(lp_Return)
    Else
        MyComCall = AddQuotes(CStr(lp_Return))
    End If
    
    'Clean up time
    If Not IsEmpty(UserParms) Then Erase UserParms
    Erase vParm
    lp_Return = ""
    ObjName = ""
    ProcName = ""
    ProcCallType = 0
    ObjIndex = 0
    Set TmpObj = Nothing
    j = 0: I = 0
    
    Exit Function
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Function

Function dmHex(lParm As Variant) As Variant
Dim k As Variant

    Dim StrA As String
    StrA = lParm
    
    k = Hex(Eval(StrA))
    
    If IsNumeric(k) Then
        dmHex = k
    Else
        dmHex = AddQuotes(CStr(k))
    End If
    
    StrA = ""
    k = 0
    
End Function
