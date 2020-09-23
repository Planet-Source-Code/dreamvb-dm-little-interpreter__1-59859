Attribute VB_Name = "modVars"
' This module as you may have guessed is for the variables.
Public Enum VarTypes
    dmNotKnown = -1
    dmInteger = 1
    dmBool = 2
    dmChar = 3
    dmFloat = 4
    dmlong = 5
End Enum

Private Type Variable ' A house :) to keep our our variables
    VariableName As String
    VariableType As VarTypes
    VariableData As Variant
    isReadOnly As Boolean
End Type

Private Type TArrayItems
    mArrayBoundData() As Variant ' used to hold the data of an array
End Type

Private Type dmArray ' our array type
    ArrayName As String
    ArrayType As VarTypes
    ArrayData() As Variant
End Type

Public Vars() As Variable
Public Arrays() As dmArray

Public VarCount As Long ' Keep a count of all the variables
Public ArrayCount As Long ' Keep a count of the arrays we have

Public Function VarNotkeyword(lpVarName As String) As Boolean
Dim vKeywords As Variant

    VarNotkeyword = False
    vKeywords = Split(KeyWords, ",")
    
    For I = 0 To UBound(vKeywords)
        If LCase(lpVarName) = vKeywords(I) Then
            VarNotkeyword = True
            Exit For
        End If
    Next
    
    I = 0
    Erase vKeywords
    
End Function

Public Sub SetSystemVariable(lpVarName As String, lpVarData As Variant)
On Error GoTo SetError:
    Select Case LCase(lpVarName)
        Case "time": Time = lpVarData
        Case "date": Date = lpVarData
    End Select
    
    Exit Sub
SetError:
    If Err Then Abort 9, Err.Description
End Sub

Public Function GetSystemVar(lpVarName As String) As Variant
    Select Case LCase(lpVarName)
        Case "time": GetSystemVar = Time
        Case "date": GetSystemVar = Date
        Case "timer": GetSystemVar = Timer
    End Select
End Function

Public Function isSystemVar(lpVarName As String) As Boolean
    Select Case LCase(lpVarName)
        Case "time": isSystemVar = True
        Case "date": isSystemVar = True
        Case "timer": isSystemVar = True
    End Select
End Function

Public Sub ResetVarStack()
    ArrayCount = -1
    VarCount = -1 ' reset the variable counter
    Erase Vars ' erase any data that Vars may contain
    Erase Arrays
End Sub

Public Function VariableIndex(sVarName As String) As Integer
Dim X As Integer, Idx As Integer

    ' Looks for a variable's name then returns it's index
    If VarCount = -1 Then VariableIndex = -1: Exit Function
    Idx = -1
    For X = 0 To UBound(Vars)
        If LCase(Trim(sVarName)) = LCase(Vars(X).VariableName) Then
            Idx = X
            Exit For
        End If
    Next
    VariableIndex = Idx
End Function

Public Sub AddVariable(sVarName As String, sVarType As String, VarReadOnly As Boolean, Optional VarData As Variant)
    VarCount = VarCount + 1 ' add one to our variable counter
    ReDim Preserve Vars(VarCount) ' resize the varstack
    Vars(VarCount).VariableName = sVarName ' add the variable name
    Vars(VarCount).VariableType = GetVariableTypeFromStr(sVarType) ' add it's datatype see Enum VarTypes
    Vars(VarCount).VariableData = VarData ' add the variables data
    Vars(VarCount).isReadOnly = VarReadOnly ' is the variable readonly eg user and built in consts are set to true
End Sub

Public Sub AddArray(sArrayName As String, sArrayType As String, Optional ArraySize As Long)
    ArrayCount = ArrayCount + 1
    ReDim Preserve Arrays(ArrayCount)
    Arrays(ArrayCount).ArrayName = sArrayName
    Arrays(ArrayCount).ArrayType = GetVariableTypeFromStr(sArrayType)
    If ArraySize <> -1 Then ReDim Preserve Arrays(ArrayCount).ArrayData(ArraySize)
End Sub

Function isArrayEx(StrArrName As String) As Boolean
    isArrayEx = ArrayIndex(StrArrName) <> -1
End Function

Function ArrayIndex(sArrayName As String) As Integer
Dim Idx As Integer, X As Integer
    ' Locate an arrays index
    Idx = -1
    ArrayIndex = -1
    If ArrayCount = -1 Then Exit Function
    
    For X = 0 To UBound(Arrays)
        If LCase(Arrays(X).ArrayName) = LCase(sArrayName) Then
            Idx = X
            Exit For
        End If
    Next
    
    ArrayIndex = Idx
    X = 0
    
End Function

Public Function GetVariableTypeFromStr(lpVarType As String) As VarTypes
    Select Case lpVarType
        Case "int"
            GetVariableTypeFromStr = dmInteger
        Case "bool"
            GetVariableTypeFromStr = dmBool
        Case "char"
            GetVariableTypeFromStr = dmChar
        Case "float"
            GetVariableTypeFromStr = dmFloat
        Case "long"
            GetVariableTypeFromStr = dmlong
        Case Else
            GetVariableTypeFromStr = dmNotKnown
    End Select
End Function

Public Function DefaultVarData(lVarType As Integer) As Variant
    ' here we use this to set a variables default data based on it's datatype
    Select Case lVarType
        Case dmInteger, dmlong: DefaultVarData = 0
        Case dmBool: DefaultVarData = False
        Case dmChar: DefaultVarData = ""
        Case dmFloat: DefaultVarData = 0
    End Select
End Function

Public Function SetVarDataX(lVarType As Integer, mVarData As Variant) As Variant
On Error Resume Next
    If Len(mVarData) = 0 Then mVarData = 0
    
    Select Case lVarType
        Case dmInteger: SetVarDataX = CInt(mVarData)
        Case dmlong: SetVarDataX = CLng(mVarData)
        Case dmBool: SetVarDataX = Abs(CBool(mVarData))
        Case dmChar: SetVarDataX = CStr(mVarData)
        Case dmFloat: SetVarDataX = CSng(mVarData)
    End Select
    
    If Err Then Abort 3, Err.Description
    
End Function

Function GetVariableData(sVarName As String) As Variant
    ' return the data from a variable
    If VariableIndex(sVarName) = -1 Then Exit Function ' No index found so we do nothing
    If isSystemVar(Trim(sVarName)) Then GetVariableData = GetSystemVar(Trim(sVarName)): Exit Function
    GetVariableData = Vars(VariableIndex(sVarName)).VariableData ' get the variables data
End Function

Public Sub SetVariable(sVarName As String, sVarData As Variant)
    If VariableIndex(sVarName) = -1 Then Abort 0, sVarName: Exit Sub ' No index found abort
    ' check if the variable is a const
    If Vars(VariableIndex(sVarName)).isReadOnly Then Abort 10: Exit Sub
    If isSystemVar(sVarName) Then SetSystemVariable sVarName, sVarData: Exit Sub
    Vars(VariableIndex(sVarName)).VariableData = sVarData ' set the variables data
End Sub

Public Sub SetArrayData(ArrayIdx As Integer, aArrySlotIdx As Integer, nArrayData As Variant)
On Error Resume Next
    ' Set the arrays data
    Arrays(ArrayIdx).ArrayData(aArrySlotIdx) = SetVarDataX(GetArrayDataType(ArrayIdx), nArrayData)
    If Err Then Abort 3, Err.Description
End Sub

Public Function GetVariableType(sVarName As String) As VarTypes
    If VariableIndex(sVarName) = -1 Then GetVariableType = -1: Exit Function ' abort
    GetVariableType = Vars(VariableIndex(sVarName)).VariableType ' get the variables datatype
End Function

Function GetArrayUBound(ArrayIdx As Integer) As Integer
On Error Resume Next
    ' Return the upper bound of an array
    If ArrayIdx = -1 Then GetArrayUBound = -1
    GetArrayUBound = UBound(Arrays(ArrayIdx).ArrayData())
    If Err Then GetArrayUBound = -1: Err.Clear
End Function

Sub DestroyArray(ArrayIdx As Integer)
    On Error Resume Next
    If ArrayIdx = -1 Then Abort 1, "array variable": Exit Sub
    Erase Arrays(ArrayIdx).ArrayData
End Sub

Function GetArrayDataType(ArrayIdx) As VarTypes
    If ArrayIdx = -1 Then GetArrayDataType = dmNotKnown: Exit Function
    GetArrayDataType = Arrays(ArrayIdx).ArrayType
End Function

Function GetArrayData(ArrayIdx As Integer, aArrySlotIdx As Integer) As Variant
    On Error Resume Next
    If ArrayIdx = -1 Then Exit Function
    GetArrayData = Arrays(ArrayIdx).ArrayData(aArrySlotIdx)
    If Err Then Abort 3, Err.Description
End Function

Function ResizeArray(ArrayIdx As Integer, nSize As Integer)
    ReDim Preserve Arrays(ArrayIdx).ArrayData(nSize)
End Function
