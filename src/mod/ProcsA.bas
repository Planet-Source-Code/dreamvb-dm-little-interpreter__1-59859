Attribute VB_Name = "ProcsA"
Private Type Proc
    ProcName As String ' it's name
    ProcParmCount As Integer ' number of parmators in the proc
    ProcCanReturn As Boolean ' true only for functions
    ProcCodeBlock As String ' Code for the proc
    ProcVarList As String
End Type

Public ProgStack() As Proc, ProcCounter As Long

Public Sub ResetProcStack()
    ProcCounter = -1
    Erase ProgStack
End Sub

Public Function ProcIndex(lpProcName As String) As Integer
Dim nIdx As Integer, I As Integer
    
    If ProcCounter = -1 Then ProcIndex = -1: Exit Function
    nIdx = -1
    
    For I = 0 To UBound(ProgStack)
        If LCase(lpProcName) = LCase(ProgStack(I).ProcName) Then
            nIdx = I
            Exit For
        End If
    Next
    I = 0
    ProcIndex = nIdx
    
End Function

Public Function AddProc(sProcName As String, sParmCount As Integer, lpCanReturn As Boolean, lpVarPrt As String)
    ProcCounter = ProcCounter + 1
    ReDim Preserve ProgStack(ProcCounter)
    ProgStack(ProcCounter).ProcName = LCase(sProcName)
    ProgStack(ProcCounter).ProcParmCount = sParmCount
    ProgStack(ProcCounter).ProcCanReturn = lpCanReturn
    ProgStack(ProcCounter).ProcVarList = lpVarPrt ' new
End Function
 
Public Sub AddProcCode(ProcIndex As Integer, ProcData As String)
    ProgStack(ProcIndex).ProcCodeBlock = ProcData
End Sub

Public Function ProcGetCodeBlock(ProcName As String) As String
    If ProcIndex(ProcName) = -1 Then ProcGetCodeBlock = vbNullChar: Exit Function
    ProcGetCodeBlock = ProgStack(ProcIndex(ProcName)).ProcCodeBlock
End Function

Public Function GetProcParmCount(ProcName As String) As Integer
    If ProcIndex(ProcName) = -1 Then GetProcParmCount = -1: Exit Function
    GetProcParmCount = ProgStack(ProcIndex(ProcName)).ProcParmCount
End Function

Public Function GetProcParmList(ProcName As String) As String
    If ProcIndex(ProcName) = -1 Then GetProcParmList = "": Exit Function
    GetProcParmList = ProgStack(ProcIndex(ProcName)).ProcVarList
End Function

