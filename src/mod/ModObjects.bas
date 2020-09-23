Attribute VB_Name = "ModObjects"
Private Type ObjType
    ObjName As String
    TheObj As Object
End Type

Public ObjCount As Long
Public ObjCollection() As ObjType

Public Sub ResetObjStack()
    ObjCount = -1
    Erase ObjCollection
End Sub

Function AddObject(lpObjName As String, lpObjRef As Object) As Boolean
On Error Resume Next
    ' This function is used to create an add the new object
    ObjCount = ObjCount + 1
    ReDim Preserve ObjCollection(ObjCount)
    ObjCollection(ObjCount).ObjName = lpObjName
    Set ObjCollection(ObjCount).TheObj = lpObjRef
End Function

Function GetObjIndex(lpObjName As String) As Integer
Dim Idx As Integer

    Idx = -1
    
    For I = 0 To ObjCount
        If UCase(lpObjName) = UCase(ObjCollection(I).ObjName) Then
            Idx = I
            Exit For
        End If
    Next
    
    GetObjIndex = Idx
    
End Function

Function GetObjectX(mObjIdx As Long) As Object
On Error Resume Next
    Set GetObjectX = ObjCollection(mObjIdx).TheObj
    If Err Then Abort 3, "Invalid procedure call or argument"
End Function

Function SetObject(mObjIdx As Integer, lpObjRef As Object)
On Error Resume Next
    If mObjIdx = -1 Then Exit Function
    Set ObjCollection(mObjIdx).TheObj = lpObjRef
    If Err Then Abort 3, Err.Description
End Function

Function DestroyObject(mObjIdx As Integer)
    On Error Resume Next
    Set ObjCollection(mObjIdx).TheObj = Nothing
    If Err Then Abort 3, Err.Description
End Function
