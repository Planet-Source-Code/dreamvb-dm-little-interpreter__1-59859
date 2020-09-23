VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Our FileIO system not a lot yet like
' OK one thing to remmber with my COM support is that
' add parms come in as an array or Variant
' see the examples below to give you an idea

Function CreateTextFile(lParm As Variant) As Integer
    If UBound(lParm) < 1 Then CreateTextFile = 0: Exit Function

    Open lParm(0) For Binary As #1
        Put #1, , CStr(lParm(1))
    Close #1
    
    CreateTextFile = 1
    
End Function

Function FixPath(lParm As Variant) As String
Dim s As String
    s = lParm(0)
    If Right(s, 1) = "\" Then FixPath = s Else FixPath = s & "\"
    s = ""
    Erase lParm
End Function

Function OpenTextFile(lParm As Variant) As String
Dim StrB As String
    
    Open lParm(0) For Binary As #1
        StrB = Space(LOF(1))
        Get #1, , StrB
    Close #1
    
    OpenTextFile = StrB
    StrB = ""
    Erase lParm
    
End Function
