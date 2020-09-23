VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' a Test COM
' OK one thing to remmber with my COM support is that
' add parms come in as an array or Variant
' see the examples below to give you an idea

Function StrFuncReturn() As String
    StrFuncReturn = "Hello from myDll " _
    & vbCrLf & "I was called from function StrFuncReturn."
End Function

Function AddTwoNums(lParm As Variant) As Variant
    AddTwoNums = Int(lParm(0)) + Int(lParm(1))
End Function

Function MessageBox(lParm As Variant) As VbMsgBoxResult
   MessageBox = MsgBox(lParm(0), Int(lParm(1)), lParm(2))
End Function


