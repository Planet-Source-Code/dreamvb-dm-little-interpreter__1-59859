Attribute VB_Name = "dmcomp"
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
                
                If Not IsFileHere(AbsFile) Then MsgBox "Include file not found:" & vbCrLf & AbsFile: End
                sBuff = OpenFile(AbsFile)
                StrA = StrA & sBuff & vbCrLf
                AbsFile = "": e_pos = 0: n_pos = 0: s_pos = 0: incFile = ""
                lzCode = Replace(lzCode, sline, "")
            End If
        End If
    Next
    GetIncludeFiles = StrA
End Function

Function GetPath(lzFile As Variant) As String
Dim e_pos As Integer
    e_pos = InStrRev(lzFile, "\", Len(lzFile), vbBinaryCompare)
    If e_pos <> 0 Then
        GetPath = Mid(lzFile, 1, e_pos)
        e_pos = 0
    End If
End Function

Function GetFileName(lzFile As Variant) As String
Dim e_pos As Integer
    e_pos = InStrRev(lzFile, "\", Len(lzFile), vbBinaryCompare)
    If e_pos <> 0 Then
        GetFileName = Mid(lzFile, e_pos + 1, Len(lzFile))
        e_pos = 0
    End If
End Function

Function FixPath(lzPath As String) As String
    If Right(lzPath, 1) = "\" Then FixPath = lzPath Else FixPath = lzPath & "\"
End Function

Public Function IsFileHere(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then IsFileHere = False: Exit Function Else IsFileHere = True
End Function

Function Encrypt(lzStr As String) As String
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

Private Function OpenFile(lzFile As String) As String
Dim ByteBuff() As Byte
    nFile = FreeFile
    Open lzFile For Binary As #nFile
        ReDim ByteBuff(1 To LOF(nFile))
        Get #nFile, , ByteBuff()
    Close #nFile
    
    OpenFile = StrConv(ByteBuff(), vbUnicode)
    Erase ByteBuff
    
End Function

Sub Main()
Dim lzCommand As String, vParmLst As Variant, iParmCnt As Integer
Dim nEncode As Boolean, sTemp1 As String, lzFileName As String, StrHead As String
Dim sTemp1uff As String, sTemp2 As String, sTemp3 As String, StrBuffer As String
Dim iFile As Long
    
    ' This is our little compiler for linking the scripts to an exe
    ' This file must remain in the same folder as Little# Interpreter
    Const ErrMsg As String = "Incorrect command line parameters."
    
    StrHead = Chr(5) & Chr(255) & "DM#"
    nEncode = False
    
    lzCommand = Command
    If Len(lzCommand) = 0 Then
        MsgBox ErrMsg
        End
    End If
    
    vParmLst = Split(lzCommand, "/", , vbBinaryCompare)
    iParmCnt = UBound(vParmLst)
    
    If iParmCnt = 0 Then
        MsgBox ErrMsg
        End
    End If

    If Not LCase(Trim(vParmLst(1))) = "make" Then
        MsgBox ErrMsg
        End
    Else
        lzFileName = vParmLst(0) 'Open the script file
        sTemp1 = OpenFile(lzFileName) ' Open the source file
        ' now we need to also add any include files to the script
        sTemp3 = Encrypt(GetIncludeFiles(sTemp1) & sTemp1)
    
        StrBuffer = StrHead & sTemp3
        sTemp1 = "": StrHead = "": sTemp3 = ""
        lzFileName = FixPath(App.Path) & "bscript.exe"
        If Not IsFileHere(lzFileName) Then
            MsgBox "File not found " & vbCrLf & vbCrLf & lzFileName
            Erase vParmLst: lzFileName = ""
            End
        Else
            sTemp1 = GetPath(vParmLst(0))
            sTemp2 = GetFileName(vParmLst(0))
            sTemp2 = Left(sTemp2, Len(sTemp2) - 4) & "exe"
            sTemp1 = sTemp1 & sTemp2
            FileCopy lzFileName, sTemp1
            
            'write the encrypted data to the end of the exe
            iFile = FreeFile
            Open sTemp1 For Binary Access Write As #iFile
                Put #iFile, LOF(iFile) + 1, StrBuffer
            Close #iFile
            
            StrBuffer = ""
            sTemp1 = ""
            sTemp2 = ""
            End
        End If
    End If
    
End Sub
