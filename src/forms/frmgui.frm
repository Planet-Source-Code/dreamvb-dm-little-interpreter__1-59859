VERSION 5.00
Begin VB.Form frmgui 
   BackColor       =   &H80000009&
   ClientHeight    =   2370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2010
   Icon            =   "frmgui.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2370
   ScaleWidth      =   2010
   StartUpPosition =   3  'Windows Default
   Tag             =   "FORM"
   Begin VB.TextBox Edit 
      Height          =   300
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Tag             =   "1"
      Text            =   "Edit"
      Top             =   630
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton button 
      Caption         =   "Button"
      Height          =   350
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LB 
      Height          =   450
      Index           =   0
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Tag             =   "1"
      Top             =   975
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox CheckBox 
      Caption         =   "CheckBox"
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "1"
      Top             =   1545
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Tmr 
      Enabled         =   0   'False
      Index           =   0
      Left            =   60
      Top             =   1935
   End
   Begin VB.Label lbStatic 
      BackStyle       =   0  'Transparent
      Caption         =   "Static"
      Height          =   240
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Tag             =   "1"
      Top             =   405
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmgui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function ControlPropRead(lzControl As String, lzControlProp As String) As Variant
Dim StrC As Variant, StrA As String, bObj As Object
On Error GoTo FlagErr:
    StrA = "": StrC = 0
    StrC = GetIndexFromStr(lzControl)
    If StrC = 0 Then Abort 3, "Control not found " & lzControl: Exit Function
    StrA = LCase(GetCtrlName(lzControl))

    Select Case StrA
        Case "button": Set cObj = button(Int(StrC))
        Case "dialog": Set cObj = frmgui
        Case "static": Set cObj = lbStatic(Int(StrC))
        Case "edit": Set cObj = Edit(Int(StrC))
        Case "tmr": Set cObj = Tmr(Int(StrC))
        Case "lb": Set cObj = LB(Int(StrC))
        Case "checkbox": Set cObj = CheckBox(Int(StrC))
        Case Else: ControlPropRead = lzControl & "." & lzControlProp: Exit Function
    End Select

    ControlPropRead = CallByName(cObj, LCase(lzControlProp), VbGet)
    StrA = "": StrC = 0: Set cObj = Nothing
    
    Exit Function
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Function

Public Sub DrawLine(lParm As String)
Dim StrTmpA As String, StrTmpB As String, vInfo As Variant
Dim DrawS As String, e_pos As Integer, n_pos As Integer

    On Error GoTo FlagErr:
    
    DrawS = ""

    StrTmpA = lParm
    
    If CountIF(StrTmpA, ",") > 2 Then
        n_pos = InStr(1, StrTmpA, "-", vbBinaryCompare)
        e_pos = InStr(n_pos + 1, StrTmpA, ")", vbBinaryCompare)
        If e_pos <> 0 Then
            StrTmpB = Mid(StrTmpA, 1, e_pos)
            StrTmpA = Trim(Mid(StrTmpA, e_pos + 2, Len(StrTmpA) - e_pos - 1))
            DrawS = UCase(Trim(Right(StrTmpA, 2)))
            If UCase(Right(StrTmpA, 2)) = ",B" Then StrTmpA = StrRemoveRight(StrTmpA, 2)
            If UCase(Right(StrTmpA, 3)) = ",BF" Then StrTmpA = StrRemoveRight(StrTmpA, 3)
            
            If Left(DrawS, 1) = "," Then DrawS = StrRemoveLeft(DrawS, 1)
            ' Clean up first part of the line function removeing any brackets
            StrTmpB = Replace(StrTmpB, ")", "")
            StrTmpB = Replace(StrTmpB, "(", "")
            StrTmpB = Replace(StrTmpB, "-", ",")
            n_pos = 0: e_pos = 0
        Else
            Abort 1, "-"
            Exit Sub
        End If
    Else
        Abort 12, "LINE"
        Exit Sub
    End If
    
    vInfo = Split(StrTmpB, ",")
    
   ' vInfo=split(StrTmpB
    ' Clean up the parts of the strings we do not need
    
    
    If DrawS = "B" Then
         frmgui.Line (Eval(vInfo(0)), _
         Eval(vInfo(1)))-(Eval(vInfo(2)), _
         Eval(vInfo(3))), Eval(StrTmpA), B
    ElseIf DrawS = "BF" Then
         frmgui.Line (Eval(vInfo(0)), _
         Eval(vInfo(1)))-(Eval(vInfo(2)), _
         Eval(vInfo(3))), Eval(StrTmpA), BF
    Else
        frmgui.Line (Eval(vInfo(0)), _
        Eval(vInfo(1)))-(Eval(vInfo(2)), _
        Eval(vInfo(3))), Eval(StrTmpA)
    End If
    
    
    Exit Sub
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Sub
Public Sub SetPixel(lParm As String)
Dim vInfo As Variant
On Error GoTo FlagErr:
    ' This sub is used to draw a pixel on to the screen
    If CountIF(lParm, ",") <> 2 Then Abort 12, "Line": Exit Sub
    
    vInfo = Split(lParm, ",")
    frmgui.PSet (Eval(vInfo(0)), Eval(vInfo(1))), Eval(vInfo(2))
    
    Erase vInfo
    Exit Sub
    
FlagErr:
    If Err Then Abort 3, Err.Description
    
End Sub

Function DoGuiError(lzError As String)
    GuiError = True
    Link_Error_Str = lzError: LinkError = True
End Function

Public Sub DoControlProp(lpObjName As String, lpProp As String)
Dim StrA As String, StrB As String, StrC As Variant, cObj As Object
Dim isDialog As Boolean, HasAssign As Boolean, HasParm As Boolean
Dim i_pos As Integer

On Error Resume Next

    HasParm = False
    GuiError = False
    isDialog = LCase(lpObjName) = "dialog"
    StrA = lpObjName
    
    StrB = TidyLine(lpProp)
    
    StrC = GetIndexFromStr(lpObjName)
    
    If Not IsNumeric(StrC) And Not isDialog Then
        DoGuiError "Type msistake " & lpObjName
        Exit Sub
    ElseIf Val(StrC) = 0 And Not isDialog Then
        DoGuiError "inavild control index " & lpObjName
        Exit Sub
    Else
        Select Case LCase(GetCtrlName(StrA))
            Case "button": Set cObj = button(Int(StrC))
            Case "dialog": Set cObj = frmgui
            Case "static": Set cObj = lbStatic(Int(StrC))
            Case "edit": Set cObj = Edit(Int(StrC))
            Case "tmr": Set cObj = Tmr(Int(StrC))
            Case "lb": Set cObj = LB(Int(StrC))
            Case "checkbox": Set cObj = CheckBox(Int(StrC))
            Case Else: Exit Sub 'DoGuiError "Control not found " & lpObjName
        End Select
        
        HasAssign = GetCharPos(StrB, "=")
        
        If HasAssign Then
            StrA = GetFromAssign(StrB, GetCharPos(StrB, "="), 0) ' Get Procname
            StrB = Eval(GetFromAssign(StrB, GetCharPos(StrB, "="), 1)) ' Get proc Value
            If LinkError Then DoGuiError Link_Error_Str & " " & lpProp: Exit Sub
            
            CallByName cObj, StrA, VbLet, StrB
        Else
            i_pos = InStr(1, StrB, " ", vbBinaryCompare)
            If i_pos <> 0 Then
                StrA = Trim(Mid(StrB, 1, i_pos))
                StrB = Eval(Trim(Mid(StrB, i_pos, Len(StrB))))
                HasParm = True
            Else
                StrA = Trim(StrB)
                StrB = ""
                HasParm = False
            End If
            
            If HasParm Then
                CallByName cObj, StrA, VbMethod, StrB
                If Err Then Abort 3, Err.Description
            Else
                CallByName cObj, StrA, VbMethod
                If Err Then Abort 3, Err.Description
            End If
            
        End If
            If Err.Number = 340 Then Err.Clear
            
        Set cObj = Nothing
        StrA = "": StrB = "": StrC = ""
        
    End If
    
    If Err Then DoGuiError Err.Description & vbCrLf & lpObjName & "." & lpProp

End Sub

Public Sub UnloadControls()
Dim X As Integer

    On Error Resume Next
    
        For X = 0 To button.Count - 1
            If X > 0 Then Unload button(X)
        Next
        For X = 0 To lbStatic.Count - 1
            If X > 0 Then Unload lbStatic(X)
        Next
        For X = 0 To Edit.Count - 1
            If X > 0 Then Unload Edit(X)
        Next
        
        For X = 0 To Tmr.Count - 1
            If X > 0 Then Unload Tmr(X)
        Next
        
        For X = 0 To LB.Count - 1
            If X > 0 Then Unload LB(X)
        Next
        
        For X = 0 To CheckBox.Count - 1
            If X > 0 Then Unload CheckBox(X)
        Next
        
        X = 0
        
End Sub

Public Sub UnloadGUI()
    Unload frmgui
End Sub

Function GUIAddControl(lParm As String) As Boolean
Dim CtrlProp As Variant
Dim iCount As Integer
Dim TheObj As Object
Dim sTmp As String
Dim ParmCount As Integer
On Error GoTo GuiError:

    ParmCount = CountIF(lParm, ",")
    If ParmCount = 0 Then Abort 12, "AddControl": Exit Function
    
    sTmp = Trim(UCase(Mid(lParm, 1, GetCharPos(lParm, ",") - 1)))
    
    If (ParmCount < 6) And sTmp <> "TMR" Then Abort 12, "AddControl": Exit Function
    
    CtrlProp = Split(lParm, ",")

    iCount = 0
 
    Select Case sTmp
        Case "BUTTON"
            iCount = button.Count
            Load button(iCount)
            Set TheObj = button(iCount)
        Case "STATIC"
            iCount = lbStatic.Count
            Load lbStatic(iCount)
            Set TheObj = lbStatic(iCount)
        Case "EDIT"
            iCount = Edit.Count
            Load Edit(iCount)
            Set TheObj = Edit(iCount)
        Case "CHECKBOX"
            iCount = CheckBox.Count
            Load CheckBox(iCount)
            Set TheObj = CheckBox(iCount)
        Case "LB"
            iCount = LB.Count
            Load LB(iCount)
            Set TheObj = LB(iCount)
        Case "TMR"
            iCount = Tmr.Count
            Load Tmr(iCount)
            Set TheObj = Tmr(iCount)
        Case Else
            Abort 3, "Inviald control name." & sTmp
      End Select
      
    If sTmp <> "TMR" Then
        TheObj.Top = Eval(CtrlProp(1))
        TheObj.Left = Eval(CtrlProp(2))
        TheObj.Height = Eval(CtrlProp(3))
        TheObj.Width = Eval(CtrlProp(4))
        TheObj.Enabled = Eval(CtrlProp(5))
        
        If sTmp = "EDIT" Or sTmp = "LB" Then
            TheObj.Text = Eval(CtrlProp(6))
        Else
            TheObj.Caption = Eval(CtrlProp(6))
        End If
        
        TheObj.Visible = True
        
    Else
        TheObj.Interval = Eval(CtrlProp(1))
        TheObj.Enabled = Eval(CtrlProp(2))
    End If
    
    Set TheObj = Nothing
    sTmp = ""
    
    Exit Function
GuiError:
      If Err Then Abort 3, Err.Description & vbCrLf & "ref:GUIAddControl"
      
      
End Function

Private Sub Button_Click(Index As Integer)
Dim StrProcPtr As String

    StrProcPtr = "button" & Index & "_click"
    If ProcIndex(StrProcPtr) = -1 Then
        StrProcPtr = ""
        Exit Sub
    Else
        Execute ProcGetCodeBlock(StrProcPtr)
    End If
    
End Sub

Private Sub CheckBox_Click(Index As Integer)
Dim StrProcPtr As String

    StrProcPtr = "checkbox" & Index & "_click"
    If ProcIndex(StrProcPtr) = -1 Then
        StrProcPtr = ""
        Exit Sub
    Else
        Execute ProcGetCodeBlock(StrProcPtr)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadControls
End Sub

Private Sub LB_Click(Index As Integer)
Dim StrProcPtr As String

    StrProcPtr = "lb" & Index & "_click"
    If ProcIndex(StrProcPtr) = -1 Then
        StrProcPtr = ""
        Exit Sub
    Else
        Execute ProcGetCodeBlock(StrProcPtr)
    End If
End Sub

Private Sub Tmr_Timer(Index As Integer)
Dim StrProcPtr As String
    StrProcPtr = "tmr" & Index & "_timer"
    If ProcIndex(StrProcPtr) = -1 Then
        StrProcPtr = ""
        Exit Sub
    Else
        Execute ProcGetCodeBlock(StrProcPtr)
        StrProcPtr = ""
    End If
    
End Sub

