VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SM_Debugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enabled As Boolean
Private obControl As Control
Public Property Set OutPutControl(new_Control As Control)
On Error GoTo ERRHDL
    Set obControl = new_Control
    With obControl
        .Text = "> ScriptMagic Debugger initialized" & vbCrLf
    End With
Exit Property
ERRHDL:
MsgBox Err.Description
Enabled = False
End Property
Friend Function GetOPControl() As Object
    Set GetOPControl = obControl
End Function
Public Sub Echo(xVarToDisplay)
Dim buff$
Dim i As VbVarType
On Error GoTo ERRHDL
If Not Enabled Then Exit Sub
buff$ = CStr(xVarToDisplay)
i = VarType(xVarToDisplay)
Select Case i
    Case vbArray
        buff$ = "{Array}"
    Case vbDataObject
        buff$ = "Output = Dataobject"
    Case vbEmpty
        buff$ = "{Empty}"
    Case vbError
        buff$ = xVarToDisplay.Description
    Case vbObject
        buff$ = "{Object}"
    Case vbUserDefinedType
        buff$ = "{User Defined Type}"
    Case vbNull
        buff$ = "{NULL}"
    Case Else
        buff$ = CStr(xVarToDisplay)
End Select
buff$ = "> " & buff$
If Not (obControl Is Nothing) Then
    obControl.SelStart = Len(obControl.Text) + 1
    obControl.SelText = buff$ & vbCrLf
End If
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear
End Sub
Public Sub Alert(strMsg, Optional strTitle)
Dim buff$, buff2$
Dim i As VbVarType
On Error GoTo ERRHDL
If Not Enabled Then Exit Sub
buff$ = CStr(strMsg)
i = VarType(strMsg)
Select Case i
    Case vbArray
        buff$ = "{Array}"
    Case vbDataObject
        buff$ = "Output = Dataobject"
    Case vbEmpty
        buff$ = "{Empty}"
    Case vbError
        buff$ = strMsg.Description
    Case vbObject
        buff$ = "{Object}"
    Case vbUserDefinedType
        buff$ = "{User Defined Type}"
    Case vbNull
        buff$ = "{NULL}"
    Case Else
        buff$ = CStr(strMsg)
End Select
buff2$ = "ScriptMagic Debugger"
If Not IsMissing(strTitle) Then
i = VarType(strTitle)
Select Case i
    Case vbArray
        buff2$ = "{Array}"
    Case vbDataObject
        buff2$ = "Output = Dataobject"
    Case vbEmpty
        buff2$ = "{Empty}"
    Case vbError
        buff2$ = strTitle.Description
    Case vbObject
        buff2$ = "{Object}"
    Case vbUserDefinedType
        buff2$ = "{User Defined Type}"
    Case vbNull
        buff2$ = "{NULL}"
    Case Else
        buff2$ = CStr(strTitle)
End Select
End If
buff2$ = IIf(Len(buff2$) > 50, Left$(buff2$, 50), buff2$)
MsgBox buff$, vbOKOnly + vbCritical, buff2$
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear

End Sub


Private Sub Class_Terminate()
Set obControl = Nothing
End Sub
