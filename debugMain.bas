Attribute VB_Name = "debugMain"
Option Explicit
Public Type MyParametersType
    stName As String
    stParms As String
    lParmCount As Long
    stParmValues As String
    bActive As Boolean
End Type

Public MyLocalFunctions() As MyParametersType


Public Function ValidParms(funcName As String) As Boolean
    Dim i As Long, j As Long
    ReDim ed1(0) As String
    i = LocalFuncIDX(funcName)
    If i < 0 Then
        ValidParms = False
        Exit Function
    End If
    With MyLocalFunctions(i)
        If .lParmCount = 0 Then
            ValidParms = True
            Exit Function
        End If
        If .stParmValues = "" Then
            ValidParms = False
            Exit Function
        End If
        ed1 = Split(.stParmValues, Chr$(0))
        If (UBound(ed1) + 1) <> .lParmCount Then
            ValidParms = False
            Exit Function
        End If
        For j = 0 To UBound(ed1)
            If ed1(j) = "" Then
                ValidParms = False
                Exit Function
            End If
        Next
        ValidParms = True
    End With
End Function

Public Function ClearFunctions()
ReDim MyLocalFunctions(0) As MyParametersType
End Function
Public Function RemoveFunction(funcName As String) As Boolean
Dim i As Long
Dim j As Long
Dim b1 As Boolean
ReDim tmp1(0) As MyParametersType
i = LocalFuncIDX(funcName)
If i < 0 Then
    RemoveFunction = True
    Exit Function
End If
MyLocalFunctions(i).bActive = False
If UBound(MyLocalFunctions) = 0 Then
    RemoveFunction = True
    Exit Function
End If
ReDim tmp1(UBound(MyLocalFunctions) - 1) As MyParametersType
j = 0
For i = 0 To UBound(MyLocalFunctions)
    If MyLocalFunctions(i).bActive Then
        tmp1(j) = MyLocalFunctions(i)
        j = j + 1
    End If
Next
ReDim MyLocalFunctions(UBound(tmp1)) As MyParametersType
For i = 0 To UBound(tmp1)
    MyLocalFunctions(i) = tmp1(i)
Next
End Function

Public Function GetParmString(strFuncName As String) As String
Dim i As Long, j As Long
Dim ret As String
ReDim ed1(0) As String
i = LocalFuncIDX(strFuncName)
If i < 0 Then
    MsgBox strFuncName & " is not a valid object", vbCritical, "Error.."
    GetParmString = False
    Exit Function
End If

With MyLocalFunctions(i)
    If .lParmCount = 0 Then
        GetParmString = ""
        Exit Function
    End If
    If .stParmValues = "" Then
        GetParmString = ""
        Exit Function
    End If
    ed1 = Split(.stParmValues, Chr$(0))
    If (UBound(ed1) + 1) <> .lParmCount Then
        GetParmString = ""
        Exit Function
    End If
    For j = 0 To UBound(ed1)
        If ed1(j) = "" Then
            GetParmString = ""
            Exit Function
        End If
        If ret = "" Then
            ret = ed1(j)
        Else
            ret = ret & "," & ed1(j)
        End If
    Next
    GetParmString = ret
    Exit Function
End With
End Function

Public Function AddFunctionName(funcName As String, parmList As String) As Boolean
Dim i As Long
Dim j As Long
ReDim ed1(0) As String
j = LocalFuncIDX(funcName)
If j >= 0 Then
    If parmList = "" Then
            RemoveFunction (funcName)
            DoEvents
    End If
End If
If parmList = "" Then
    AddFunctionName = True
    Exit Function
End If
If j >= 0 Then
        If MyLocalFunctions(j).stName = funcName And MyLocalFunctions(j).stParms = parmList Then
            AddFunctionName = True
            Exit Function
        Else
            RemoveFunction (funcName)
            DoEvents
        End If
End If
i = UBound(MyLocalFunctions) + 1
ReDim Preserve MyLocalFunctions(i) As MyParametersType
With MyLocalFunctions(i)
    .stName = funcName
    .stParms = parmList
    If .stParms = "" Then
        .lParmCount = 0
    Else
        ed1 = Split(.stParms, ",")
        .lParmCount = UBound(ed1) + 1
    End If
    .stParmValues = ""
    .bActive = True
End With

End Function
Public Function GetParmArray(funcName As String, pArray() As Variant) As Boolean
Dim i As Long, j As Long
ReDim pArray(0)
ReDim ed1(0) As String
pArray(0) = 0
i = LocalFuncIDX(funcName)
If i < 0 Then
    MsgBox funcName & " is not a valid object", vbCritical, "Error.."
    GetParmArray = False
    Exit Function
End If
With MyLocalFunctions(i)
    If .lParmCount = 0 Then
        pArray(0) = 0
        GetParmArray = True
        Exit Function
    End If
    If .stParmValues = "" Then
        pArray(0) = .lParmCount
        GetParmArray = False
        Exit Function
    End If
    ed1 = Split(.stParmValues, Chr$(0))
    If (UBound(ed1) + 1) <> .lParmCount Then
        pArray(0) = .lParmCount
        GetParmArray = False
        Exit Function
    End If
    ReDim pArray(UBound(ed1))
    For j = 0 To UBound(ed1)
        If ed1(j) = "" Then
            pArray(j) = .lParmCount
            GetParmArray = False
            Exit Function
        End If
        pArray(j) = ed1(j)
    Next
    GetParmArray = True
    Exit Function
End With
End Function

Public Function InListBox(cbList As Control, strValue As String) As Long
Dim i As Long
For i = 0 To cbList.ListCount - 1
    If UCase(cbList.List(i)) = UCase(strValue) Then
        InListBox = i
        Exit Function
    End If
Next
InListBox = -1
End Function
Public Function BlankClassCode() As String
Dim buff$
buff$ = "'Private Variables" & vbLf & vbLf
buff$ = buff$ & "'Class Constructor" & vbLf
buff$ = buff$ & "Private Sub Class_Initialize()" & _
vbLf & "  'Automatically called when class is created.  Place initialization code here" & vbLf & "End Sub" & vbLf
buff$ = buff$ & "'Public Properties And Methods" & vbLf & vbLf & _
"'Class Destructor" & vbLf
buff$ = buff$ & "Private Sub Class_Terminate()" & _
vbLf & "  'Automatically called when class is Destroyed.  Place any cleanup code here" & vbLf & "End Sub" & vbLf

BlankClassCode = buff$

End Function
Public Function OpenAnyFileURL(hwndParent As Long, strFilename As String) As Boolean
Dim apInst As Long
On Error GoTo ERRHDL
            MouseOn
            apInst = ShellExecute(hwndParent, "Open", strFilename, "", "", 5)
            OpenAnyFileURL = (apInst > 32)
            MouseOff
Exit Function
ERRHDL:
MsgBox Err.Description
Err.Clear
MouseOff
            MouseOff
End Function
Public Function LocalFuncIDX(funcName As String)
Dim i As Long
For i = 0 To UBound(MyLocalFunctions)
    If MyLocalFunctions(i).bActive And MyLocalFunctions(i).stName = funcName Then
        LocalFuncIDX = i
        Exit Function
    End If
Next
LocalFuncIDX = -1
End Function

Public Function PromptForPassword(strToMatch As String) As Boolean
Dim iCnt As Long
Dim buff$
Dim ret$
Dim msg1$
msg1$ = "This script magic project is password protected.  Enter the password below"
iCnt = 0
ret$ = ""
Do While iCnt < 3
ret$ = Trim$(InputBox(msg1$, "Enter Password", ""))
If ret$ = "" Then
    PromptForPassword = False
    Exit Function
End If
iCnt = iCnt + 1
If UCase$(ret$) = UCase$(strToMatch) Then
    PromptForPassword = True
    Exit Function
End If
If iCnt < 3 Then
If MsgBox("Incorrect: Try again?", vbQuestion + vbYesNo, "Project Password") = vbNo Then
    PromptForPassword = False
    Exit Function
End If
End If
Loop
MsgBox "Retry count exceeded.  Access to this Script Magic Project is denied", vbCritical, "Error.."
    PromptForPassword = False
    Exit Function
End Function
Public Function OpenScriptMagicString(strFilename As String, Optional strPassword As String = "", Optional verbose As Boolean = True) As String
Dim x As QSXML
Dim y As Object
Dim buff$
On Error GoTo ERRHDL
Set x = New QSXML
x.Initialize pavAUTO
If Not x.OpenFromString(strFilename, verbose) Then
    OpenScriptMagicString = ""
    Set x = Nothing
    Exit Function
End If
With x
    Set y = .GetRootElement()
    If UCase$(y.nodename) <> "SCRIPTMAGIC" Then
        If verbose Then
            MsgBox "Invalid file format.", vbCritical, "Error.."
        End If
        Set x = Nothing
        OpenScriptMagicString = ""
        Exit Function
    End If
        buff$ = .GetAttributeValue(y, "PASSWORD")
        If buff$ <> "" Then
            buff$ = sm_DecodeText(buff$)
            If strPassword <> "" Then
                If UCase$(buff$) <> UCase$(strPassword) Then
                    If verbose Then
                        MsgBox "Invalid password for this project"
                    End If
                End If
            Else
                If Not PromptForPassword(buff$) Then
                    Set x = Nothing
                    OpenScriptMagicString = ""
                    Exit Function
                End If
            End If
        End If
        OpenScriptMagicString = .XML
End With
Set x = Nothing
Exit Function
ERRHDL:
If verbose Then
    MsgBox Err.Description, vbCritical, "OpenScriptMagicString"
End If
    Err.Clear
OpenScriptMagicString = ""
End Function

Public Sub MouseOn()
Screen.MousePointer = vbHourglass
End Sub
Public Sub MouseOff()
Screen.MousePointer = vbDefault
End Sub
Public Function ValidateParameterList(sList As String) As Boolean
Dim buff$, i As Long, j As Long
ReDim ed1(0) As String
buff$ = Trim$(sList)
If buff$ = "" Then
    ValidateParameterList = True
    Exit Function
End If
ed1 = Split(buff$, ",")
For i = 0 To UBound(ed1)
    ed1(i) = Trim$(ed1(i))
    If ed1(i) = "" Then
        MsgBox "No empty parameters allowed", vbCritical, "ValidateParameters()"
        ValidateParameterList = False
        Exit Function
    End If
    If InStr(ed1(i), " ") > 0 Then
        MsgBox "Parameter names may not contain spaces", vbCritical, "ValidateParameters()"
        ValidateParameterList = False
        Exit Function
    End If
    If InStr(CALPHA, UCase$(Left$(ed1(i), 1))) = 0 Then
        MsgBox "Parameter names must begin with A-Z or a-z", vbCritical, "ValidateParameters()"
        ValidateParameterList = False
        Exit Function
    End If
    If AlphaNumFormat(ed1(i), "_") <> ed1(i) Then
        MsgBox "Parameter names cannot contain punctuation or special characters (except _).", vbCritical, "ValidateParameters()"
        ValidateParameterList = False
        Exit Function
    End If
Next
If UBound(ed1) > 0 Then
   For i = 0 To UBound(ed1)
        For j = 0 To UBound(ed1)
            If j <> i Then
                If UCase$(ed1(i)) = UCase$(ed1(j)) Then
                    MsgBox "Duplicate parameter names are not allowed.", vbCritical, "ValidateParameters()"
                    ValidateParameterList = False
                    Exit Function
                End If
            End If
        Next
   Next
End If
ValidateParameterList = True
End Function
Public Function FormatParameterList(sList As String) As String
Dim buff$, ret As String, i As Long
ReDim ed1(0) As String
buff$ = Trim$(sList)
If buff$ = "" Then
    FormatParameterList = ""
    Exit Function
End If
ed1 = Split(buff$, ",")
For i = 0 To UBound(ed1)
    ed1(i) = AlphaNumFormat(Trim$(ed1(i)), "_")
    If i = 0 Then
    ret = ed1(i)
    Else
        ret = ret & ", " & ed1(i)
    End If
Next
FormatParameterList = ret
End Function

