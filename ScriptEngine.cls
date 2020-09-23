VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScriptEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const m_VersionMajor = 1
Private Const m_VersionMinor = 0
Private Type MyRequiredObjectsType
    Active As Boolean
    objScriptName As String
    obj As Object
End Type
Public Enum SMProjectOpenMode
    OpenFile = 0
    OpenString = 1
End Enum
Private SystemObjects() As MyRequiredObjectsType
Private ProjectObjects() As MyRequiredObjectsType
Private m_OpenMode As SMProjectOpenMode

Public Enum SMReadyStateType
    Waiting = -1
    Initialized = 0
    Ready = 1
    Executing = 2
End Enum
Private m_ReadyState As SMReadyStateType

Public Type SMLastErrorType
    ErrorNumber As Long
    ScriptLineNumber As Long
    ErrorDescription As String
    ErrorSource As String
    OffendingCode As String
End Type
Private m_LastError As SMLastErrorType
Private m_ParentForm As Object
Private m_VBScript As String
Private m_ProjectPassWord As String
Private m_Tag As String
Private m_Verbose As Boolean
Private m_Initialized As Boolean
Private WithEvents MyProjectEvents As SM_Event
Attribute MyProjectEvents.VB_VarHelpID = -1
Private MyProjectDebugger As SM_Debugger
'Private MyScriptHostFrm As FRMScriptCTLHost
'Private SC1 As ScriptControl
Private SC1
Private MyProjectXML As QSXML
Private MyErrorTextBox As Object
Private MyEventTextBox As Object
Private MySaveReqTextBox As Object
Private MyReadyStateTextBox As Object
'***********Event Declarations***********
Event SaveRequest(strScriptMagic As String)
Event ScriptEvent(evtName As String, evtParms As String)
Event SMError(smErr As SMLastErrorType)
Event SMReadyStateChange(smReadyState As SMReadyStateType)
'***************Properties***************
Public Property Let ReadyStateTextBox(newEVTBox As Object)
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End If
Set MyReadyStateTextBox = newEVTBox
MyReadyStateTextBox.Text = ""
Exit Property
ERRHDL:
    Exit Property
End Property
Public Property Let SaveRequestTextBox(newEVTBox As Object)
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End If
Set MySaveReqTextBox = newEVTBox
MySaveReqTextBox.Text = ""
Exit Property
ERRHDL:
    Exit Property
End Property
Public Property Let EventTextBox(newEVTBox As Object)
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End If
Set MyEventTextBox = newEVTBox
MyEventTextBox.Text = ""
Exit Property
ERRHDL:
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End Property
Public Property Let ErrorTextBox(newErrorTextBox As Object)
On Error GoTo ERRHDL
If Not m_Initialized Then
    'in this case for this error we do a message box
    MsgBox "Script magic not initialized", vbCritical, "Error.."
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End If
Set MyErrorTextBox = newErrorTextBox
MyErrorTextBox.Text = ""
Exit Property
ERRHDL:
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End Property
Public Property Let EchoTextBox(newEchoCtl As Object)
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Property
End If
Set MyProjectDebugger.OutPutControl = newEchoCtl
MyProjectDebugger.Enabled = True
End Property
Public Property Get HasParentForm() As Boolean
HasParentForm = IIf(m_ParentForm Is Nothing, False, True)
End Property
Public Property Get ScriptSystemObjectCount() As Long
    ScriptSystemObjectCount = UBound(SystemObjects)
End Property
Public Property Get ScriptSystemObjectName(Index As Long) As String
    If Index <= 0 Or Index > UBound(SystemObjects) Then
        ScriptSystemObjectName = ""
    Else
        If SystemObjects(Index).Active Then
        ScriptSystemObjectName = SystemObjects(Index).objScriptName
        Else
        ScriptSystemObjectName = ""
        End If
    End If
End Property
Public Property Get ScriptUserObjectCount() As Long
    ScriptUserObjectCount = UBound(ProjectObjects)
End Property
Public Property Get ScriptUserObjectName(Index As Long) As String
    If Index <= 0 Or Index > UBound(ProjectObjects) Then
        ScriptUserObjectName = ""
    Else
        If ProjectObjects(Index).Active Then
        ScriptUserObjectName = ProjectObjects(Index).objScriptName
        Else
        ScriptUserObjectName = ""
        End If
    End If
End Property

Public Property Get ParentForm() As Object
    Set ParentForm = m_ParentForm
End Property
Public Property Let ParentForm(new_ParentForm As Object)
    Set m_ParentForm = new_ParentForm
End Property
Public Property Get MSVBScriptObject() As Object
    Set MSVBScriptObject = SC1
End Property
Public Property Get Tag() As String
    Tag = m_Tag
End Property
Public Property Let Tag(new_Tag As String)
    m_Tag = new_Tag
End Property
Public Property Get SM_VBScript() As String
    SM_VBScript = m_VBScript
End Property
Public Property Get EchoEnabled() As Boolean
    EchoEnabled = MyProjectDebugger.Enabled
End Property
Public Property Get XMLString() As String
If Not m_Initialized Then
    RaiseSMError 452342, "Script Magic Not Initialized"
    Exit Property
End If
If m_ReadyState = Waiting Then
    RaiseSMError 452342, "Invalid Ready State", "XMLString"
    Exit Property
End If
XMLString = MyProjectXML.XML
Exit Property
End Property
Public Property Let Password(newPassword As String)
    m_ProjectPassWord = newPassword
End Property
Public Property Get verbose() As Boolean
    verbose = m_Verbose
End Property
Public Property Let verbose(bNewVerbose As Boolean)
    m_Verbose = bNewVerbose
End Property
Public Property Get OpenMode() As SMProjectOpenMode
    OpenMode = m_OpenMode
End Property
Public Property Get XMLObject() As Object
    Set XMLObject = MyProjectXML
End Property
Public Property Get LastError() As SMLastErrorType
    LastError = m_LastError
End Property
Public Property Get ReadyState() As SMReadyStateType
    ReadyState = m_ReadyState
End Property
Private Property Let ReadyState(new_ReadyState As SMReadyStateType)
    m_ReadyState = new_ReadyState
    MyProjectDebugger.Echo "Project ready state changed to " & new_ReadyState
    RaiseEvent SMReadyStateChange(m_ReadyState)
    If Not (MyReadyStateTextBox Is Nothing) Then
        With MyReadyStateTextBox
           .Tag = CStr(new_ReadyState)
           If new_ReadyState = Executing Then
            .Text = "Executing"
           ElseIf new_ReadyState = Initialized Then
            .Text = "Initialized"
           ElseIf new_ReadyState = Ready Then
            .Text = "Ready"
           ElseIf new_ReadyState = Waiting Then
            .Text = "Waiting"
           Else
            .Text = "Unknown"
           End If
        End With
    End If
End Property
Public Property Get Version() As String
    Version = CStr(m_VersionMajor) & "." & CStr(m_VersionMinor)
End Property
'***************Public Methods***************
Public Function GenerateProjectVBScript(Optional Compacted As Boolean = False) As String
Dim buff$, i As Long
On Error GoTo ERRHDL
If Not m_Initialized Then
    GenerateProjectVBScript = ""
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Initialized Then
    GenerateProjectVBScript = ""
    RaiseSMError -2005, "No Script Magic Project Loaded"
    Exit Function
End If
GenerateProjectVBScript = ExtractVBScript(MyProjectXML, Compacted)
Exit Function
ERRHDL:
RaiseSMError Err.Number, Err.Description, "Generate Project Script"
Err.Clear
GenerateProjectVBScript = ""

End Function
Public Function AddScriptObject(strObjectName As String, smScriptobj As Object) As Boolean
Dim buff$, i As Long
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Waiting Then
    RaiseSMError -2005, "No Script Magic Project Loaded"
    Exit Function
End If
If strObjectName = "" Then
    RaiseSMError -2012, "Invalid object name"
    Exit Function
End If
If Not isValidObjName(buff$, verbose) Then
    RaiseSMError -2012, "Invalid object name"
    Exit Function
End If

For i = 0 To UBound(ProjectObjects) - 1
    With ProjectObjects(i)
        If UCase$(.objScriptName) = strObjectName Then
            Set .obj = smScriptobj
            .Active = True
            AddScriptObject = True
            Exit Function
        End If
    End With
Next
i = UBound(ProjectObjects) + 1
With ProjectObjects(i)
    .Active = True
    .objScriptName = strObjectName
    Set .obj = smScriptobj
End With
AddScriptObject = True
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    AddScriptObject = False
    Exit Function
End If
If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, "GetScriptValue"
    AddScriptObject = False
End If
Err.Clear

End Function
Public Function ClearScriptObjects() As Boolean
    ReDim ProjectObjects(0) As MyRequiredObjectsType
End Function

Public Function GetScriptValue(varName As String) As Variant
Dim buff$
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Initialized Then
    RaiseSMError -2005, "No Script Magic Project Loaded"
    Exit Function
End If
buff$ = varName
GetScriptValue = SC1.Eval(buff$)
Exit Function
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Function
End If
If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, "GetScriptValue"
End If
Err.Clear
End Function
Public Function GetScriptObject(varObjName As String) As Object
Dim buff$
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Initialized Then
    RaiseSMError -2005, "No Script Magic Project Loaded"
    Exit Function
End If
buff$ = varObjName
Set GetScriptObject = SC1.Eval(buff$)
Exit Function
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Function
End If

If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, "GetScriptValue"
End If
Err.Clear
Set GetScriptObject = Nothing
End Function

Public Function IsMethod(strMethodName As String) As Boolean
Dim nd As Object
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
Set nd = GetItemNode(strMethodName)
If (nd Is Nothing) Then
    IsMethod = False
Else
    If InStr("FUNCTIONSUBROUTINE", UCase$(nd.nodename)) = 0 Then
        IsMethod = False
    Else
        IsMethod = True
    End If
End If
Set nd = Nothing
End Function
Public Function GetMethod(strMethodName As String) As SM_ScriptMethod
Dim ret As SM_ScriptMethod
Dim nd As Object
Dim buff$
ReDim ed1(0) As String
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
    
With MyProjectXML
    If Not IsMethod(strMethodName) Then
        Set GetMethod = ret  'Nothing
        Exit Function
    End If
    Set nd = GetItemNode(strMethodName)
    Set ret = New SM_ScriptMethod
    buff$ = .GetAttributeValue(nd, "PARAMETERS")
    If buff$ = "" Then
        ret.SetMethodName strMethodName
    Else
        ed1 = Split(buff$, ",")
        ret.SetMethodName strMethodName, UBound(ed1) + 1
    End If
End With
Set GetMethod = ret
End Function
Public Sub ExecuteScript()
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Sub
End If

If m_VBScript = "" Then
    RaiseSMError 683442, "No Executable Script", "Script Magic {ExecuteScript}"
     Exit Sub
End If
If m_ReadyState <> Ready Then
    ResetProject
    If m_LastError.ErrorNumber <> 0 Then
        Exit Sub
    End If
End If
ReadyState = Executing
With SC1
.AddCode m_VBScript
End With
Exit Sub
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Sub
End If

If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, Err.Source
End If
Err.Clear
End Sub
Public Sub ExecuteScriptDebug()
On Error GoTo ERRHDL
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Sub
End If

If m_VBScript = "" Then
    RaiseSMError 683442, "No Executable Script", "Script Magic {ExecuteScript}"
     Exit Sub
End If
If m_ReadyState <> Ready Then
    ResetProject
    If m_LastError.ErrorNumber <> 0 Then
        Exit Sub
    End If
End If
ReadyState = Executing
With SC1
m_VBScript = ExtractVBScript(MyProjectXML, False)
.AddCode m_VBScript
End With
Exit Sub
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Sub
End If

If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, Err.Source
End If
Err.Clear
End Sub

Public Function RunSub(strMethodName As String, ParamArray MethodParms() As Variant) As Boolean
Dim i As Long, buff$
Dim objMethod As SM_ScriptMethod

On Error GoTo ERRHDL
ClearErrors
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Initialized Then
    RaiseSMError 52634, "An active Script Magic Script is not running", "ExecuteScriptMethod()"
    RunSub = False
    Exit Function
End If
If strMethodName = "" Then
    RaiseSMError 52634, "An active Script Magic Script is not running", "ExecuteScriptMethod()"
    RunSub = False
    Exit Function
End If
Set objMethod = New SM_ScriptMethod
With objMethod
    If IsMissing(MethodParms) Then
        .SetMethodName strMethodName, 0, False
    Else
        .SetMethodName strMethodName, UBound(MethodParms) + 1, False
    End If
    If Not IsMissing(MethodParms) Then
        For i = 0 To UBound(MethodParms)
            Select Case VarType(MethodParms(i))
                Case vbEmpty
                    .AddParameter i + 1, ParmScriptObj, "vbEmpty"
                Case vbNull
                    .AddParameter i + 1, ParmScriptObj, "vbNull"
                Case vbInteger, vbLong, vbCurrency, vbSingle, vbDouble, vbDecimal
                    .AddParameter i + 1, parmNumber, CStr(MethodParms(i))
                Case vbString, vbByte
                    .AddParameter i + 1, ParmString, CStr(MethodParms(i))
                Case vbDate
                    .AddParameter i + 1, ParmString, CStr(MethodParms(i) & "")
                Case vbBoolean
                    .AddParameter i + 1, parmNumber, CStr(CLng(MethodParms(i)))
                Case vbArray
                    RaiseSMError -6001, "Invalid Parameter Type: Array()"
                    RunSub = False
                    Set objMethod = Nothing
                    Exit Function
                Case vbObject
                    RaiseSMError -6001, "Invalid Parameter Type: {Object}"
                    RunSub = False
                    Set objMethod = Nothing
                    Exit Function
                Case vbUserDefinedType
                    RaiseSMError -6001, "Invalid Parameter Type: {user defined type}"
                    RunSub = False
                    Set objMethod = Nothing
                    Exit Function
                Case Else
                    RaiseSMError -6001, "Invalid Parameter Type: " & VarType(MethodParms(i))
                    RunSub = False
                    Set objMethod = Nothing
                    Exit Function
            End Select
        Next
    End If
End With
RunSub = ExecuteScriptMethod(objMethod)
Set objMethod = Nothing
Exit Function
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    RunSub = False
    Exit Function
End If
If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, Err.Source
End If
Err.Clear
RunSub = False
End Function
Public Function RunFunction(strMethodName As String, ParamArray MethodParms() As Variant) As Variant
Dim i As Long, buff$
Dim objMethod As SM_ScriptMethod

On Error GoTo ERRHDL
ClearErrors
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
If m_ReadyState = Initialized Then
    RaiseSMError 52634, "An active Script Magic Script is not running", "ExecuteScriptMethod()"
    RunFunction = False
    Exit Function
End If
If strMethodName = "" Then
    RaiseSMError 52634, "An active Script Magic Script is not running", "ExecuteScriptMethod()"
    RunFunction = False
    Exit Function
End If
Set objMethod = New SM_ScriptMethod
With objMethod
    If IsMissing(MethodParms) Then
        .SetMethodName strMethodName, 0, True
    Else
        .SetMethodName strMethodName, UBound(MethodParms) + 1, True
    End If
    If Not IsMissing(MethodParms) Then
        For i = 0 To UBound(MethodParms)
            Select Case VarType(MethodParms(i))
                Case vbEmpty
                    .AddParameter i + 1, ParmScriptObj, "vbEmpty"
                Case vbNull
                    .AddParameter i + 1, ParmScriptObj, "vbNull"
                Case vbInteger, vbLong, vbCurrency, vbSingle, vbDouble, vbDecimal
                    .AddParameter i + 1, parmNumber, CStr(MethodParms(i))
                Case vbString, vbByte
                    .AddParameter i + 1, ParmString, CStr(MethodParms(i))
                Case vbDate
                    .AddParameter i + 1, ParmString, CStr(MethodParms(i) & "")
                Case vbBoolean
                    .AddParameter i + 1, parmNumber, CStr(CLng(MethodParms(i)))
                Case vbArray
                    RaiseSMError -6001, "Invalid Parameter Type: Array()"
                    RunFunction = False
                    Set objMethod = Nothing
                    Exit Function
                Case vbObject
                    RaiseSMError -6001, "Invalid Parameter Type: {Object}"
                    RunFunction = False
                    Set objMethod = Nothing
                    Exit Function
                Case vbUserDefinedType
                    RaiseSMError -6001, "Invalid Parameter Type: {user defined type}"
                    RunFunction = False
                    Set objMethod = Nothing
                    Exit Function
                Case Else
                    RaiseSMError -6001, "Invalid Parameter Type: " & VarType(MethodParms(i))
                    RunFunction = False
                    Set objMethod = Nothing
                    Exit Function
            End Select
        Next
    End If
End With
If ExecuteScriptMethod(objMethod) Then
    RunFunction = objMethod.ReturnValue
End If
Set objMethod = Nothing
Exit Function
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Function
End If

If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, Err.Source
End If
Err.Clear

End Function


Public Function ExecuteScriptMethod(objMethod As SM_ScriptMethod) As Boolean
Dim buff$, i As Long
Dim prm As SMParameterType
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Function
End If
On Error GoTo ERRHDL
ClearErrors
If m_ReadyState = Initialized Then
    RaiseSMError 52634, "An active Script Magic Script is not running", "ExecuteScriptMethod()"
    ExecuteScriptMethod = False
    Exit Function
End If
If Not objMethod.Ready Then
    RaiseSMError 423674, "Method: {" & objMethod.MethodName & "} not set correctly", "ExecuteScriptMethod"
    Exit Function
End If
With objMethod
    If .MethodReturnsValue Then
        buff$ = .MethodName
    Else
        buff$ = "Call " & .MethodName
    End If
    If .ParmsRequired > 0 Then
        buff$ = buff$ & "("
        For i = 1 To .smParameters.Count
            prm = .smParameters.Item(i)
            If i = 1 Then
                buff$ = buff$ & prm.ParmValue
            Else
                buff$ = buff$ & "," & prm.ParmValue
            End If
        Next
        buff$ = buff$ & ")"
    End If
    .ReturnValue = ""
    If Not .MethodReturnsValue Then
        SC1.ExecuteStatement buff$
    Else
        .ReturnValue = SC1.Eval(buff$)
    End If
End With
ExecuteScriptMethod = True
Exit Function
ERRHDL:
If SC1.Error.Number <> 0 Then
    SC1_Error
    Err.Clear
    RaiseEvent SMError(m_LastError)
    If Not (MyErrorTextBox Is Nothing) Then
        RaiseTextBoxError m_LastError
    End If
    
    Exit Function
End If
If m_LastError.ErrorNumber <> Err.Number Then
    RaiseSMError Err.Number, Err.Description, Err.Source
End If
Err.Clear

End Function
Public Sub Halt()
On Error GoTo ERRHDL:
    ClearErrors
    With SC1
        .Reset
        ReadyState = Waiting
    End With
    Exit Sub
ERRHDL:
RaiseSMError Err.Number, Err.Description, Err.Source
Err.Clear
End Sub

Public Sub ResetProject()
Dim i As Long
On Error GoTo ERRHDL:
If Not m_Initialized Then
    RaiseSMError -2004, "Script Magic Object Not Initialized"
    Exit Sub
End If
If m_ReadyState <> Waiting Then
    ReadyState = Initialized
End If
With SC1
    .Reset
    If Not (m_ParentForm Is Nothing) Then
        .AddObject "ParentForm", m_ParentForm
    End If
    For i = 1 To UBound(SystemObjects)
        If SystemObjects(i).Active Then
            .AddObject SystemObjects(i).objScriptName, SystemObjects(i).obj
        End If
    Next
    If UBound(ProjectObjects) > 0 Then
        For i = 1 To UBound(ProjectObjects)
            If ProjectObjects(i).Active Then
                .AddObject ProjectObjects(i).objScriptName, ProjectObjects(i).obj
            End If
        Next
    End If
End With
ClearErrors
Exit Sub
ERRHDL:
RaiseSMError Err.Number, Err.Description, Err.Source
Err.Clear
End Sub

Public Function GetDebuggerControl() As Object
    Set GetDebuggerControl = MyProjectDebugger.GetOPControl()
End Function
Public Function OpenProject(strProject As String, eOpenMode As SMProjectOpenMode) As Boolean
Dim buff$
On Error GoTo ERRHDL:
If Not m_Initialized Then
    RaiseSMError 9234215, "Invalid Ready State", "Script Magic {OpenProject}"
    OpenProject = False
    Exit Function
End If
ResetProject

If eOpenMode = OpenFile Then
    buff$ = OpenScriptMagicFile(strProject, m_ProjectPassWord, m_Verbose)
    If buff$ = "" Then
        OpenProject = False
    End If
    With MyProjectXML
        If Not .OpenFromString(buff$, m_Verbose) Then
            RaiseSMError 34673243, "Error loading project", "OpenProject()"
            OpenProject = False
            Exit Function
        End If
    End With
Else
    buff$ = OpenScriptMagicString(strProject, m_ProjectPassWord, m_Verbose)
    If buff$ = "" Then
        OpenProject = False
        Exit Function
    End If
    With MyProjectXML
        If Not .OpenFromString(buff$, m_Verbose) Then
            RaiseSMError 34673243, "Error loading project", "OpenProject()"
            OpenProject = False
            Exit Function
        End If
    End With
End If
m_VBScript = CompactVBScript(MyProjectXML)
ReadyState = Ready
OpenProject = True
Exit Function
ERRHDL:
RaiseSMError Err.Number, Err.Description, "OpenProject()"
Err.Clear
OpenProject = False
End Function
Public Function ExportScriptToFile(strFilename As String) As Boolean
If Not m_Initialized Then
    RaiseSMError 9234215, "Invalid Ready State", "Script Magic {ExportScript}"
    ExportScriptToFile = False
    Exit Function
End If
If m_ReadyState = Waiting Then
    RaiseSMError 9234215, "Invalid Ready State", "Script Magic {Export Script}"
    ExportScriptToFile = False
    Exit Function
End If
    With MyProjectXML
        ExportScriptToFile = .Save(strFilename)
    End With
End Function

Public Sub SaveProject()
If Not m_Initialized Then
    RaiseSMError 9234215, "Invalid Ready State", "Script Magic {OpenProject}"
    Exit Sub
End If
RaiseEvent SaveRequest(MyProjectXML.XML)
If Not (MySaveReqTextBox Is Nothing) Then
    With MySaveReqTextBox
        .Text = MyProjectXML.XML
        .Tag = "SAVEREQUEST"
    End With
End If
End Sub
Public Sub Echo(xmsgStr As String)
On Error GoTo ERRHDL
MyProjectDebugger.Echo xmsgStr
Exit Sub
ERRHDL:
Err.Clear
End Sub
Public Sub ShowEditor()
On Error GoTo ERRHDL
Dim x As New Editor
If ReadyState = Waiting Then
    RaiseSMError -34234, "Invalid ready state to show editor", "ShowEditor()"
    Exit Sub
End If
x.Show Me
Set x = Nothing

Exit Sub
ERRHDL:
RaiseSMError Err.Number, Err.Description, "{ShowEditor()}"
Err.Clear
End Sub
Public Sub RaiseSMError(eNumber As Long, Optional eDescription As String = "", _
                         Optional eSource As String = "", Optional eLineNum As Long = 0)
Dim eCode As String

With m_LastError
    .ErrorNumber = eNumber
    .ErrorDescription = eDescription
    .ErrorSource = eSource
    .ScriptLineNumber = eLineNum
    If eLineNum <> 0 Then
        eCode = GetCodeLine(eLineNum)
    End If
    .OffendingCode = eCode
End With
RaiseEvent SMError(m_LastError)
If Not (MyErrorTextBox Is Nothing) Then
    RaiseTextBoxError m_LastError
End If
End Sub
Private Sub RaiseTextBoxError(mErr As SMLastErrorType)
On Error GoTo ERRHDL
With MyErrorTextBox
    .Text = CStr(mErr.ErrorNumber) & Chr$(176) & mErr.ScriptLineNumber & _
    Chr$(176) & mErr.ErrorDescription & Chr$(176) & _
    mErr.ErrorSource & Chr$(176) & mErr.OffendingCode & Chr$(176)
    .Tag = CStr(mErr.ErrorNumber)
End With
Exit Sub
ERRHDL:
Err.Clear
End Sub
Public Sub ClearErrors()
With m_LastError
    .ErrorNumber = 0
    .ErrorDescription = ""
    .ErrorSource = ""
    .OffendingCode = ""
    .ScriptLineNumber = 0
End With
If Not (MyErrorTextBox Is Nothing) Then
    MyErrorTextBox.Text = ""
    MyErrorTextBox.Tag = ""
End If
End Sub
'****************Constructor / Destructor********************
Private Sub Class_Initialize()
On Error GoTo ERRHDL
m_Initialized = True
Set MyProjectDebugger = New SM_Debugger
Set MyProjectEvents = New SM_Event
ReDim SystemObjects(2) As MyRequiredObjectsType
ReDim ProjectObjects(0) As MyRequiredObjectsType
With SystemObjects(1)
    .Active = True
    .objScriptName = "SMDebug"
    Set .obj = MyProjectDebugger
End With
With SystemObjects(2)
    .Active = True
    .objScriptName = "SMEvent"
    Set .obj = MyProjectEvents
End With
Set MyProjectXML = New QSXML
Set SC1 = CreateObject("ScriptControl")
SC1.Language = "VBScript"
m_ReadyState = Waiting
SC1.Reset
'Set MyScriptHostFrm = New FRMScriptCTLHost
'Set MyScriptHostFrm.MyHostObj = Me
'MyScriptHostFrm.Visible = False
'Set SC1 = MyScriptHostFrm.GetScriptControl()

MyProjectXML.Initialize pavAUTO
Exit Sub
ERRHDL:
m_Initialized = False
End Sub
Private Sub Class_Terminate()
Set MyProjectEvents = Nothing
Set MyProjectDebugger = Nothing
Set MyProjectXML = Nothing
'Unload FRMScriptCTLHost
'Set FRMScriptCTLHost = Nothing
Set SC1 = Nothing
Erase SystemObjects
Erase ProjectObjects
End Sub
'***************Private Methods*****************
Private Function OpenScriptMagicFile(strFilename As String, Optional strPassword As String = "", Optional verbose As Boolean = True) As String
Dim x As QSXML
Dim y As Object
Dim buff$
On Error GoTo ERRHDL
Set x = New QSXML
x.Initialize pavAUTO
If Not x.OpenFromFile(strFilename, verbose) Then
    OpenScriptMagicFile = ""
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
        OpenScriptMagicFile = ""
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
                If verbose Then
                    If Not PromptForPassword(buff$) Then
                        Set x = Nothing
                        OpenScriptMagicFile = ""
                        Exit Function
                    End If
                Else
                    OpenScriptMagicFile = ""
                    Exit Function
                End If
            End If
        End If
        OpenScriptMagicFile = .XML
End With
Set x = Nothing
Exit Function
ERRHDL:
If verbose Then
    MsgBox Err.Description, vbCritical, "OpenScriptMagicFile"
End If
    Err.Clear
OpenScriptMagicFile = ""
End Function
Private Function PromptForPassword(strToMatch As String) As Boolean
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

Private Function OpenScriptMagicString(strFilename As String, Optional strPassword As String = "", Optional verbose As Boolean = True) As String
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
               If verbose Then
                If Not PromptForPassword(buff$) Then
                    Set x = Nothing
                    OpenScriptMagicString = ""
                    Exit Function
                End If
               Else
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
Private Function GetCodeLine(lnNumber As Long) As String
Dim i As Long
i = lnNumber - 1
ReDim ed1(0) As String
GetCodeLine = ""
ed1 = Split(m_VBScript, vbLf)
If i <= UBound(ed1) Then
    GetCodeLine = ed1(i)
End If
Erase ed1
End Function

Private Function GetItemNode(itmName As String, Optional isRootItem As Boolean = False) As Object
Dim nd As Object
Dim ndl As Object
Dim ret As Object
Dim i As Long, j As Long
Dim rootl As Object
With MyProjectXML
Set rootl = .GetRootChildren()
If itmName = "" Then
    MsgBox "Invalid item name", vbCritical, "GetItemNode()"
    Set GetItemNode = ret
    Exit Function
End If
If isRootItem Then
    For i = 0 To rootl.Length - 1
        If UCase$(rootl(i).nodename) = UCase$(itmName) Then
            Set GetItemNode = rootl(i)
            Exit Function
        End If
    Next
    Set GetItemNode = Nothing
    Exit Function
End If
For i = 0 To rootl.Length - 1
    If CLng("0" & .GetAttributeValue(rootl(i), "COUNT")) > 0 Then
        Set ndl = .GetChildNodeList(rootl(i))
        For j = 0 To ndl.Length - 1
            If UCase$(.GetAttributeValue(ndl(j), "NAME")) = UCase$(itmName) Then
            Set GetItemNode = ndl(j)
            Exit Function
            End If
        Next
    End If
Next
End With
    Set GetItemNode = ret
    Exit Function
End Function
Private Function ItemExists(itmName As String) As Boolean
Dim nd As Object
Dim ndl As Object
Dim i As Long, j As Long
Dim rootl As Object
If InStr(SM_BUILTINOBJECTS, " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is the name of a built in ScriptMagic object", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
If InStr(UCase$(SM_RESERVEDWORDS), " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is a vbscript reserved word", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
If InStr(UCase$(SM_FUNCTIONCONST), " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is the name of a vbscript function", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
With MyProjectXML
Set rootl = .GetRootChildren()
If itmName = "" Then
    MsgBox "Invalid item name", vbCritical, "ItemExists()"
    ItemExists = True
    Exit Function
End If
For i = 0 To rootl.Length - 1
    If CLng("0" & .GetAttributeValue(rootl(i), "COUNT")) > 0 Then
        Set ndl = .GetChildNodeList(rootl(i))
        For j = 0 To ndl.Length - 1
            If UCase$(.GetAttributeValue(ndl(j), "NAME")) = UCase$(itmName) Then
            ItemExists = True
            Exit Function
            End If
        Next
    End If
Next
End With
ItemExists = False
Exit Function
End Function

Private Sub SC1_Error()
Dim i As Long
On Error GoTo ERRHDL
With SC1.Error
    m_LastError.ErrorNumber = .Number
    m_LastError.ErrorDescription = .Description
    m_LastError.ScriptLineNumber = .Line
    If .Line > 0 Then
        m_LastError.OffendingCode = GetCodeLine(.Line)
    End If
    .Clear
End With
Exit Sub
ERRHDL:
    m_LastError.ErrorNumber = Err.Number
    m_LastError.ErrorDescription = Err.Description
    m_LastError.ScriptLineNumber = 0
    Err.Clear
'SC1.Reset
End Sub

Private Sub MyProjectEvents_SMScriptEvent(evtName As String, evtParms As String)
If evtParms = "" Then
    MyProjectDebugger.Echo "Event: " & evtName
Else
    MyProjectDebugger.Echo "Event: " & evtName & " Parms: " & evtParms
End If
RaiseEvent ScriptEvent(evtName, evtParms)
If Not (MyEventTextBox Is Nothing) Then
    With MyEventTextBox
        .Text = ""
        .Text = evtName & Chr$(176) & evtParms & Chr$(176)
        .Tag = evtName
    End With
End If
End Sub
Friend Function ExtractVBScript(locQSXML As QSXML, Optional bNoWhiteSpace As Boolean = False) As String
Dim nd As Object
Dim rootNDL As Object
Dim ndc As Object
Dim ndp As Object
Dim ndl As Object
Dim locLn As String
Dim buff$, i As Long

ReDim ed1(0) As String

With locQSXML
    Set nd = .GetRootElement()
    If .GetAttributeValue(nd, "NAME") = "" Then
        ExtractVBScript = ""
        Exit Function
    End If
    buff$ = "'" & Chr$(171) & "Project: " & .GetAttributeValue(nd, "NAME") & Chr$(187) & vbLf
    buff$ = buff$ & "'" & String$(50, "*") & vbLf
    buff$ = buff$ & "'Created: " & .GetAttributeValue(nd, "CREATED") & _
    " Author: " & .GetAttributeValue(nd, "AUTHOR") & vbLf
    buff$ = buff$ & "'Script Magic Run Mode: " & .GetAttributeValue(nd, "RUNMODE") & vbLf & _
    "'Project Description:" & vbLf
    Set rootNDL = .GetChildNodeList(nd)
        If .IsChildNode(nd, "DESCRIPTION") Then
            Set ndc = .GetChildNode(rootNDL, "DESCRIPTION")
            buff$ = buff$ & ndc.Text
        End If
        buff$ = buff$ & vbLf
        If .GetAttributeValue(nd, "EXPLICIT") = "1" Then
        buff$ = buff$ & "Option Explicit" & vbLf
        End If
        buff$ = buff$ & "'" & String$(50, "*") & vbLf & "'" & vbLf
        buff$ = buff$ & GetSectionHeader("Public Constants") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "CONSTANTS")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    buff$ = buff$ & "Public Const " & .GetAttributeValue(ndl(i), "NAME") & _
                    " = "
                    If .GetAttributeValue(ndl(i), "TYPE") = "NUMBER" Then
                    buff$ = buff$ & .GetAttributeValue(ndl(i), "VALUE") & vbLf
                    Else
                    buff$ = buff$ & Dquote(.GetAttributeValue(ndl(i), "VALUE")) & vbLf
                    End If
                Next
            End If
        buff$ = buff$ & GetSectionHeader("Public Variables") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "VARIABLES")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    buff$ = buff$ & "Public " & .GetAttributeValue(ndl(i), "NAME") & vbLf
                Next
            End If
            buff$ = buff$ & GetSectionHeader("Active Global Objects") & vbLf
            buff$ = buff$ & "'SCRIPT MAGIC OBJECTS: SMDebug, SMEvent" & vbLf
            If HasParentForm Then
            buff$ = buff$ & "'OBJECT: ParentForm (should be a reference to the currently active window)" & vbLf
            Else
            buff$ = buff$ & "'WARNING: NO Parent form object set by calling application" & vbLf
            End If
            If ScriptUserObjectCount = 0 Then
                buff$ = buff$ & "'USER OBJECTS: <None>" & vbLf
            Else
                buff$ = buff$ & "'USER OBJECTS: "
                For i = 1 To ScriptUserObjectCount
                    buff$ = buff$ & ScriptUserObjectName(i)
                Next
                buff$ = buff$ & vbLf
            End If
            buff$ = buff$ & GetSectionHeader("Initialization Code") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "INITIALIZATION")
                buff$ = buff$ & ndc.Text & vbLf
            buff$ = buff$ & GetSectionHeader("ALL SubRoutines") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "SUBROUTINES")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    buff$ = buff$ & GetSubHeader(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                    buff$ = buff$ & .GetAttributeValue(ndl(i), "SCOPE") & " Sub " & .GetAttributeValue(ndl(i), "NAME") & _
                    "(" & .GetAttributeValue(ndl(i), "PARAMETERS") & ")" & vbLf & _
                    ndl(i).Text & vbLf & _
                    "End Sub" & vbLf
                    buff$ = buff$ & GetSubFooter(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                Next
            End If
            buff$ = buff$ & GetSectionHeader("ALL Functions") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "FUNCTIONS")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    buff$ = buff$ & GetFunctionHeader(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                    buff$ = buff$ & .GetAttributeValue(ndl(i), "SCOPE") & " Function " & .GetAttributeValue(ndl(i), "NAME") & _
                    "(" & .GetAttributeValue(ndl(i), "PARAMETERS") & ")"
                    If .GetAttributeValue(ndl(i), "RETURNTYPE") <> "" Then
                        buff$ = buff$ & " As " & .GetAttributeValue(ndl(i), "RETURNTYPE") & vbLf
                    Else
                        buff$ = buff$ & vbLf
                    End If
                    buff$ = buff$ & ndl(i).Text & vbLf & "End Function" & vbLf
                    buff$ = buff$ & GetFunctionFooter(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                Next
            End If
            buff$ = buff$ & GetSectionHeader("ALL Classes") & vbLf & vbLf
            Set ndc = .GetChildNode(rootNDL, "CLASSES")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    buff$ = buff$ & GetClassHeader(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                    buff$ = buff$ & "Class " & .GetAttributeValue(ndl(i), "NAME") & vbLf
'                    buff$ = buff$ & .GetAttributeValue(ndl(i), "SCOPE") & " Class " & .GetAttributeValue(ndl(i), "NAME") & vbLf
                    buff$ = buff$ & ndl(i).Text & vbLf & "End Class" & vbLf
                    buff$ = buff$ & GetClassFooter(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                Next
            End If
            buff$ = buff$ & GetSectionHeader("END OF FILE") & vbLf
End With
buff$ = Replace(buff$, Chr$(171), "")
buff$ = Replace(buff$, Chr$(187), "")
If bNoWhiteSpace Then
    locLn = " '" & Chr$(171)
    ed1 = Split(buff$, vbLf)
    buff$ = ""
    For i = 0 To UBound(ed1)
        ed1(i) = Trim$(ed1(i))
        If Len(ed1(i)) > 0 Then
            If Left$(ed1(i), 1) <> "'" Then
                buff$ = buff$ & ed1(i) & locLn & "LN:" & (i + 1) & Chr$(187) & vbLf
            End If
        End If
    Next
End If
DoEvents
buff$ = Replace(buff$, vbLf, vbCrLf)
ExtractVBScript = buff$
End Function


