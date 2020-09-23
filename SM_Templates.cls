VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SM_Templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private m_TemplateXML As String
Public Enum SMTemplateMethodTypeEnum
    eSubRoutine = 0
    eFunction = 1
End Enum
Public Type SMTemplateMethodType
    SubFunction As SMTemplateMethodTypeEnum
    MethodName As String
    ParmString As String
    Code As String
End Type
Public Enum SMTemplateAddErrorEnum
    eNoError = 0
    eObjectExists = 1
    eInvalidParms = 2
    eInvalidCode = 3
    eInvalidProjectCode = 4
End Enum
Public Function BlankMethod()
Dim ret As SMTemplateMethodType
BlankMethod = ret
End Function
'Public Function AddMethod(methodType As SMTemplateMethodType) As SMTemplateAddErrorEnum
Public Function AddMethod(methodType) As SMTemplateAddErrorEnum
Dim buff$, i As Long
Dim ret As SMTemplateAddErrorEnum
Dim xmlObj As QSXML
Dim nd As Object
Dim ndc As Object
Dim ndp As Object
On Error GoTo ERRHDL
If methodType.SubFunction = eFunction Then
    buff$ = "<FUNCTION></FUNCTION>"
Else
    buff$ = "<SUBROUTINE></SUBROUTINE>"
End If
ret = eInvalidProjectCode
Set xmlObj = New QSXML
xmlObj.Initialize pavAUTO
If Not xmlObj.OpenFromString(m_TemplateXML) Then
    GoTo ERRHDL
End If
With xmlObj
    If methodType.SubFunction = eFunction Then
        Set nd = .GetChildNode(.GetRootChildren, "FUNCTIONS")
    Else
        Set nd = .GetChildNode(.GetRootChildren, "SUBROUTINES")
    End If
    If nd Is Nothing Then
        GoTo ERRHDL
    End If
End With
If Not ValidateParmString(methodType) Then
    AddMethod = eInvalidParms
    Exit Function
End If
If methodType.SubFunction <> eFunction And methodType.SubFunction <> eSubRoutine Then
    AddMethod = eInvalidCode
    Exit Function
End If
If ItemExists(methodType.MethodName) Then
    ret = eObjectExists
    GoTo ERRHDL
End If
With xmlObj
    i = CLng("0" & .GetAttributeValue(nd, "COUNT")) + 1
    Set ndc = .XMLAddNode(nd, buff$)
    .SetAttribute ndc, "NAME", methodType.MethodName
    .SetAttribute ndc, "PARAMETERS", methodType.ParmString
    .SetAttribute ndc, "SCOPE", "Public"
    .SetAttribute ndc, "id", ""
    ndc.Text = methodType.Code
    .SetAttribute nd, "COUNT", CStr(i)
    m_TemplateXML = .XML
End With
AddMethod = eNoError
Exit Function
ERRHDL:
Err.Clear
Set xmlObj = Nothing
AddMethod = ret
End Function
Public Function GenerateShellProject(Optional NewProjectName As String = "New Project") As String
Dim buff$
Dim xmlObj As QSXML
Dim nd As Object
Dim ndc As Object
Dim ndp As Object
Set xmlObj = New QSXML
xmlObj.Initialize pavAUTO
buff$ = "<SCRIPTMAGIC></SCRIPTMAGIC>"
With xmlObj
    .OpenFromString buff$
    Set nd = .GetRootElement()
    .SetAttribute nd, "NAME", NewProjectName
    .SetAttribute nd, "HOSTID", ""
    .SetAttribute nd, "RUNMODE", "INTERACTIVE"
    .SetAttribute nd, "PASSWORD", ""
    .SetAttribute nd, "TIMEOUT", "10"
    .SetAttribute nd, "EXPLICIT", "0"
    .SetAttribute nd, "CREATED", Format$(Now, "dd mmm yyyy hh:nn AMPM")
    .SetAttribute nd, "CREATEDBY", "SYSTEM"
    .SetAttribute nd, "LASTMODIFIED", ""
    .SetAttribute nd, "LASTMODIFIEDBY", ""
    Set ndc = .AddNode(nd, "", "DESCRIPTION")
       ndc.Text = "'Script Magic Project: " & NewProjectName
    Set ndc = .AddNode(nd, "", "CONSTANTS")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "VARIABLES")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "INPUT")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "INITIALIZATION")
        buff$ = "'Step 1 Call the Main Processing Sub Routine" & vbLf
        buff$ = buff$ & "'Call Main() " & vbLf & vbLf
        ndc.Text = buff$
    Set ndc = .AddNode(nd, "", "SUBROUTINES")
        .SetAttribute ndc, "COUNT", "1"
    Set ndp = .AddNode(ndc, "", "SUBROUTINE")
        .SetAttribute ndp, "NAME", "Main"
        .SetAttribute ndp, "PARAMETERS", ""
        .SetAttribute ndp, "SCOPE", "Public"
        ndp.Text = "'Add code here"
    Set ndc = .AddNode(nd, "", "FUNCTIONS")
        .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "CLASSES")
        .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "VBSCRIPT")
m_TemplateXML = .XML
GenerateShellProject = m_TemplateXML
End With
Set xmlObj = Nothing
Exit Function
End Function
Public Property Get SM_TemplateProject() As String
    SM_TemplateProject = m_TemplateXML
End Property
Public Property Let SM_TemplateProject(new_SMProject As String)
m_TemplateXML = new_SMProject
End Property
Private Function ValidateParmString(mtObj) As Boolean
'mtObj declared as a variant but should be a MethodType structure
    Dim i As Long, j As Long
    ReDim ed1(0) As String
    With mtObj
        If .ParmString = "" Then
            ValidateParmString = True
            Exit Function
        End If
        ed1 = Split(.ParmString, ",")
        For j = 0 To UBound(ed1)
            If ed1(j) = "" Then
                ValidateParmString = False
                Exit Function
            End If
            If AlphaNumFormat(ed1(j), "_") <> ed1(j) Then
                ValidateParmString = False
                Exit Function
            End If
            If InStr(CALPHA, UCase$(Left$(ed1(j), 1))) = 0 Then
                ValidateParmString = False
                Exit Function
            End If
        Next
        ValidateParmString = True
    End With
End Function

Private Sub Class_Initialize()
InitGlobals
GenerateShellProject "New ScriptMagic Project"
End Sub
Private Function ItemExists(itmName As String) As Boolean
Dim nd As Object
Dim ndl As Object
Dim i As Long, j As Long
Dim locXML As New QSXML
locXML.Initialize pavAUTO
locXML.OpenFromString m_TemplateXML
Dim rootl As Object
If InStr(SM_BUILTINOBJECTS, " " & UCase$(itmName) & " ") > 0 Then
    'Msgbox itmName & " is the name of a built in ScriptMagic object", vbCritical, "Error.."
    ItemExists = True
            Set locXML = Nothing
    Exit Function
End If
If InStr(UCase$(SM_RESERVEDWORDS), " " & UCase$(itmName) & " ") > 0 Then
    'Msgbox itmName & " is a vbscript reserved word", vbCritical, "Error.."
    ItemExists = True
            Set locXML = Nothing
    Exit Function
End If
If InStr(UCase$(SM_FUNCTIONCONST), " " & UCase$(itmName) & " ") > 0 Then
    'Msgbox itmName & " is the name of a vbscript function", vbCritical, "Error.."
    ItemExists = True
            Set locXML = Nothing
    Exit Function
End If
With locXML
Set rootl = .GetRootChildren()
If itmName = "" Then
    ''Msgbox "Invalid item name", vbCritical, "ItemExists()"
    ItemExists = True
            Set locXML = Nothing
    Exit Function
End If
For i = 0 To rootl.Length - 1
    If CLng("0" & .GetAttributeValue(rootl(i), "COUNT")) > 0 Then
        Set ndl = .GetChildNodeList(rootl(i))
        For j = 0 To ndl.Length - 1
            If UCase$(.GetAttributeValue(ndl(j), "NAME")) = UCase$(itmName) Then
            ItemExists = True
            Set locXML = Nothing
            Exit Function
            End If
        Next
    End If
Next
End With
ItemExists = False
Set locXML = Nothing
Exit Function
End Function


