VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SM_ScriptMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum SMParameterTypeEnum
    ParmString = 0
    parmNumber = 1
    ParmScriptObj = 2
End Enum

Public Type SMParameterType
    ParmValid As Boolean
    ParmType As Long
    ParmValue As String
End Type
Private m_ReturnValue As Variant
Private m_MethodName
Private m_ReturnsValue As Boolean
Private m_MethodParms As Collection
Private m_ParmsRequired As Long
Private m_Ready As Boolean
Private m_Parms As Collection
Public Property Get ReturnValue() As Variant
    ReturnValue = m_ReturnValue
End Property
Public Property Let ReturnValue(new_return As Variant)
    m_ReturnValue = new_return
End Property
Public Property Let MethodReturnsValue(new_RetValue As Boolean)
    m_ReturnsValue = new_RetValue
End Property
Public Property Get MethodReturnsValue() As Boolean
    MethodReturnsValue = m_ReturnsValue
End Property
Public Property Get smParameters() As Collection
    Set smParameters = m_Parms
End Property
Public Property Get MethodName() As String
    MethodName = m_MethodName
End Property
Public Property Get Ready() As Boolean
    Ready = m_Ready
End Property
Public Property Get ParmsRequired() As Long
    ParmsRequired = m_ParmsRequired
End Property
Public Sub SetMethodName(new_MethodName As String, Optional lNumberofParms As Long = 0, Optional bReturnsValue As Boolean = False)
    Dim prmType As SMParameterType
    Dim i As Long
    m_ReturnsValue = bReturnsValue
    m_MethodName = new_MethodName
    m_ParmsRequired = lNumberofParms
    ClearParameters
    With m_Parms
        Do While .Count < lNumberofParms
            .Add prmType
        Loop
    End With
    m_Ready = IsReady()
End Sub
Public Sub ClearParameters()
Dim i As Long
Dim prmType As SMParameterType
With m_Parms
    Do While .Count > 0
        .Remove 1
    Loop
    For i = 1 To m_ParmsRequired
        prmType.ParmValid = False
        .Add prmType
    Next
End With
m_Ready = IsReady()
End Sub
Public Function AddParameter(parmNumber As Long, ParmType As SMParameterTypeEnum, ParmValue As String) As Boolean
On Error GoTo ERRHDL
Dim prmType As SMParameterType
m_Ready = IsReady()
If m_MethodName = "" Then
    AddParameter = False
    Exit Function
End If
If parmNumber < 1 Or parmNumber > m_ParmsRequired Then
    AddParameter = False
    Exit Function
End If
With prmType
    .ParmType = ParmType
    .ParmValid = True
    If ParmType = ParmString Then
    .ParmValue = Dquote(ParmValue)
    Else
    .ParmValue = ParmValue
    End If
End With
With m_Parms
    .Add prmType, , , parmNumber
    .Remove parmNumber
End With
m_Ready = IsReady()
AddParameter = True
Exit Function
ERRHDL:
Err.Clear
AddParameter = False
End Function
Private Function IsReady() As Boolean
Dim i As Long
Dim prmType As SMParameterType
    If m_MethodName = "" Then
        IsReady = False
        Exit Function
    End If
    If m_ParmsRequired <> m_Parms.Count Then
        IsReady = False
        Exit Function
    End If
    For i = 1 To m_Parms.Count
        prmType = m_Parms.Item(i)
        If Not prmType.ParmValid Then
            IsReady = False
            Exit Function
        End If
    Next
IsReady = True
End Function
Private Sub Class_Initialize()
Set m_Parms = New Collection
End Sub
Private Sub Class_Terminate()
Set m_Parms = Nothing
End Sub
