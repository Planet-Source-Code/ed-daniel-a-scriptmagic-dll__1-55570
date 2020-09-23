VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox EventBoxes 
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "Event Boxes: This one is for READYSTATECHANGE"
      Top             =   3420
      Width           =   4035
   End
   Begin VB.TextBox EventBoxes 
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Event Boxes: This one is for SAVEREQUEST"
      Top             =   2940
      Width           =   4035
   End
   Begin VB.TextBox EventBoxes 
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Event Boxes: This one is for EVENT"
      Top             =   2460
      Width           =   4035
   End
   Begin VB.TextBox EventBoxes 
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Event Boxes: This one is for ERROR"
      Top             =   1980
      Width           =   4035
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Load A Template Test"
      Height          =   375
      Left            =   60
      TabIndex        =   12
      Top             =   2220
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Show Editor"
      Height          =   375
      Left            =   2340
      TabIndex        =   11
      Top             =   1740
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "RunSub Test"
      Height          =   375
      Left            =   60
      TabIndex        =   10
      Top             =   1740
      Width           =   2175
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Show an Object"
      Height          =   375
      Left            =   2340
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Show a Variable Name"
      Height          =   375
      Left            =   60
      TabIndex        =   8
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Execute MyFunction1"
      Height          =   375
      Left            =   2340
      TabIndex        =   7
      Top             =   900
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop Script"
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   900
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Execute Script"
      Height          =   375
      Left            =   2340
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Script"
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "XML String"
      Height          =   375
      Left            =   2340
      TabIndex        =   3
      Top             =   60
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2715
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   4320
      Width           =   9075
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5220
      Pattern         =   "*.mgk"
      TabIndex        =   1
      Top             =   240
      Width           =   3675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Version"
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click to load a project"
      Height          =   195
      Left            =   5220
      TabIndex        =   17
      Top             =   60
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************These are the relevant enums defined in script magic DLL********
'Public Enum SMReadyStateType
'    Waiting = -1
'    Initialized = 0
'    Ready = 1
'    Executing = 2
'End Enum
'Public Enum SMProjectOpenMode
'    OpenFile = 0
'    OpenString = 1
'End Enum
'Public Type SMLastErrorType
'    ErrorNumber As Long
'    ScriptLineNumber As Long
'    ErrorDescription As String
'    ErrorSource As String
'    OffendingCode As String
'End Type

Private oScriptMagic As Object
Attribute oScriptMagic.VB_VarHelpID = -1
'Private WithEvents oScriptMagic As ScriptMagicDLL.ScriptEngine

Private Sub Command1_Click()
With oScriptMagic
    MsgBox .Version
End With
End Sub

Private Sub Command10_Click()

With oScriptMagic
If .ReadyState = -1 Then
    MsgBox "No Script is running"
    Exit Sub
End If
.ShowEditor
End With
End Sub

Private Sub Command11_Click()
Dim i As Long
Dim buff$
Dim mt
'Public Enum SMTemplateAddErrorEnum
'    eNoError = 0
'    eObjectExists = 1
'    eInvalidParms = 2
'    eInvalidCode = 3
'    eInvalidProjectCode = 4
'End Enum

'Dim xTmp As New SM_Templates
Dim xTmp As Object
Set xTmp = CreateObject("ScriptMagicDLL.SM_Templates")
xTmp.GenerateShellProject

mt = xTmp.BlankMethod()
For i = 1 To 50
    mt.MethodName = "MyFunction_" & i
    mt.SubFunction = 1
    mt.ParmString = "a,b,c"
    mt.Code = "MsgBox a+b+c"
    If xTmp.AddMethod(mt) <> 0 Then
        MsgBox "Error Occurred "
        Set xTmp = Nothing
        Exit Sub
    End If
Next
For i = 1 To 50
    mt.MethodName = "MySubRoutine_" & i
    mt.SubFunction = 0
    mt.ParmString = ""
    mt.Code = ""
    'See Enum above
    If xTmp.AddMethod(mt) <> 0 Then
        MsgBox "Error Occurred "
        Set xTmp = Nothing
        Exit Sub
    End If
Next
'
MsgBox "Template Generated...."
buff$ = xTmp.SM_TemplateProject
'Text1.Text = buff$
'Public Enum SMProjectOpenMode
'    OpenFile = 0
'    OpenString = 1
'End Enum
'oScriptMagic.OpenProject buff$, OpenString
oScriptMagic.OpenProject buff$, 1
oScriptMagic.ExecuteScript
Set xTmp = Nothing
End Sub

Private Sub Command2_Click()
With oScriptMagic
    .verbose = True
    Text1.Text = .XMLString
    
End With
End Sub

Private Sub Command3_Click()
Dim buff$
With oScriptMagic
    .verbose = True
    buff$ = .SM_VBScript
    If InStr(buff$, vbCrLf) = 0 Then
        buff$ = Replace(buff$, vbLf, vbCrLf)
    End If
    Text1.Text = buff$
End With

End Sub

Private Sub Command4_Click()
With oScriptMagic
    If .ReadyState = -1 Then
        MsgBox "Load a project and try again."
        Exit Sub
    End If
    .ExecuteScript
    
End With
End Sub

Private Sub Command5_Click()
If oScriptMagic.ReadyState <> -1 Then
    oScriptMagic.Halt
End If

Exit Sub
Dim xSub As Object
With oScriptMagic
    .RunSub "MySubroutine2", "Ed Daniel", "Dynamic"
    Exit Sub
End With
With oScriptMagic
'Public Enum SMReadyStateType
'    Waiting = -1
'    Initialized = 0
'    Ready = 1
'    Executing = 2
'End Enum
    If .ReadyState <> 2 Then
        MsgBox "Script is not executing: ReadyState = " & .ReadyState
        Exit Sub
    End If
    Set xSub = .GetMethod("MySubRoutine2")
    If xSub Is Nothing Then Exit Sub
    If .LastError.ErrorNumber <> 0 Then Exit Sub
End With
'Public Enum SMParameterTypeEnum
'    ParmString = 0
'    parmNumber = 1
'    ParmScriptObj = 2
'End Enum

With xSub
    .AddParameter 1, 0, "Ed Daniel"
    .AddParameter 2, 0, "473-5902"
    If .Ready Then
        MsgBox .MethodName & " is ready to execute with " & .ParmsRequired & " parameters"
    Else
        MsgBox .MethodName & " is not ready"
    End If
    If .Ready Then
        oScriptMagic.ExecuteScriptMethod xSub
    End If
End With

End Sub

Private Sub Command6_Click()
Dim xSub As Object
With oScriptMagic
    If .ReadyState <> 2 Then
        MsgBox "Script is not executing: ReadyState = " & .ReadyState
        Exit Sub
    End If
    Set xSub = .GetMethod("MyFunction1")
    If xSub Is Nothing Then
        MsgBox "Error no MyFunction1"
        Exit Sub
    End If
End With
With xSub
    .MethodReturnsValue = True
'Public Enum SMParameterTypeEnum
'    ParmString = 0
'    parmNumber = 1
'    ParmScriptObj = 2
'End Enum
    
    .AddParameter 1, 1, 50
    .AddParameter 2, 1, 10
    .AddParameter 3, 1, 5
    If .Ready Then
        MsgBox .MethodName & " is ready to execute with " & .ParmsRequired & " parameters"
    Else
        MsgBox .MethodName & " is not ready"
    End If
    If .Ready Then
        oScriptMagic.ExecuteScriptMethod xSub
    End If
    MsgBox .MethodName & " returned: " & CStr(.ReturnValue & "")
End With

Set xSub = Nothing

End Sub

Private Sub Command7_Click()
Dim buff$
Static vName As String
Dim ret As Variant
With oScriptMagic
If .ReadyState <> 2 Then
    MsgBox "No Script is running"
    Exit Sub
End If
buff$ = InputBox("Enter Variable Name below", "GetValue", vName)

If buff$ = "" Then Exit Sub
vName = buff$
    ret = .GetScriptValue(buff$)
    If .LastError.ErrorNumber <> 0 Then
        MsgBox .LastError.ErrorDescription
        .ClearErrors
    Else
        MsgBox buff$ & " = " & CStr(ret & "")
    End If
End With
End Sub

Private Sub Command8_Click()
Dim buff$
Static vName As String
Dim ret As Object
With oScriptMagic
If .ReadyState <> 2 Then
    MsgBox "No Script is running"
    Exit Sub
End If
If vName = "" Then
    vName = "MTObject"
End If
buff$ = InputBox("Enter Object Variable Name below", "GetValue", vName)
If buff$ = "" Then Exit Sub
vName = buff$
    Set ret = .GetScriptObject(buff$)
    If .LastError.ErrorNumber <> 0 Then
        MsgBox .LastError.ErrorDescription
        .ClearErrors
    Else
        If (ret Is Nothing) Then
            MsgBox "Ret is nothing"
        Else
            MsgBox "Hello the script did return a valid object named: " & buff$
        End If
    End If
End With

End Sub

Private Sub Command9_Click()
Dim buff$
Static vName As String
Dim ret As Object
With oScriptMagic
If .ReadyState <> 2 Then
    MsgBox "No Script is running"
    Exit Sub
End If
If .RunSub("MySub", "1 parameter") Then
    MsgBox "Success"
Else
    MsgBox "Failure"
End If
End With
End Sub

Private Sub EventBoxes_Change(Index As Integer)
'If you declare script magic in your project as a WithEvents object you get
'several different events.  However, I have a large project of which
'script magic is only a small piece, and I didn't want to have to recompile
'and redeploy the entire application because I made one change to the
'script magic dll.  This was my simple solution.  For each event you can
'bind a text box control.  Then put your event handle in the Change event
'of that text box.  The text box can be invisible at run time.  Now I can
'use LATE binding in my project, and just redeploy new ScriptMagic DLLs
Dim deLimiter As String
ReDim evt(0) As String
deLimiter = Chr$(176)
'When using the actual events structures are used.  When using textboxes
'if a structure is supposed to be passed, a string array is passed instead
'using Chr(176) as the delimiter
With EventBoxes(Index)
    If .Text = "" Then
        'Do nothing the event box is being cleared
        Exit Sub
    End If
    Select Case Index
        Case 0
        'In my example I'm using eventboxes(0) as an error event trapper.
        'The error event returns an SMLastErrorType structure.
            evt = Split(.Text, deLimiter)
            MsgBox "Error# " & evt(0) & " occurred at line: " & evt(1) & " Description: " & evt(2)
        Case 1
        'Text 1 is the event handler.  This means a script used the SMEvent.Raise method
        'to send an event back to the host.
            evt = Split(.Text, deLimiter)
            'Element 0 is the event name (defined in the script)
            'Element 1 contains an optional string of extra data
            MsgBox "Script Created Event: " & evt(0) & vbCrLf & "Parms: " & evt(1)
        Case 2
          'Textbox 2 is for the save request.  The text is set to the entire scriptmagic project xml string
            oScriptMagic_SaveRequest .Text
        Case 3
          'Textbox 3 is for a scriptobject readystate change, do whatever you like
    End Select
   ' .Text = ""  'Clear the box for the next event
End With


End Sub

Private Sub File1_DblClick()
Dim buff$
With File1
    If .ListCount = 0 Then Exit Sub
    buff$ = App.Path & "\" & .List(.ListIndex)
End With
With oScriptMagic
    If Not .OpenProject(buff$, 0) Then
        MsgBox "Failed To Open"
    Else
        MsgBox "Project file opened"
    End If
End With
End Sub

Private Sub Form_Load()
Set oScriptMagic = CreateObject("ScriptMagicDLL.ScriptEngine")
With oScriptMagic
    .ParentForm = Me
    .verbose = True
    .EchoTextBox = Text1
    .ErrorTextBox = EventBoxes(0)
    .EventTextBox = EventBoxes(1)
    .SaveRequestTextBox = EventBoxes(2)
    .ReadyStateTextBox = EventBoxes(3)
    '.BindImmediateResultsBox Me.Text1
End With
File1.Path = App.Path
File1.Pattern = "*.mgk"
File1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oScriptMagic = Nothing
End Sub

Private Sub oScriptMagic_SaveRequest(strScriptMagic As String)
MsgBox "Save requested by Debug Editor"
Dim buff$
With File1
    If .ListCount = 0 Then Exit Sub
    buff$ = App.Path & "\" & .List(.ListIndex)
End With
With oScriptMagic
    If Not .ExportScriptToFile(buff$) Then
        MsgBox "Failed To Save"
    Else
        MsgBox "Project file Saved as: " & buff$
    End If
End With
End Sub

Private Sub oScriptMagic_ScriptEvent(evtName As String, evtParms As String)
MsgBox "Event: " & evtName & " just occurred with a parmstring of '" & evtParms & "'"
End Sub

Private Sub oScriptMagic_SMError(smErr)
With smErr
    MsgBox "Error: " & CStr(.ErrorNumber) & " - " & .ErrorDescription
    If smErr.ScriptLineNumber > 1 Then
        MsgBox "Error was at line: " & .ScriptLineNumber & vbLf & "Code: " & .OffendingCode
    End If
End With
End Sub
