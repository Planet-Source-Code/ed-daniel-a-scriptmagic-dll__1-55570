VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_smdllobj As ScriptMagicDLL.ScriptEngine
Private m_rebindObj As Object
Private m_isLoaded As Boolean
Friend Property Get IsLoaded() As Boolean
    IsLoaded = m_isLoaded
End Property
Public Sub Show(smDLLObj As ScriptMagicDLL.ScriptEngine)
    Set m_smdllobj = smDLLObj
    Set m_rebindObj = smDLLObj.GetDebuggerControl()
    frmCodeMain.SetScriptDLLObject smDLLObj
    smDLLObj.EchoTextBox = frmCodeMain.EXEResult
    frmCodeMain.Show 1
    m_isLoaded = True
End Sub
Public Sub CloseEditor()

    m_isLoaded = False
End Sub
Private Sub Class_Initialize()
    InitGlobals
End Sub
Private Sub InitGlobals()
SM_FUNCTIONWORDS = SM_FUNCTIONCONST
SectionStartString = "'" & Chr$(171) & " <SECTIONNAME> " & Chr$(187)
SectionEndString = "'" & Chr$(171) & " END <SECTIONNAME> " & Chr$(187)
FunctionStartString = "'" & Chr$(171) & "FUNCTION: <FUNCTIONNAME> " & Chr$(187)
FunctionEndString = "'" & Chr$(171) & "END FUNCTION: <FUNCTIONNAME> " & Chr$(187)
ClassStartString = "'" & Chr$(171) & "CLASS: <FUNCTIONNAME> " & Chr$(187)
ClassEndString = "'" & Chr$(171) & "END CLASS: <FUNCTIONNAME> " & Chr$(187)
SubStartString = "'" & Chr$(171) & "SUBROUTINE: <FUNCTIONNAME> " & Chr$(187)
SubEndString = "'" & Chr$(171) & "END SUBROUTINE: <FUNCTIONNAME> " & Chr$(187)
End Sub

Private Sub Class_Terminate()
m_smdllobj.EchoTextBox = m_rebindObj
End Sub
