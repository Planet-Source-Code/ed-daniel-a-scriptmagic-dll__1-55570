Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SM_BUILTINOBJECTS = " SMDEBUG SMEVENT PARENTFORM "
Public Const SM_RESERVEDWORDS = " Declare And ByRef ByVal Call Case Class Const Dim Do Each Else ElseIf Empty End Eqv Erase Error Exit Explicit False For Function Get If Imp In Is Let Loop Mod Next Not Nothing Null On Option Or Private Property Public Randomize ReDim Resume Select Set Step Sub Then To True Until Wend While Xor As Long Integer Boolean New "
Public Const SM_FUNCTIONCONST = " Anchor Array Asc Atn CBool CByte CCur CDate CDbl Chr CInt CLng Cos CreateObject CSng CStr Date DateAdd DateDiff DatePart DateSerial DateValue Day Dictionary Document Element Err Exp FileSystemObject  Filter Fix Int Form FormatCurrency FormatDateTime FormatNumber FormatPercent GetObject Hex History Hour InputBox InStr InstrRev IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase Left Len Link LoadPicture Location Log LTrim RTrim Trim Mid Minute Month MonthName MsgBox Navigator Now Oct Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second Sgn Sin Space Split Sqr StrComp String StrReverse Tan Time TextStream TimeSerial TimeValue TypeName UBound UCase VarType Weekday WeekDayName Window Year "
Public SM_FUNCTIONWORDS As String
Public Const CALPHA = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Public Const CNUMBERS = "0123456789"
Public SectionStartString As String
Public SectionEndString As String
Public FunctionStartString As String
Public FunctionEndString As String
Public ClassStartString As String
Public ClassEndString As String
Public SubStartString As String
Public SubEndString As String
Public Function MyStringReplace(str2Change As String, str2Insert As String, lStart As Long, lLength As Long) As String
Dim ret As String
If lStart = 1 Then
    ret = str2Insert & Mid$(str2Change, lLength + 1)
Else
    ret = Left$(str2Change, lStart - 1) & str2Insert & Mid$(str2Change, lStart + lLength)
End If
MyStringReplace = ret
End Function

Public Function sm_DecodeText(vText As String)
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)


    Do While CurSpc <= varLen


        DoEvents
            varChr = Mid(vText, CurSpc, 3)


            Select Case varChr
                'lower case
                Case "coe"
                varChr = "a"
                Case "wer"
                varChr = "b"
                Case "ibq"
                varChr = "c"
                Case "am7"
                varChr = "d"
                Case "pm1"
                varChr = "e"
                Case "mop"
                varChr = "f"
                Case "9v4"
                varChr = "g"
                Case "qu6"
                varChr = "h"
                Case "zxc"
                varChr = "i"
                Case "4mp"
                varChr = "j"
                Case "f88"
                varChr = "k"
                Case "qe2"
                varChr = "l"
                Case "vbn"
                varChr = "m"
                Case "qwt"
                varChr = "n"
                Case "pl5"
                varChr = "o"
                Case "13s"
                varChr = "p"
                Case "c%l"
                varChr = "q"
                Case "w$w"
                varChr = "r"
                Case "6a@"
                varChr = "s"
                Case "!2&"
                varChr = "t"
                Case "(=c"
                varChr = "u"
                Case "wvf"
                varChr = "v"
                Case "dp0"
                varChr = "w"
                Case "w$-"
                varChr = "x"
                Case "vn&"
                varChr = "y"
                Case "c*4"
                varChr = "z"
                
                'numbers
                Case "aq@"
                varChr = "1"
                Case "902"
                varChr = "2"
                Case "2.&"
                varChr = "3"
                Case "/w!"
                varChr = "4"
                Case "|pq"
                varChr = "5"
                Case "ml|"
                varChr = "6"
                Case "t'?"
                varChr = "7"
                Case ">^s"
                varChr = "8"
                Case "<s^"
                varChr = "9"
                Case ";&c"
                varChr = "0"
                
                'caps
                Case "$)c"
                varChr = "A"
                Case "-gt"
                varChr = "B"
                Case "|p*"
                varChr = "C"
                Case "1" & Chr(34) & "r"
                varChr = "D"
                Case "c>:"
                varChr = "E"
                Case "@+x"
                varChr = "F"
                Case "v^a"
                varChr = "G"
                Case "]eE"
                varChr = "H"
                Case "aP0"
                varChr = "I"
                Case "{=1"
                varChr = "J"
                Case "cWv"
                varChr = "K"
                Case "cDc"
                varChr = "L"
                Case "*,!"
                varChr = "M"
                Case "fW" & Chr(34)
                varChr = "N"
                Case ".?T"
                varChr = "O"
                Case "%<8"
                varChr = "P"
                Case "@:a"
                varChr = "Q"
                Case "&c$"
                varChr = "R"
                Case "WnY"
                varChr = "S"
                Case "{Sh"
                varChr = "T"
                Case "_%M"
                varChr = "U"
                Case "}'$"
                varChr = "V"
                Case "QlU"
                varChr = "W"
                Case "Im^"
                varChr = "X"
                Case "l|P"
                varChr = "Y"
                Case ".>#"
                varChr = "Z"
                'Special characters
                Case "\" & Chr(34) & "]"
                varChr = "!"
                Case "cY,"
                varChr = "@"
                Case "x%B"
                varChr = "#"
                Case "a*v"
                varChr = "$"
                Case "'&T"
                varChr = "%"
                Case ";%R"
                varChr = "^"
                Case "eG_"
                varChr = "&"
                Case "Z/e"
                varChr = "*"
                Case "rG\"
                varChr = "("
                Case "]*F"
                varChr = ")"
                Case "@B*"
                varChr = "_"
                Case "+Hc"
                varChr = "-"
                Case "&|D"
                varChr = "="
                Case "(:#"
                varChr = "+"
                Case "SlW"
                varChr = "["
                Case "'QB"
                varChr = "]"
                Case "{D>"
                varChr = "{"
                Case "+c%"
                varChr = "}"
                Case "(s:"
                varChr = ":"
                Case "^a("
                varChr = ";"
                Case "16."
                varChr = "'"
                Case "s.*"
                varChr = Chr(34)
                Case "&?W"
                varChr = ","
                Case "GPQ"
                varChr = "."
                Case "SK*"
                varChr = "<"
                Case "RL^"
                varChr = ">"
                Case "40C"
                varChr = "/"
                Case "?#9"
                varChr = "?"
                Case "_?/"
                varChr = "\"
                Case "(_@"
                varChr = "|"
                Case "=#B"
                varChr = " "
            End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3


        DoEvents
        Loop
        sm_DecodeText = varFin
    End Function

Public Function sm_EncodeText(vText As String)
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    
    varLen = Len(vText)


    Do While CurSpc <= varLen


        DoEvents
            CurSpc = CurSpc + 1
            varChr = Mid(vText, CurSpc, 1)

            
            Select Case varChr
                'lower case
                Case "a"
                varChr = "coe"
                Case "b"
                varChr = "wer"
                Case "c"
                varChr = "ibq"
                Case "d"
                varChr = "am7"
                Case "e"
                varChr = "pm1"
                Case "f"
                varChr = "mop"
                Case "g"
                varChr = "9v4"
                Case "h"
                varChr = "qu6"
                Case "i"
                varChr = "zxc"
                Case "j"
                varChr = "4mp"
                Case "k"
                varChr = "f88"
                Case "l"
                varChr = "qe2"
                Case "m"
                varChr = "vbn"
                Case "n"
                varChr = "qwt"
                Case "o"
                varChr = "pl5"
                Case "p"
                varChr = "13s"
                Case "q"
                varChr = "c%l"
                Case "r"
                varChr = "w$w"
                Case "s"
                varChr = "6a@"
                Case "t"
                varChr = "!2&"
                Case "u"
                varChr = "(=c"
                Case "v"
                varChr = "wvf"
                Case "w"
                varChr = "dp0"
                Case "x"
                varChr = "w$-"
                Case "y"
                varChr = "vn&"
                Case "z"
                varChr = "c*4"
                
                'numbers
                Case "1"
                varChr = "aq@"
                Case "2"
                varChr = "902"
                Case "3"
                varChr = "2.&"
                Case "4"
                varChr = "/w!"
                Case "5"
                varChr = "|pq"
                Case "6"
                varChr = "ml|"
                Case "7"
                varChr = "t'?"
                Case "8"
                varChr = ">^s"
                Case "9"
                varChr = "<s^"
                Case "0"
                varChr = ";&c"
                
                'caps
                Case "A"
                varChr = "$)c"
                Case "B"
                varChr = "-gt"
                Case "C"
                varChr = "|p*"
                Case "D"
                varChr = "1" & Chr(34) & "r"
                Case "E"
                varChr = "c>:"
                Case "F"
                varChr = "@+x"
                Case "G"
                varChr = "v^a"
                Case "H"
                varChr = "]eE"
                Case "I"
                varChr = "aP0"
                Case "J"
                varChr = "{=1"
                Case "K"
                varChr = "cWv"
                Case "L"
                varChr = "cDc"
                Case "M"
                varChr = "*,!"
                Case "N"
                varChr = "fW" & Chr(34)
                Case "O"
                varChr = ".?T"
                Case "P"
                varChr = "%<8"
                Case "Q"
                varChr = "@:a"
                Case "R"
                varChr = "&c$"
                Case "S"
                varChr = "WnY"
                Case "T"
                varChr = "{Sh"
                Case "U"
                varChr = "_%M"
                Case "V"
                varChr = "}'$"
                Case "W"
                varChr = "QlU"
                Case "X"
                varChr = "Im^"
                Case "Y"
                varChr = "l|P"
                Case "Z"
                varChr = ".>#"
                'Special characters
                Case "!"
                varChr = "\" & Chr(34) & "]"
                Case "@"
                varChr = "cY,"
                Case "#"
                varChr = "x%B"
                Case "$"
                varChr = "a*v"
                Case "%"
                varChr = "'&T"
                Case "^"
                varChr = ";%R"
                Case "&"
                varChr = "eG_"
                Case "*"
                varChr = "Z/e"
                Case "("
                varChr = "rG\"
                Case ")"
                varChr = "]*F"
                Case "_"
                varChr = "@B*"
                Case "-"
                varChr = "+Hc"
                Case "="
                varChr = "&|D"
                Case "+"
                varChr = "(:#"
                Case "["
                varChr = "SlW"
                Case "]"
                varChr = "'QB"
                Case "{"
                varChr = "{D>"
                Case "}"
                varChr = "+c%"
                Case ":"
                varChr = "(s:"
                Case ";"
                varChr = "^a("
                Case "'"
                varChr = "16."
                Case Chr(34)
                varChr = "s.*"
                Case ","
                varChr = "&?W"
                Case "."
                varChr = "GPQ"
                Case "<"
                varChr = "SK*"
                Case ">"
                varChr = "RL^"
                Case "/"
                varChr = "40C"
                Case "?"
                varChr = "?#9"
                Case "\"
                varChr = "_?/"
                Case "|"
                varChr = "(_@"
                Case " "
                varChr = "=#B"
            End Select
        varFin = varFin & varChr


        DoEvents
        Loop
        
        sm_EncodeText = varFin
    End Function


Public Function CompactVBScript(locQSXML As QSXML) As String
Dim nd As Object
Dim rootNDL As Object
Dim ndc As Object
Dim ndp As Object
Dim ndl As Object
Dim locLn As String
Dim buff$, i As Long
Dim bNoWhiteSpace As Boolean
bNoWhiteSpace = True
ReDim ed1(0) As String

With locQSXML
    Set nd = .GetRootElement()
    If .GetAttributeValue(nd, "NAME") = "" Then
        CompactVBScript = ""
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
            buff$ = buff$ & GetSectionHeader("Script Input Variables") & vbLf & vbLf
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
                If Right$(ed1(i), 1) <> "_" Then
                buff$ = buff$ & ed1(i) & locLn & "LN:" & (i + 1) & Chr$(187) & vbLf
                Else
                buff$ = buff$ & ed1(i) & vbLf
                End If
            End If
        End If
    Next
End If
DoEvents
CompactVBScript = buff$
End Function
Public Function GetSectionHeader(sectName As String) As String
Dim ret As String
ret = Replace(SectionStartString, "<SECTIONNAME>", sectName)
GetSectionHeader = ret
End Function
Public Function GetSectionFooter(sectName As String) As String
Dim ret As String
ret = Replace(SectionEndString, "<SECTIONNAME>", sectName)
GetSectionFooter = ret
End Function
Public Function GetFunctionHeader(funcName As String) As String
Dim ret As String
ret = Replace(FunctionStartString, "<FUNCTIONNAME>", funcName)
GetFunctionHeader = ret
End Function
Public Function GetClassHeader(funcName As String) As String
Dim ret As String
ret = Replace(ClassStartString, "<FUNCTIONNAME>", funcName)
GetClassHeader = ret
End Function
Public Function GetSubHeader(funcName As String) As String
Dim ret As String
ret = Replace(SubStartString, "<FUNCTIONNAME>", funcName)
GetSubHeader = ret
End Function
Public Function GetFunctionFooter(funcName As String) As String
Dim ret As String
ret = Replace(FunctionEndString, "<FUNCTIONNAME>", funcName)
GetFunctionFooter = ret
End Function
Public Function GetClassFooter(funcName As String) As String
Dim ret As String
ret = Replace(ClassEndString, "<FUNCTIONNAME>", funcName)
GetClassFooter = ret
End Function
Public Function GetSubFooter(funcName As String) As String
Dim ret As String
ret = Replace(SubEndString, "<FUNCTIONNAME>", funcName)
GetSubFooter = ret
End Function

Public Function Dquote(strVal As String) As String
Dquote = Chr$(34) & strVal & Chr$(34)
End Function
Public Function GetPathFromFileName(strFilename As String) As String
Dim i As Long
Dim ret As String
i = InStrRev(strFilename, "\")
If i > 0 Then
    ret = Left$(strFilename, i - 1)
Else
    ret = ""
End If
GetPathFromFileName = ""
End Function
Public Function isValidObjName(strName As String, Optional bVerbose As Boolean = True) As Boolean
Dim i As Long, buff$
If strName = "" Then
    If bVerbose Then
        MsgBox "Object names cannot be blank", vbCritical, "Error.."
    End If
    isValidObjName = False
    Exit Function
End If
If InStr(strName, " ") > 0 Then
    If bVerbose Then
        MsgBox "Object names cannot contain spaces", vbCritical, "Error.."
    End If
    isValidObjName = False
    Exit Function
End If
If InStr(CALPHA, UCase$(Left$(strName, 1))) = 0 Then
    If bVerbose Then
        MsgBox "Object names must begin with A-Z or a-z", vbCritical, "Error.."
    End If
    isValidObjName = False
    Exit Function
End If
buff$ = AlphaNumFormat(strName, "_")
If buff$ <> strName Then
    If bVerbose Then
        MsgBox "Object names may not contain punctuation or special characters", vbCritical, "Error.."
    End If
    isValidObjName = False
    Exit Function
End If
isValidObjName = True
End Function

Public Function AlphaNumFormat(strValue As String, Optional cExtra As String = "")
Dim i As Long
Dim j As Long
Dim buff$, c$
Dim ret As String
buff$ = UCase$(CALPHA & CNUMBERS & cExtra)
For i = 1 To Len(strValue)
    c$ = Mid$(strValue, i, 1)
    If InStr(buff$, UCase$(c$)) > 0 Then
        ret = ret & c$
    End If
Next
AlphaNumFormat = ret
End Function
Public Function NumFormat(strValue As String, Optional cExtra As String = "")
Dim i As Long
Dim j As Long
Dim buff$, c$
Dim ret As String
buff$ = UCase$(CNUMBERS & cExtra)
For i = 1 To Len(strValue)
    c$ = Mid$(strValue, i, 1)
    If InStr(buff$, UCase$(c$)) > 0 Then
        ret = ret & c$
    End If
Next
NumFormat = ret
End Function

Public Sub InitGlobals()
If SectionStartString = "" Then
SM_FUNCTIONWORDS = SM_FUNCTIONCONST
SectionStartString = "'" & Chr$(171) & " <SECTIONNAME> " & Chr$(187)
SectionEndString = "'" & Chr$(171) & " END <SECTIONNAME> " & Chr$(187)
FunctionStartString = "'" & Chr$(171) & "FUNCTION: <FUNCTIONNAME> " & Chr$(187)
FunctionEndString = "'" & Chr$(171) & "END FUNCTION: <FUNCTIONNAME> " & Chr$(187)
ClassStartString = "'" & Chr$(171) & "CLASS: <FUNCTIONNAME> " & Chr$(187)
ClassEndString = "'" & Chr$(171) & "END CLASS: <FUNCTIONNAME> " & Chr$(187)
SubStartString = "'" & Chr$(171) & "SUBROUTINE: <FUNCTIONNAME> " & Chr$(187)
SubEndString = "'" & Chr$(171) & "END SUBROUTINE: <FUNCTIONNAME> " & Chr$(187)
End If
End Sub

