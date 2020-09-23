VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Project Properties"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      Caption         =   "Option Explicit (All variables must be declared)"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   2220
      Width           =   5115
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      ItemData        =   "frmProperties.frx":030A
      Left            =   3000
      List            =   "frmProperties.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3420
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   1
      Left            =   6900
      TabIndex        =   9
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Save"
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2700
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5400
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Password"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   2220
      Width           =   2955
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   60
      MaxLength       =   1024
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   1020
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   2700
      Width           =   4035
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "frmProperties.frx":0330
      Left            =   120
      List            =   "frmProperties.frx":033A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3420
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   60
      MaxLength       =   50
      TabIndex        =   0
      Top             =   420
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "minutes"
      Height          =   255
      Index           =   5
      Left            =   5700
      TabIndex        =   19
      Top             =   3420
      Width           =   2595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Script Timeout (Seconds)"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   18
      Top             =   3180
      Width           =   2595
   End
   Begin VB.Label Label2 
      Caption         =   "Created:"
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   17
      Top             =   480
      Width           =   5355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Desc"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   3840
      UseMnemonic     =   0   'False
      Width           =   8835
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Execute Method"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   3180
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Created:"
      Height          =   255
      Index           =   0
      Left            =   3660
      TabIndex        =   14
      Top             =   120
      Width           =   5355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Description"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   780
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   2460
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   11
      Top             =   180
      Width           =   2595
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MyPrjXML As String
Private MyXMLOBJ As QSXML

Private Sub Check1_Click()
    Text3.Enabled = CBool(Check1.Value)
    Command1.Enabled = Text3.Enabled
End Sub

Private Sub Combo1_Click()
Dim buff$
Select Case Combo1.ListIndex
    Case 0
     buff$ = "IMMEDIATE Script Magic scripts are immediately when run in the user interface.  So unless you have executable " & _
     "code in the 'Initialization' section to run you will see nothing happen.  When you select run from the toolbar the " & _
     "Microsoft VB Scripting is invoked to execute all of your script as a whole."
    Case 1
     buff$ = "INTERACTIVE Script Magic scripts run in an interactive debugger. Code in the " & _
     "'Initialization' section is immediately executed when you select run from the toolbar.  " & _
     "All of the code is loaded into memory and you can selectively execute " & _
     " PUBLIC SubRoutines or Functions in the Interactive Script Testing Window."
    
End Select
Label3.Caption = buff$
End Sub

Private Sub Combo2_Click()
Dim x As Currency
Label1(5).Caption = ""
With Combo2
    If .ListIndex < 1 Then Exit Sub
    x = (CLng(.Text) / 60)
    Label1(5).Caption = " = " & Mid$(Format$(x, "Currency"), 2) & " minutes"
End With
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
    Case 0
        If SaveFormData() Then
            Unload Me
            Exit Sub
        End If
    Case 1
    Unload Me
End Select
End Sub
Private Function SaveFormData() As Boolean
Dim nd As Object
Dim ndc As Object
If CBool(Check1.Value) Then
    Text3.Text = Trim$(Text3.Text)
    If Text3.Text = "" Then
        MsgBox "Password is checked but no password is entered.", vbCritical, "Error.."
        Text3.SetFocus
        SaveFormData = False
        Exit Function
    End If
Else
    If Trim$(Text3.Text) <> "" Then
        If MsgBox("Save project without password?", vbYesNo + vbQuestion, "Remove Password") = vbNo Then
        SaveFormData = False
        Exit Function
        End If
    End If
End If
For i = 0 To Text1.UBound
Text1(i).Text = Trim$(Text1(i).Text)
Next
If Text1(0).Text = "" Then
    MsgBox "Project name is required", vbInformation, "Error.."
    Text1(0).SetFocus
    SaveFormData = False
    Exit Function
End If
With MyXMLOBJ
    Set nd = .GetRootElement
    Set ndc = .GetChildNode(nd.childNodes, "DESCRIPTION")
    .SetAttribute nd, "NAME", Text1(0).Text
    .SetAttribute nd, "AUTHOR", Text1(1).Text
    .SetAttribute nd, "RUNMODE", Combo1.Text
    .SetAttribute nd, "TIMEOUT", CStr(Combo2.ListIndex)
    .SetAttribute nd, "EXPLICIT", CStr(Check2.Value)
    If CBool(Check1.Value) Then
    .SetAttribute nd, "PASSWORD", sm_EncodeText(Text3.Text)
    Else
    .SetAttribute nd, "PASSWORD", ""
    End If
    ndc.Text = MySingleQuote(Text2.Text)
    SaveFormData = frmCodeMain.UpdateProject(.XML)
End With
End Function
Private Function MySingleQuote(strText) As String
Dim i As Long
Dim buff$
Dim ret As String
ReDim ed1(0) As String

buff$ = Replace(strText, vbCrLf, vbLf)
ed1 = Split(buff$, vbLf)
For i = 0 To UBound(ed1)
    ed1(i) = Trim$(ed1(i))
    If ed1(i) <> "" Then
        If Left$(ed1(i), 1) <> "'" Then
            ed1(i) = "'" & ed1(i)
        End If
        ret = ret & ed1(i) & vbLf
    End If
Next
MySingleQuote = ret
End Function
Private Sub Form_Load()
Dim i As Long
Combo2.Clear
Combo2.AddItem "<Unlimited>"
For i = 1 To 1200
    Combo2.AddItem CStr(i)
Next
Set MyXMLOBJ = New QSXML
MyXMLOBJ.Initialize pavAUTO
MyXMLOBJ.OpenFromString MyPrjXML
LoadFormData
End Sub
Private Sub LoadFormData()
Dim nd As Object
Dim ndc As Object
Dim buff$
With MyXMLOBJ
    Set nd = .GetRootElement
    Check2.Value = CLng("0" & .GetAttributeValue(nd, "EXPLICIT"))
    Text1(0).Text = .GetAttributeValue(nd, "NAME")
    Text1(1).Text = .GetAttributeValue(nd, "AUTHOR")
    Label2(0).Caption = "Created: " & .GetAttributeValue(nd, "CREATED")
    Label2(1).Caption = "Last Modified: " & .GetAttributeValue(nd, "LASTMODIFIED")
    buff$ = .GetAttributeValue(nd, "PASSWORD")
    If buff$ = "" Then
        Check1.Value = vbUnchecked
        Text3.Enabled = False
        Command1.Enabled = False
    Else
        Check1.Value = vbChecked
        Text3.Text = sm_DecodeText(buff$)
        Command1.Enabled = True
    End If
    If .IsChildNode(nd, "DESCRIPTION") Then
        Set ndc = .GetChildNode(nd.childNodes, "DESCRIPTION")
        If InStr(ndc.Text, vbLf) > 0 And InStr(ndc.Text, vbCrLf) = 0 Then
        buff$ = Replace(ndc.Text, vbLf, vbCrLf)
        Else
        buff$ = ndc.Text
        End If
        Text2.Text = buff$
    Else
        buff$ = "<DESCRIPTION></DESCRIPTION>"
        Set ndc = .XMLAddNode(nd, buff$)
        ndc.Text = "'Script Magic Project: " & Text1(0).Text
        Text2.Text = ndc.Text
    End If
    buff$ = .GetAttributeValue(nd, "TIMEOUT")
    If buff$ = "" Then buff$ = "10"
    Combo2.ListIndex = CLng(buff$)
    
    buff$ = .GetAttributeValue(nd, "RUNMODE")
    If buff$ = "" Then
        Combo1.ListIndex = 0
    Else
        If buff$ = "INTERACTIVE" Then
            Combo1.ListIndex = 1
        Else
            Combo1.ListIndex = 0
        End If
        
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set MyXMLOBJ = Nothing
End Sub

