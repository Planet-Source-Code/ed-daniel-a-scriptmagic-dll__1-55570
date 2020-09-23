VERSION 5.00
Begin VB.Form EditSubDefinition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New SubRoutine"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "EditSubDefinition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      Caption         =   "Function"
      Height          =   255
      Index           =   1
      Left            =   2940
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   360
      Width           =   1755
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SubRoutine"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Value           =   -1  'True
      Width           =   1755
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "EditSubDefinition.frx":0442
      Left            =   1080
      List            =   "EditSubDefinition.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      MaxLength       =   32
      TabIndex        =   1
      Text            =   "MySubRoutine"
      Top             =   1740
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SubRoutine Scope"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   840
      Width           =   3795
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   5520
      Picture         =   "EditSubDefinition.frx":0461
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   0
      Picture         =   "EditSubDefinition.frx":08A3
      Stretch         =   -1  'True
      Top             =   60
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Prototype"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   2820
      Width           =   3195
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      UseMnemonic     =   0   'False
      Width           =   5835
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   6
      Top             =   2220
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SubRoutine Name"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   5
      Top             =   1500
      Width           =   3855
   End
End
Attribute VB_Name = "EditSubDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IsLoaded As Boolean
Public InitClassType As Integer


Private Sub Combo1_Click()
SetLabel2
End Sub

Private Sub Command1_Click(Index As Integer)
Dim buff$
Dim clsName As String
clsName = IIf(Option1(0).Value = True, "SUBROUTINES", "FUNCTIONS")
Select Case Index
    Case 0
        If ValidateData() Then
            If frmCodeMain.ItemExists(Text1(0).Text) Then
                MsgBox "An object called '" & Text1(0).Text & "' already exists in your project.  Choose another name.", vbCritical, "Error.."
                Text1(0).SetFocus
                Exit Sub
            End If
            If Not ValidateParameterList(Text1(1).Text) Then
                Text1(1).SetFocus
                Exit Sub
            End If
            Text1(1).Text = FormatParameterList(Text1(1).Text)
            buff$ = MakeSubXML()
            If frmCodeMain.AddProjectItem(clsName, buff$) Then
                Unload Me
                Exit Sub
            End If
        End If
    Case 1
        Unload Me
End Select
End Sub
Private Function MakeSubXML() As String
Dim ob1 As New QSXML
Dim nd As Object
Dim strXML As String
Dim buff$
With ob1
    .Initialize pavAUTO
    If Option1(0).Value Then
    .CreateRootElement "", "SUBROUTINE"
    Else
    .CreateRootElement "", "FUNCTION"
    End If
    Set nd = .GetRootElement()
    .SetAttribute nd, "NAME", Text1(0).Text
    .SetAttribute nd, "PARAMETERS", Text1(1).Text
    .SetAttribute nd, "SCOPE", Combo1.Text
    If Option1(0).Value Then
        buff$ = "'" & vbLf
        buff$ = buff$ & "MsgBox " & Dquote("TO DO: Add processing code for " & Text1(0).Text) & vbLf
        buff$ = buff$ & "'" & vbLf
    Else
        buff$ = "'" & vbLf
        buff$ = buff$ & "Dim retValue " & vbLf & Text1(0).Text & " = retValue " & vbLf
        buff$ = buff$ & "MsgBox " & Dquote("TO DO: Add processing code for " & Text1(0).Text) & vbLf
        buff$ = buff$ & "'" & vbLf
    End If
    nd.Text = buff$
    strXML = .XML
End With
MakeSubXML = strXML
Set ob1 = Nothing
End Function
Private Function ValidateData() As Boolean
Dim buff$
Text1(0) = Trim$(Text1(0))
Text1(1) = Trim$(Text1(1))
buff$ = Text1(0)
If buff$ = "" Then
    MsgBox "Enter an object name", vbCritical, "Error.."
    Text1(0).SetFocus
    ValidateData = False
    Exit Function
End If
If InStr(buff$, " ") > 0 Then
    MsgBox "Names may not contain spaces.", vbCritical, "Error.."
    ValidateData = False
    Exit Function
End If
If InStr(CALPHA, UCase$(Left$(buff$, 1))) = 0 Then
    MsgBox "Object names must begin with A-Z or a-z", vbCritical, "Error.."
    ValidateData = False
    Exit Function
End If
ValidateData = True
End Function

Private Sub Form_Activate()
If Not IsLoaded Then
    Text1(0).SetFocus
    IsLoaded = True
End If
End Sub

Private Sub Form_Load()
Dim i As Long
Option1(InitClassType).Value = True
If Option1(0).Value Then
    i = frmCodeMain.CountChildren("SUBROUTINES")
    Text1(0).Text = "MySubRoutine" & i + 1
Else
    i = frmCodeMain.CountChildren("FUNCTIONS")
    Text1(0).Text = "MyFunction" & i + 1
End If
Combo1.ListIndex = 0
SetLabel2
Text1(0).SelLength = Len(Text1(0).Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
InitClassType = 0
IsLoaded = False
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value Then
    Label1(0).Caption = "SubRoutine Scope"
    Label1(1).Caption = "SubRoutine Name"
    Me.Caption = "New SubRoutine"
Else
    Label1(0).Caption = "Function Scope"
    Label1(1).Caption = "Function Name"
    Me.Caption = "New Function"
End If
SetLabel2
End Sub

Private Sub Text1_Change(Index As Integer)
SetLabel2
End Sub
Private Sub SetLabel2()
Dim buff$
If Option1(0).Value Then
buff$ = Combo1.Text & " Sub " & Trim$(Text1(0).Text) & "(" & _
Trim$(Text1(1).Text) & ")" & vbLf & vbLf & "End Sub"
Label2.Caption = buff$
Else
buff$ = Combo1.Text & " Function " & Trim$(Text1(0).Text) & "(" & _
Trim$(Text1(1).Text) & ")" & vbLf
buff$ = buff$ & "Dim retValue" & vbLf & _
Text1(0).Text & " = retValue"
buff$ = buff$ & vbLf & "End Sub"
Label2.Caption = buff$
End If
End Sub
