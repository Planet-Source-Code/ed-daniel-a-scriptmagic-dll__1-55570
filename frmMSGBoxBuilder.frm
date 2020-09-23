VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmMSGBoxBuilder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBScript Message Box Builder"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8040
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMSGBoxBuilder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List3 
      Height          =   1440
      IntegralHeight  =   0   'False
      ItemData        =   "frmMSGBoxBuilder.frx":01CA
      Left            =   3000
      List            =   "frmMSGBoxBuilder.frx":01DD
      TabIndex        =   3
      Top             =   4200
      Width           =   2355
   End
   Begin VB.ListBox List2 
      Height          =   1440
      IntegralHeight  =   0   'False
      ItemData        =   "frmMSGBoxBuilder.frx":0217
      Left            =   5400
      List            =   "frmMSGBoxBuilder.frx":0227
      TabIndex        =   4
      Top             =   4200
      Width           =   2475
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   2
      Left            =   5820
      TabIndex        =   8
      Top             =   6480
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Test Message Box"
      Height          =   435
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Generate Code"
      Height          =   435
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   6480
      Width           =   1515
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Generate stub code to process the response"
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   5880
      Value           =   1  'Checked
      Width           =   6375
   End
   Begin MSScriptControlCtl.ScriptControl SC1 
      Left            =   4800
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ListBox List1 
      Height          =   1440
      IntegralHeight  =   0   'False
      ItemData        =   "frmMSGBoxBuilder.frx":0253
      Left            =   540
      List            =   "frmMSGBoxBuilder.frx":0269
      TabIndex        =   2
      Top             =   4200
      Width           =   2355
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Index           =   1
      Left            =   480
      MaxLength       =   2048
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1860
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   480
      MaxLength       =   50
      TabIndex        =   0
      Text            =   "Script Magic Message Box"
      Top             =   1140
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Picture to Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   14
      Top             =   3960
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Default button"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   5460
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Buttons to Display (Check all that apply)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   7035
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Buttons to Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   540
      TabIndex        =   11
      Top             =   3960
      Width           =   2355
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   1
      Left            =   7560
      Picture         =   "frmMSGBoxBuilder.frx":02C8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   0
      Left            =   0
      Picture         =   "frmMSGBoxBuilder.frx":070A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Message to Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Message Box Title"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "frmMSGBoxBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        'Ok
        Clipboard.Clear
        Clipboard.SetText CodeGen, vbCFText
        Unload Me
    Case 1
        'Test
        TestMsgBox
    Case 2
        'Exit
        Clipboard.Clear
        Unload Me
End Select
End Sub
Private Sub TestMsgBox()
Dim buff$
buff$ = CodeGen
On Error GoTo ERRHDL
With Me.SC1
    If MsgBox("Generated Script Below: (Click Ok to Execute)" & vbLf & vbLf & buff$, vbOKCancel, "VBScript Code Preview") = vbCancel Then
        Exit Sub
    End If
    .Reset
    .ExecuteStatement buff$
    .Reset
End With
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear
SC1.Reset
Exit Sub
End Sub
Private Function locFormatMessage(strMsg As String) As String
Dim ret As String
Dim i As Long
ret = Replace(strMsg, vbCrLf, vbLf)
ret = Replace(ret, vbLf, Dquote(" & vblf & "))
locFormatMessage = Dquote(ret)

End Function
Private Function CodeGen() As String
Dim ret As String
Dim strTitle As String
Dim strMsg As String
Dim strButtons As String
Dim buff$, buff2$
Dim stb As Boolean
Dim i As VbMsgBoxResult
Dim j As VbMsgBoxStyle
stb = CBool(Check1.Value)
If stb Then
    ret = "Select Case MsgBox("
Else
    ret = "MsgBox "
End If
buff2$ = Replace(Text1(0).Text, Chr$(34), String$(2, Chr$(34)))
buff$ = Dquote(buff2$)
strTitle = buff$
buff2$ = Replace(Text1(1).Text, Chr$(34), String$(2, Chr$(34)))
strMsg = locFormatMessage(buff2$)


j = List1.ItemData(List1.ListIndex)
Select Case List1.ItemData(List1.ListIndex)
    Case vbAbortRetryIgnore
        strButtons = "vbAbortRetryIgnore"
    Case vbOKCancel
        strButtons = "vbOKCancel"
    Case vbOKOnly
        strButtons = "vbOKOnly"
    Case vbRetryCancel
        strButtons = "vbRetryCancel"
    Case vbYesNo
        strButtons = "vbYesNo"
    Case vbYesNoCancel
        strButtons = "vbYesNoCancel"
End Select
If List3.ListIndex > 0 Then
    Select Case List3.ItemData(List3.ListIndex)
        Case vbCritical
            strButtons = strButtons & " + vbCritical"
        Case vbInformation
            strButtons = strButtons & " + vbCritical"
        Case vbQuestion
            strButtons = strButtons & " + vbCritical"
        Case vbExclamation
            strButtons = strButtons & " + vbCritical"
    End Select
End If
Select Case List2.ItemData(List2.ListIndex)
    Case vbDefaultButton1
            strButtons = strButtons & " + vbDefaultButton1"
    Case vbDefaultButton2
            strButtons = strButtons & " + vbDefaultButton2"
    Case vbDefaultButton3
            strButtons = strButtons & " + vbDefaultButton3"
    Case vbDefaultButton4
            strButtons = strButtons & " + vbDefaultButton4"
End Select
ret = ret & strMsg & "," & strButtons & "," & strTitle
If stb Then
    ret = ret & ") " & vbLf
    If j = vbOKOnly Then
        ret = ret & "   Case " & "vbOK" & vbLf & "      Msgbox " & Dquote("User Pressed Ok") & vbLf
    End If
    If j = vbOKCancel Then
        ret = ret & "   Case " & "vbOK" & vbLf & "      Msgbox " & Dquote("User Pressed Ok") & vbLf
        ret = ret & "   Case " & "vbCancel" & vbLf & "      Msgbox " & Dquote("User Pressed Cancel") & vbLf
    End If
    If j = vbYesNo Then
        ret = ret & "   Case " & "vbYes" & vbLf & "      Msgbox " & Dquote("User Pressed Yes") & vbLf
        ret = ret & "   Case " & "vbNo" & vbLf & "      Msgbox " & Dquote("User Pressed No") & vbLf
    End If
    If j = vbYesNoCancel Then
        ret = ret & "   Case " & "vbYes" & vbLf & "      Msgbox " & Dquote("User Pressed Yes") & vbLf
        ret = ret & "   Case " & "vbNo" & vbLf & "      Msgbox " & Dquote("User Pressed No") & vbLf
        ret = ret & "   Case " & "vbCancel" & vbLf & "      Msgbox " & Dquote("User Pressed Cancel") & vbLf
    End If
    If j = vbAbortRetryIgnore Then
        ret = ret & "   Case " & "vbRetry" & vbLf & "      Msgbox " & Dquote("User Pressed Retry") & vbLf
        ret = ret & "   Case " & "vbIgnore" & vbLf & "      Msgbox " & Dquote("User Pressed Ignore") & vbLf
        ret = ret & "   Case " & "vbAbort" & vbLf & "      Msgbox " & Dquote("User Pressed Abort") & vbLf
    End If
    If j = vbRetryCancel Then
        ret = ret & "   Case " & "vbRetry" & vbLf & "      Msgbox " & Dquote("User Pressed Retry") & vbLf
        ret = ret & "   Case " & "vbCancel" & vbLf & "      Msgbox " & Dquote("User Pressed Cancel") & vbLf
    End If
    ret = ret & "End Select"
End If
CodeGen = ret
End Function

Private Sub Form_Load()
Dim buff$
buff$ = "This Message Box wizard will let you create simple message boxes quickly.  " & _
"Enter data and select the options you want then click 'Generate Code'.  Click the " & _
"'Test Message Box' button to see a preview of your message box in action."
Label1(3).Caption = buff$
LoadFormData
End Sub
Private Sub LoadFormData()
Dim i As VbMsgBoxResult
Dim j As VbMsgBoxStyle
Clipboard.Clear
Text1(1).Text = "Enter the message you want to appear here"
Text1(0).SelStart = 0
Text1(0).SelLength = Len(Text1(0).Text)
Text1(1).SelStart = 0
Text1(1).SelLength = Len(Text1(1).Text)
List2.ListIndex = 0
List3.ListIndex = 1
List1.ListIndex = 0
List2.ItemData(0) = vbDefaultButton1
List2.ItemData(1) = vbDefaultButton2
List2.ItemData(2) = vbDefaultButton3
List2.ItemData(3) = vbDefaultButton4
List1.ItemData(0) = vbOKOnly
List1.ItemData(1) = vbOKCancel
List1.ItemData(2) = vbYesNo
List1.ItemData(3) = vbYesNoCancel
List1.ItemData(4) = vbRetryCancel
List1.ItemData(5) = vbAbortRetryIgnore

List3.ItemData(1) = vbInformation
List3.ItemData(2) = vbExclamation
List3.ItemData(3) = vbQuestion
List3.ItemData(4) = vbCritical



End Sub

