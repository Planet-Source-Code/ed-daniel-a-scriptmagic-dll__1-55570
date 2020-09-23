VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmConstants 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add / Edit Public Constants"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConstants.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6435
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Add To List"
      Enabled         =   0   'False
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1740
      Width           =   6315
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Number"
      Height          =   210
      Index           =   1
      Left            =   3840
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "String"
      Height          =   210
      Index           =   0
      Left            =   2580
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Value           =   -1  'True
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   435
      Index           =   1
      Left            =   4860
      TabIndex        =   5
      Top             =   5280
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   60
      MaxLength       =   30
      TabIndex        =   0
      Top             =   840
      Width           =   1875
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   2760
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":0442
            Key             =   "PROJECT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":0894
            Key             =   "CODE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":0CE6
            Key             =   "BUTTON"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":0EC0
            Key             =   "SUBROUTINE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":1312
            Key             =   "SUBROUTINES"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":1764
            Key             =   "FUNCTIONS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":1BB6
            Key             =   "CLASS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":2148
            Key             =   "API"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":259A
            Key             =   "TYPEDEFS"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":29EC
            Key             =   "ENUM"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":2E3E
            Key             =   "VARIABLE"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":3290
            Key             =   "ITEM"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":36E2
            Key             =   "CONSTANTS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConstants.frx":3B34
            Key             =   "INPUT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TV1 
      Height          =   2955
      Left            =   60
      TabIndex        =   3
      Top             =   2100
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5212
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   1
      Left            =   5940
      Picture         =   "frmConstants.frx":3D0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   0
      Left            =   0
      Picture         =   "frmConstants.frx":4150
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Public Const = """""
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1260
      UseMnemonic     =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Value"
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   8
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1980
      TabIndex        =   7
      Top             =   780
      Width           =   555
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmConstants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private IsLoaded As Boolean
Private MyXMLOBJ As QSXML
Public strXML As String


Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        If frmCodeMain.UpdateConstants(MyXMLOBJ.XML) Then
            Unload Me
            Exit Sub
        End If
    Case 1
        Unload Me
End Select
End Sub

Private Sub Command2_Click()
Dim buff$
Dim nd As Object
Dim ob1 As QSXML
If Not locValidate() Then
    Exit Sub
End If
If InList(Text1(0).Text) Then
    MsgBox Text1(0).Text & " is already defined."
    Text1(0).SetFocus
    Exit Sub
End If
Set ob1 = New QSXML
ob1.Initialize pavAUTO
ob1.OpenFromString "<CONSTANT></CONSTANT>"
With ob1
    Set nd = .GetRootElement()
    .SetAttribute nd, "NAME", Text1(0).Text
    .SetAttribute nd, "VALUE", Text1(1).Text
    .SetAttribute nd, "TYPE", IIf(Option1(0).Value, "STRING", "NUMBER")
    buff$ = .XML
End With
Set ob1 = Nothing
With MyXMLOBJ
    Set nd = .GetRootElement
    .XMLAddNode nd, buff$
    .SetAttribute nd, "COUNT", CStr(TV1.Nodes.Count + 1)
End With
LoadTree
Text1(0).Text = ""
Text1(1).Text = ""
Text1(0).SetFocus
Exit Sub
End Sub
Private Function InList(strVal) As Boolean
Dim buff$, i As Long
buff$ = "NAME=" & Chr$(34) & strVal & Chr$(34)
With TV1
For i = 1 To .Nodes.Count
    If InStr(UCase$(.Nodes(i).Tag), UCase$(buff$)) > 0 Then
        InList = True
        Exit Function
    End If
Next
InList = False
Exit Function
End With
End Function
Private Function locValidate() As Boolean
Dim buff$
Text1(0).Text = Trim$(Text1(0).Text)
If Option1(1).Value Then
    Text1(1).Text = Trim$(Text1(1).Text)
End If
If Text1(0).Text = "" Then
    MsgBox "Enter a variable name"
    Text1(0).SetFocus
    locValidate = False
    Exit Function
End If
If Option1(1).Value And Text1(1).Text = "" Then
    MsgBox "Enter a number value"
    Text1(1).SetFocus
    locValidate = False
    Exit Function
End If
If Option1(1).Value Then
    buff$ = NumFormat(Text1(1).Text, ".-")
    If buff$ <> Text1(1).Text Then
        MsgBox "Only numeric values allowed."
        Text1(1).SetFocus
        locValidate = False
        Exit Function
    End If
    On Error Resume Next
    If CDbl(buff$) > 0 Then
    Else
    End If
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Only numeric values allowed."
        Text1(1).SetFocus
        locValidate = False
        Exit Function
    End If
End If
If InStr(CALPHA, UCase$(Left$(Text1(0).Text, 1))) = 0 Then
        MsgBox "Constant variable names must begin with A-Z or a-z"
        Text1(0).SetFocus
        locValidate = False
        Exit Function
End If
buff$ = AlphaNumFormat(Text1(0).Text, "_")
If buff$ <> Text1(0).Text Then
        MsgBox "Constant variable names may contain only alphanumeric characters (and _)"
        Text1(0).SetFocus
        locValidate = False
        Exit Function
End If
If frmCodeMain.ItemExists(Text1(0).Text) Then
        MsgBox "A public object called '" & Text1(0).Text & "' already exists.", vbInformation, "Error.."
        Text1(0).SetFocus
        locValidate = False
        Exit Function
End If
locValidate = True
End Function
Private Sub Form_Load()
Set MyXMLOBJ = New QSXML
MyXMLOBJ.Initialize pavAUTO
MyXMLOBJ.OpenFromString strXML
LoadTree
End Sub
Private Sub LoadTree()
Dim nd As Object
Dim ndl As Object
Dim nod1 As Node
Dim i As Long, buff$
With TV1
.Nodes.Clear
End With
With MyXMLOBJ
    Set nd = .GetRootElement
    If .GetAttributeValue(nd, "COUNT") = "0" Then
        Exit Sub
    End If
    Set ndl = .GetRootChildren()
    For i = 0 To ndl.length - 1
        buff$ = "Public Const " & .GetAttributeValue(ndl(i), "NAME") & " = "
        If .GetAttributeValue(ndl(i), "TYPE") = "NUMBER" Then
        buff$ = buff$ & .GetAttributeValue(ndl(i), "VALUE")
        Else
        buff$ = buff$ & Dquote(.GetAttributeValue(ndl(i), "VALUE"))
        End If
        Set nod1 = TV1.Nodes.Add(, , , buff$, "CONSTANTS", "CONSTANTS")
        nod1.Tag = ndl(i).XML
    Next
End With
End Sub

Private Sub Option1_Click(Index As Integer)
SetLabel2
End Sub

Private Sub Text1_Change(Index As Integer)
SetLabel2
End Sub
Private Sub SetLabel2()
Dim buff$
buff$ = "Public Const " & Trim$(Text1(0).Text) & " = "
If Option1(0).Value Then
    buff$ = buff$ & Chr$(34) & Text1(1).Text & Chr$(34)
Else
    buff$ = buff$ & Trim$(Text1(1).Text)
End If
Label1(3).Caption = buff$
If Text1(0).Text <> "" And Text1(1).Text <> "" Then
    Command2.Enabled = True
Else
    Command2.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If Command2.Enabled Then
        Command2_Click
    End If
End If
End Sub

Private Sub TV1_KeyDown(KeyCode As Integer, Shift As Integer)
With TV1
If KeyCode = vbKeyDelete Then
    If (.SelectedItem Is Nothing) Then
        Exit Sub
    End If
    If MsgBox("Remove " & .SelectedItem.Text & "?", vbYesNo + vbQuestion, "Remove Object") = vbNo Then
        Exit Sub
    End If
    DeleteNode .SelectedItem.Tag
End If
End With

End Sub
Private Function DeleteNode(stNDXML As String) As Boolean
Dim ob1 As New QSXML
Dim nd As Object
Dim ndc As Object
Dim obName As String
ob1.Initialize pavAUTO
ob1.OpenFromString stNDXML
Set nd = ob1.GetRootElement()
obName = ob1.GetAttributeValue(nd, "NAME")
With MyXMLOBJ
    Set nd = .GetNodeFromAttribute("CONSTANT", "NAME", obName)
    If .RemoveNode(nd) Then
        Set nd = .GetRootElement()
        .SetAttribute nd, "COUNT", CStr(TV1.Nodes.Count - 1)
        LoadTree
    End If
End With

End Function

