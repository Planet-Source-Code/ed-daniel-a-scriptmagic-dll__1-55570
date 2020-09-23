VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form MainCodeEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Script Initialization Code Editor"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10440
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainCodeEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   10440
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox SearchImg 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   10410
      TabIndex        =   4
      Top             =   420
      Visible         =   0   'False
      Width           =   10440
      Begin VB.CheckBox Check1 
         Caption         =   "Restrict to selected text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   7320
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   2835
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Replace All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5460
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Replace"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5460
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   420
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   5460
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox FindText 
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
         HideSelection   =   0   'False
         Index           =   1
         Left            =   1680
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   3555
      End
      Begin VB.TextBox FindText 
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
         HideSelection   =   0   'False
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   120
         Width           =   3555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Match Case"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   7320
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image CloseSearch 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   9300
         Picture         =   "MainCodeEditor.frx":01CA
         Top             =   0
         Width           =   240
      End
      Begin VB.Label SearchResult 
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7080
         TabIndex        =   13
         Top             =   780
         UseMnemonic     =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Replace with: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Find what: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":0548
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":065A
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":076C
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":0946
            Key             =   "CODE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":0B20
            Key             =   "FORM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":0CFA
            Key             =   "TOOLBOX"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":0ED4
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":10AE
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":1288
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":139A
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":14AC
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":15BE
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":16D0
            Key             =   "EXIT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":1818
            Key             =   "PROPERTIES"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":1B32
            Key             =   "WIZARD1"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainCodeEditor.frx":1F84
            Key             =   "FIND2"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Update project"
            ImageKey        =   "SAVE"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "FIND2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "WIZARD1"
            Object.ToolTipText     =   "Message Box Wizard"
            ImageKey        =   "WIZARD1"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CUT"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "CUT"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "COPY"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "PASTE"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Close window"
            ImageKey        =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   8265
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14499
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8580
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1755
   End
   Begin RichTextLib.RichTextBox CodeMain 
      Height          =   7875
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   13891
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"MainCodeEditor.frx":23D6
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnFileSave 
         Caption         =   "&Update Project"
      End
      Begin VB.Menu mnF3462 
         Caption         =   "-"
      End
      Begin VB.Menu mnFileUAndE 
         Caption         =   "Update and Exit"
      End
      Begin VB.Menu mnFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnEditSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnEditFind 
         Caption         =   "&Find (and or Replace)"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnUtils 
      Caption         =   "&Utilities"
      Begin VB.Menu mnUtilMsgBox 
         Caption         =   "&Message Box Wizard"
      End
   End
End
Attribute VB_Name = "MainCodeEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private IsLoaded As Boolean
Private isDirty As Boolean
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'wMSG For Find Line Position
Const EM_LINEINDEX = &HBB
Const WM_SETREDRAW = &HB

Dim bKey As Boolean
' True If The RTF Is Change
Dim bChange As Boolean
' Last Line Of RTF
Dim LastLine As Integer

' Color
Dim K_COLOR(1 To 2) As Long
Dim C_COLOR As Long
Dim Q_COlOR As Long
Dim N_Color As Long

Dim strDelimiter As String
Dim Delimiter(27) As String

Dim LastStart As Long

' Keyword

Public MyObjName As String
Public MyObjXML As String
Private MyXMLOBJ As QSXML


Private Sub CloseSearch_Click()
SearchOff
End Sub

Private Sub CodeMain_Change()
isDirty = True
bChange = True

' Update Color
Dim OStart As Long
Dim OLen As Long

Dim StartPos As Long
Dim EndPos As Long

Dim EndLine As Integer
Dim StartLine As Integer
Dim x As Long

Dim Text As String

With CodeMain

If .Text = "" Then Exit Sub

x = SendMessage(.hWnd, WM_SETREDRAW, 0, 0)

If LastStart > .SelStart Then
    EndLine = .GetLineFromChar(LastStart)
    StartLine = .GetLineFromChar(.SelStart)
Else
    StartLine = .GetLineFromChar(LastStart)
    EndLine = .GetLineFromChar(.SelStart)
End If

StartPos = SendMessage(.hWnd, EM_LINEINDEX, StartLine, 0&)
EndPos = SendMessage(.hWnd, EM_LINEINDEX, EndLine + 1, 0&)

If EndPos <= 0 Then EndPos = Len(.Text)

OStart = .SelStart
OLen = .SelLength

.SelStart = StartPos
.SelLength = EndPos - StartPos

.SelColor = N_Color
.SelBold = False
Text = .SelText
.SelRTF = ColorIt(Text)

.SelStart = OStart
.SelLength = OLen

LastStart = .SelStart

x = SendMessage(.hWnd, WM_SETREDRAW, 1, 0)
.Refresh

End With

End Sub

Private Sub CodeMain_Click()
If SearchImg.Visible Then
    SearchOff
End If
End Sub

Private Sub CodeMain_SelChange()
With CodeMain
SB1.Panels(3).Text = "LN: " & GetCurrentLine(CodeMain)
End With
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        If SearchImg.Visible Then
            SearchOff
            Exit Sub
        End If
        mnFileExit_Click
        Exit Sub
End Select
End Sub
Private Function SaveObject() As Boolean
Dim nd As Object, buff$
Dim ret As Boolean
With MyXMLOBJ
    Set nd = .GetRootElement()
    buff$ = Trim$(CodeMain.Text)
    buff$ = Replace(buff$, Chr$(171), "")
    buff$ = Replace(buff$, Chr$(187), "")
    nd.Text = buff$
    'frmCodeMain.strXMLToUpdate = .XML
    
    ret = frmCodeMain.UpdateInitialization("INITIALIZATION", .XML)
    If ret Then
        isDirty = False
    End If
    SaveObject = ret
End With
End Function

Private Sub Command2_Click(Index As Integer)
If FindText(0).Text = "" Then Exit Sub
Select Case Index
    Case 0
        SMFindText FindText(0).Text
    Case 1
        SMReplaceText FindText(0).Text, FindText(1).Text, False
    Case 2
        SMReplaceText FindText(0).Text, FindText(1).Text, True
End Select
End Sub
Private Sub SMReplaceText(strText2Find As String, strReplaceWith As String, bAllofThem As Boolean)
Dim i As Long
Dim selSt As Long
Dim stCount As Long
Dim lStart As Long
Dim lMax As Long
Dim cmSelstart As Long
Dim buff1$, buff2$, buff3$, buff4$

'Step 1  Make sure that they are not trying to replace same with same
If CBool(Check1(0).Value) Then 'Match Case
    If strText2Find = strReplaceWith Then
        MsgBox "Find Text and Replace Text must be different", vbInformation, "Error.."
        Exit Sub
    End If
Else
    If UCase$(strText2Find) = UCase$(strReplaceWith) Then
        MsgBox "Find Text and Replace Text must be different", vbInformation, "Error.."
        Exit Sub
    End If
End If

With CodeMain
cmSelstart = .SelStart
If CBool(Check1(1).Value) Then
    lStart = .SelStart
    lMax = lStart + .SelLength
Else
    lStart = 1
    lMax = Len(.Text)
End If
If bAllofThem Then
    'Count them first
    stCount = 0
    selSt = lStart
    If Not CBool(Check1(0).Value) Then
        i = InStr(selSt, UCase$(.Text), UCase$(strText2Find))
    Else
        i = InStr(selSt, .Text, strText2Find)
    End If
    Do While i >= lStart And i < lMax
        stCount = stCount + 1
        selSt = i + Len(strText2Find)
        If selSt >= lMax Then
            Exit Do
        End If
        If Not CBool(Check1(0).Value) Then
            i = InStr(selSt, UCase$(.Text), UCase$(strText2Find))
        Else
            i = InStr(selSt, .Text, strText2Find)
        End If
    Loop
    If stCount = 0 Then
        SearchResult.Caption = "String Not found"
        Exit Sub
    End If
    SearchResult.Caption = "Replaced " & stCount & " occurrences"
    If CBool(Check1(1).Value) Then
        buff1$ = .SelText
    Else
        buff1$ = .Text
    End If
    If Not CBool(Check1(0).Value) Then
        buff2$ = UCase$(buff1$)
        buff3$ = UCase$(strText2Find)
        buff4$ = UCase$(strReplaceWith)
    Else
        buff2$ = buff1$
        buff3$ = strText2Find
        buff4$ = strReplaceWith
    End If
    selSt = 1
    i = InStr(selSt, buff2$, buff3$)
    Do While i > 0
        buff1$ = MyStringReplace(buff1$, strReplaceWith, i, Len(strText2Find))
        buff2$ = MyStringReplace(buff2$, buff4$, i, Len(buff3$))
        selSt = i + IIf(Len(strText2Find) > Len(strReplaceWith), Len(strText2Find), Len(strReplaceWith))
        i = InStr(selSt, buff2$, buff3$)
    Loop
    If CBool(Check1(1).Value) Then
        .SelText = buff1$
        If cmSelstart < Len(.Text) Then
        .SelStart = cmSelstart
        End If
    Else
        .TextRTF = ColorIt(buff1$)
            If cmSelstart < Len(.Text) Then
               .SelStart = cmSelstart
            End If
    End If
    Exit Sub
End If

'First check if the currently selected text matches
    If Not CBool(Check1(0).Value) Then
        If UCase$(.SelText) = UCase$(strText2Find) Then
            .SelText = strReplaceWith
            'update lMax in case we have to continue
                SearchResult.Caption = "Replaced 1 occurrence"
                Exit Sub
        End If
    Else
        If .SelText = strText2Find Then
            .SelText = strReplaceWith
            If Not bAllofThem Then
                SearchResult.Caption = "Replaced 1 occurrence"
                Exit Sub
            End If
        End If
    End If
'Now continue


    selSt = .SelStart + 2
    If Not CBool(Check1(0).Value) Then
        i = InStr(selSt, UCase$(.Text), UCase$(strText2Find))
    Else
        i = InStr(selSt, .Text, strText2Find)
    End If
        If i >= lStart And i < lMax Then
            .SelStart = i - 1
            .SelLength = Len(strText2Find)
            .SelText = strReplaceWith
            SearchResult.Caption = "Replaced 1 occurrence"
        Else
            If selSt > 1 Then 'Check again from the top
                If Not CBool(Check1(0).Value) Then
                    i = InStr(lStart, UCase$(.Text), UCase$(strText2Find))
                Else
                    i = InStr(lStart, .Text, strText2Find)
                End If
                If i >= lStart And i < lMax Then
                    .SelStart = i - 1
                    .SelLength = Len(strText2Find)
                    .SelText = strReplaceWith
                    SearchResult.Caption = "Replaced 1 occurrence"
                Else
                    SearchResult.Caption = "Not Found!"
                End If
            Else
                SearchResult.Caption = "Not Found!"
            End If
        End If
End With

End Sub
Private Sub SMFindText(strText2Find As String)
Dim i As Long
Dim selSt As Long
Dim lStart As Long
Dim lMax As Long
With CodeMain
If CBool(Check1(1).Value) Then
    lStart = .SelStart
    lMax = lStart + .SelLength
Else
    lStart = 1
    lMax = Len(.Text)
End If
    selSt = .SelStart + 2
    If Not CBool(Check1(0).Value) Then
        i = InStr(selSt, UCase$(.Text), UCase$(strText2Find))
    Else
        i = InStr(selSt, .Text, strText2Find)
    End If
        If i >= lStart And i < lMax Then
            .SelStart = i - 1
            .SelLength = Len(strText2Find)
            SearchResult.Caption = "Found!"
        Else
            If selSt > lStart Then 'Check again from the top
                If Not CBool(Check1(0).Value) Then
                    i = InStr(UCase$(.Text), UCase$(strText2Find))
                Else
                    i = InStr(.Text, strText2Find)
                End If
                If i >= lStart And i < lMax Then
                    .SelStart = i - 1
                    .SelLength = Len(strText2Find)
                    SearchResult.Caption = "Found!"
                Else
                    SearchResult.Caption = "Not Found!"
                End If
            Else
                SearchResult.Caption = "Not Found!"
            End If
        End If
End With
End Sub

Private Sub FindText_Change(Index As Integer)
SearchResult.Caption = ""
End Sub

Private Sub FindText_GotFocus(Index As Integer)
With FindText(Index)
    If .Text <> "" Then
        .SelStart = 0
        .SelLength = Len(.Text)
    End If
End With
End Sub

Private Sub FindText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Select Case Index
    Case 0
        Command2_Click 0
    Case 1
        Command2_Click 1
End Select
End If
End Sub

Private Sub Form_Activate()
Dim nd As Object
If Not IsLoaded Then
    
    'This is only to get the Cancel property without having to monitor
    'KeyPress or KeyDown for the Esc key.
    'So we make the cancel button invisible by moving it off the screen
    Command1(0).Left = Command1(0).Width * -2
    IsLoaded = True
    With MyXMLOBJ
        Set nd = .GetRootElement()
        CodeMain.TextRTF = ColorIt(Trim$(nd.Text))
        CodeMain.SelStart = 0
        isDirty = False
        DoEvents
    End With
    SetLabel2
    Form_Resize
End If
End Sub
Private Sub SetLabel2()
If isDirty Then
Else
End If
End Sub
Private Sub Form_Load()
Dim i As Long
N_Color = vbBlack   '' Normal Text Color
C_COLOR = RGB(0, 150, 0) '' Comment Text Color
Q_COlOR = RGB(0, 128, 128) ''Quoation Text Color
K_COLOR(1) = RGB(0, 0, 200) '' SM_RESERVEDWORDS Color
K_COLOR(2) = RGB(128, 0, 64) '' Function Wrold Color
CodeMain.RightMargin = Me.TextWidth("A") * 3000
strDelimiter = ",(){}[]-+*%/='~!&|<>?:;.#@" & Chr(34) & vbTab
For i = 0 To Len(strDelimiter) - 1
    'Delimiter
    Delimiter(i) = Mid(strDelimiter, i + 1, 1)
    
    Select Case Delimiter(i)
        
        Case "\"
            Delimiter(i) = "\\"
        Case "}"
            Delimiter(i) = "\}"
        Case "{"
            Delimiter(i) = "\{"
        
    End Select
    
Next i

'Loading...
LastLine = -1
Set MyXMLOBJ = New QSXML
MyXMLOBJ.Initialize pavAUTO
MyXMLOBJ.OpenFromString MyObjXML

End Sub
Private Function ColorIt(Text As String) As String

Dim strLines() As String
Dim strLine As String
Dim strWord() As String
Dim intWord As Integer
Dim strWord1 As String

Dim strRTF As String
Dim strAllRTF As String
Dim strHeader As String

Dim onComment As Boolean
Dim onQuotation As Boolean

Dim i As Integer
Dim j As Integer

strLines = Split(Text, vbLf)

' Color
For i = LBound(strLines) To UBound(strLines)

    'Reset
    onComment = False
    onQuotation = False
    
    strLine = strLines(i)
    
    strLine = Replace(strLine, "\", "\\")
    strLine = Replace(strLine, "}", "\}")
    strLine = Replace(strLine, "{", "\{")
    
    ' Replace space to strline
    For j = 0 To 27
        
        strLine = Replace(strLine, Delimiter(j), Delimiter(j) & " ", , , vbTextCompare)
        
    Next j
    
    ' Split line to word
    strWord = Split(strLine, " ")
    
    For j = LBound(strWord) To UBound(strWord)
        
        Select Case UCase(strWord(j))
        
            ' Comment
            Case "'"
                
                If onQuotation = False Then
                    If onComment = False Then
                
                        onComment = True
                        strWord(j) = "\cf4 " & strWord(j)
                        
                        GoTo EndLine
                    
                    End If
                End If
            
            ' Quotation
            Case Chr(34)
            
                If onComment = False Then
                    If onQuotation = False Then
                
                        onQuotation = True
                        strWord(j) = "\cf5 " & strWord(j)
                        
                        GoTo EndIt
                
                    Else
                
                        onQuotation = False
                        strWord(j) = strWord(j) & "\cf0"
                        
                        GoTo EndIt
                
                    End If
                End If
                
            ' Comment
            Case "REM"
                
                If onQuotation = False Then
                    If onComment = False Then
                
                        onComment = True
                        strWord(j) = "\cf4 " & strWord(j)
                        
                        GoTo EndLine
                    
                    End If
                End If
            
            Case Else
                
                intWord = InStr(1, strDelimiter, Right(strWord(j), 1))
                
                If intWord > 0 Then
                
                    strWord1 = Delimiter(intWord - 1)
                    If Len(strWord(j)) <= 0 Then GoTo EndIt
                    strWord(j) = Left(strWord(j), Len(strWord(j)) - Len(strWord1))
                    
                End If
                
                If onQuotation = False Then
                
                    If InStr(1, SM_RESERVEDWORDS, " " & strWord(j) & " ", vbTextCompare) > 0 Then
                        strWord(j) = "\cf2\b1 " & strWord(j) & "\b0\cf0 "
                    End If
                    
                    If InStr(1, SM_FUNCTIONWORDS, " " & strWord(j) & " ", vbTextCompare) > 0 Then
                        strWord(j) = "\cf3 " & strWord(j) & "\cf0 "
                    End If
                    
                End If
                
                If intWord > 0 Then
                
                    ' Comment and Quotation
                    Select Case strWord1
                    
                        ' Comment
                        Case "'"
                            If onQuotation = False Then
                                If onComment = False Then
                
                                    onComment = True
                                    strWord1 = "\cf4 " & strWord1
                                    
                                    GoTo EndColor
                    
                                End If
                            End If
                            
                        'Quotation
                        Case Chr(34)
                        
                            If onComment = False Then
                                If onQuotation = False Then
                
                                    onQuotation = True
                                    strWord1 = "\cf5 " & strWord1
                                    
                                    GoTo EndColor
                        
                
                                Else
                
                                onQuotation = False
                                strWord1 = strWord1 & "\cf0"
                                
                                GoTo EndColor
                
                                End If
                            End If
                    
                    End Select
                    
EndColor:
                
                    strWord(j) = strWord(j) & strWord1
                    
                    If onComment = True Then
                        GoTo EndLine
                    End If
                    
                End If
                
        End Select
        
EndIt:
    
    Next j
EndLine:
    
    strLine = Join(strWord, " ")
    
    For j = 0 To 27
        
        strLine = Replace(strLine, Delimiter(j) & " ", Delimiter(j), , , vbTextCompare)
        
    Next j
    
    
    If onComment = True Then
        strLine = strLine & "\cf0"
    End If
    
    If onQuotation = True Then
        strLine = strLine & "\cf0"
    End If
    
    strLines(i) = strLine
    
Next i

strRTF = Join(strLines, vbLf & "\par ")
strHeader = CreateHeader

strAllRTF = strHeader & strRTF & vbLf & "}"

ColorIt = strAllRTF


End Function

Private Function CreateHeader() As String

Dim H1 As String
Dim H2 As String
Dim ColorH As String
Dim i As Integer

' Color Header
ColorH = "{\colortbl " & ConverColorToRTF(N_Color)
For i = 1 To 2
    ColorH = ColorH & ConverColorToRTF(K_COLOR(i))
Next i

ColorH = ColorH & ConverColorToRTF(C_COLOR)
ColorH = ColorH & ConverColorToRTF(Q_COlOR)
ColorH = ColorH & ";}"

' Header
H1 = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fnil\fcharset0 " & Me.Font.Name & ";}}"
H2 = "\viewkind4\uc1\pard\f0\fs" & Round(Me.Font.Size * 2) & " "

CreateHeader = H1 & vbLf & ColorH & vbLf & H2

End Function
Private Function ConverColorToRTF(LongColor As Long) As String

    Dim ColorRTFCode As String
    Dim lc As Long
    
    lc = LongColor And &H10000FF
    ColorRTFCode = ";\red" & lc
    lc = (LongColor And &H100FF00) / (2 ^ 8)
    ColorRTFCode = ColorRTFCode & "\green" & lc
    lc = (LongColor And &H1FF0000) / (2 ^ 16)
    ColorRTFCode = ColorRTFCode & "\blue" & lc
    ColorRTFCode = ColorRTFCode & ""
    
    ' Return Var
    ConverColorToRTF = ColorRTFCode
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If isDirty Then
    If MsgBox("Update Project with changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
        If Not SaveObject() Then
            Cancel = True
            Exit Sub
        End If
    End If
    isDirty = False
End If
End Sub

Private Sub Form_Resize()
Dim myClientHeight As Long
Dim myClientTop As Long
If SearchImg.Visible Then
    CloseSearch.Left = SearchImg.ScaleWidth - CloseSearch.Width
    myClientTop = SearchImg.Top + SearchImg.Height
Else
    myClientTop = Me.Toolbar1.Height
End If
myClientHeight = (Me.ScaleHeight - Me.SB1.Height) - myClientTop
CloseSearch.Top = 0
With CodeMain
    .Left = 0
    .Width = Me.ScaleWidth
    .Top = myClientTop
    .Height = myClientHeight
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
IsLoaded = False
isDirty = False
Set MyXMLOBJ = Nothing
End Sub

Private Sub Text1_Change()
SetLabel2
End Sub

Private Sub mnEditCopy_Click()
        With CodeMain
            If .SelLength > 0 Then
                Clipboard.Clear
                Clipboard.SetText .SelRTF, vbCFRTF
                Clipboard.SetText .SelText, vbCFText
            End If
        End With
        Exit Sub

End Sub

Private Sub mnEditCut_Click()
        With CodeMain
            If .SelLength > 0 Then
                Clipboard.Clear
                Clipboard.SetText .SelRTF, vbCFRTF
                Clipboard.SetText .SelText, vbCFText
                .SelText = ""
            End If
        End With
        Exit Sub
End Sub

Private Sub mnEditFind_Click()
If Not SearchImg.Visible Then
SearchOn
Else
SearchOff
End If
End Sub
Private Sub SearchOn()
    SearchImg.Visible = True
    SearchResult.Caption = ""
    FindText(0).TabStop = True
    FindText(1).TabStop = True
    Command2(0).TabStop = True
    Command2(1).TabStop = True
    Command2(2).TabStop = True
    Check1(0).TabStop = True
    Check1(1).TabStop = True
    CodeMain.TabStop = False
    FindText(0).TabIndex = 1
    Command2(0).TabIndex = 2
    FindText(1).TabIndex = 3
    Command2(1).TabIndex = 4
    Command2(2).TabIndex = 5
    Check1(0).TabIndex = 6
    Check1(1).TabIndex = 7
    If CodeMain.SelLength > 1 Then
        Check1(1).Enabled = True
    Else
        Check1(1).Value = vbUnchecked
        Check1(1).Enabled = False
    End If
    Form_Resize
    FindText(0).SetFocus
End Sub
Private Sub SearchOff()
    FindText(0).TabStop = False
    FindText(1).TabStop = False
    Command2(0).TabStop = False
    Command2(1).TabStop = False
    Command2(2).TabStop = False
    Check1(0).TabStop = False
    Check1(1).TabStop = False
    CodeMain.TabStop = True

    SearchImg.Visible = False
    Form_Resize
    CodeMain.SetFocus

End Sub

Private Sub mnEditPaste_Click()
Dim buff$
        With CodeMain
            buff$ = Clipboard.GetText(vbCFText)
            If Len(buff$) > 0 Then
                .SelText = buff$
            End If
        End With
Exit Sub
End Sub

Private Sub mnEditSelectAll_Click()
CodeMain.SelStart = 0
CodeMain.SelLength = Len(CodeMain.Text)
End Sub

Private Sub mnFileExit_Click()
If isDirty Then
    If MsgBox("Update Project with changes?", vbYesNo + vbQuestion, "Save Changes") = vbYes Then
        If Not SaveObject() Then
            Exit Sub
        End If
    End If
    isDirty = False
End If
Unload Me
End Sub

Private Sub mnFileSave_Click()
SaveObject
End Sub

Private Sub mnFileUAndE_Click()
If SaveObject() Then
    Unload Me
End If
End Sub

Private Sub mnUtilMsgBox_Click()
Dim buff$, str1 As String
Dim i As Long
frmMSGBoxBuilder.Show 1
buff$ = Clipboard.GetText(vbCFText)
If buff$ <> "" Then
    str1 = "Your message box code has been placed on the " & _
    "clipboard.  Would you like to paste it directly into the code window?"
    If MsgBox(str1, vbOKCancel, "Insert Message Box Code") = vbCancel Then
        Exit Sub
    End If
    buff$ = vbLf & buff$ & vbLf
    With CodeMain
        i = .SelStart
        If .SelLength = 0 Then
            If .SelStart > 0 Then
                str1 = Left$(.Text, .SelStart) & buff$ & Mid$(.Text, .SelStart + 1)
            Else
                str1 = buff$ & .Text
            End If
            .Text = str1
        Else
        .SelText = buff$
        End If
        .SelStart = Len(.Text)
        .TextRTF = ColorIt(.Text)
        .SelStart = i + 1
    End With
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "FIND"
        mnEditFind_Click
        Exit Sub
    Case "EXIT"
        mnFileExit_Click
        Exit Sub
    Case "SAVE"
        mnFileSave_Click
        Exit Sub
    Case "CUT"
        mnEditCut_Click
        Exit Sub
    Case "COPY"
        mnEditCopy_Click
        Exit Sub
    Case "PASTE"
        mnEditPaste_Click
        Exit Sub
    Case "WIZARD1"
        mnUtilMsgBox_Click
        Exit Sub
End Select
End Sub
