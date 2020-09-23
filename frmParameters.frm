VERSION 5.00
Begin VB.Form frmParameters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Run-Time Parameters"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmParameters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Index           =   0
      Left            =   2460
      TabIndex        =   13
      Top             =   6960
      Width           =   1995
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   11
      Left            =   1860
      TabIndex        =   12
      Top             =   6480
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   10
      Left            =   1860
      TabIndex        =   11
      Top             =   6060
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   9
      Left            =   1860
      TabIndex        =   10
      Top             =   5640
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   8
      Left            =   1860
      TabIndex        =   9
      Top             =   5220
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   7
      Left            =   1860
      TabIndex        =   8
      Top             =   4800
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   6
      Left            =   1860
      TabIndex        =   7
      Top             =   4380
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   5
      Left            =   1860
      TabIndex        =   6
      Top             =   3960
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   4
      Left            =   1860
      TabIndex        =   5
      Top             =   3540
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   3
      Left            =   1860
      TabIndex        =   4
      Top             =   3120
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   2
      Left            =   1860
      TabIndex        =   3
      Top             =   2700
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   1
      Left            =   1860
      TabIndex        =   2
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox ParmData 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   1
      Top             =   1860
      Width           =   4935
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   0
      Left            =   0
      Picture         =   "frmParameters.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   1
      Left            =   6360
      Picture         =   "frmParameters.frx":0884
      Stretch         =   -1  'True
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label3 
      Caption         =   "**IMPORTANT** If your function expects a parameter to be a STRING you must enclose it in double quotes EXAMPLE: ""Script Magic"""
      Height          =   375
      Left            =   60
      TabIndex        =   29
      Top             =   1020
      Width           =   6675
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   11
      Left            =   60
      TabIndex        =   28
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   10
      Left            =   60
      TabIndex        =   27
      Top             =   6060
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   9
      Left            =   60
      TabIndex        =   26
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   8
      Left            =   60
      TabIndex        =   25
      Top             =   5220
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   60
      TabIndex        =   24
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   6
      Left            =   60
      TabIndex        =   23
      Top             =   4380
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   5
      Left            =   60
      TabIndex        =   22
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   60
      TabIndex        =   21
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   3
      Left            =   60
      TabIndex        =   20
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   2
      Left            =   60
      TabIndex        =   19
      Top             =   2700
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   60
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label ParmNames 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   17
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Parameter Value"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1860
      TabIndex        =   16
      Top             =   1500
      Width           =   4875
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Parameter Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   15
      Top             =   1500
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Script objects that require parameters to run"
      Height          =   255
      Left            =   480
      TabIndex        =   14
      Top             =   360
      Width           =   6015
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strInitialFuncName As String
Private IsLoaded As Boolean
Private isLoading As Boolean
Private selIDX As Long

Public Sub PrepareToShow()
Dim i As Long
Dim j As Long
    ClearData
    j = Combo1.ListIndex
    Combo1.Clear
    If UBound(MyLocalFunctions) < 0 Then Exit Sub
    For i = 0 To UBound(MyLocalFunctions)
        If MyLocalFunctions(i).bActive Then
        Combo1.AddItem MyLocalFunctions(i).stName
        Combo1.ItemData(Combo1.NewIndex) = i
        End If
    Next
    If j < Combo1.ListCount Then
        Combo1.ListIndex = j
    End If
End Sub
Public Sub Initialize()
'ReDim MyLocalFunctions(-1) As MyParametersType
ReDim MyLocalFunctions(0) As MyParametersType
ClearData
End Sub
Private Sub ClearData()
Dim i As Long
isLoading = True
selIDX = -1
For i = 0 To ParmData.UBound
    ParmNames(i).Caption = "Parameter " & i + 1
    ParmData(i).Text = ""
    ParmData(i).Enabled = False
    ParmData(i).BackColor = Me.BackColor
Next
isLoading = False
End Sub

Private Sub Combo1_Click()
Dim i As Long, j As Long
ReDim ed1(0) As String
ReDim ed2(0) As String
ClearData
With Combo1
    If .ListIndex < 0 Then Exit Sub
    i = .ItemData(.ListIndex)
End With
selIDX = i
With MyLocalFunctions(i)
    If .lParmCount = 0 Then Exit Sub
    If .stParmValues = "" Then
        If .lParmCount > 1 Then
            .stParmValues = String$(.lParmCount - 1, Chr$(0))
        End If
    End If
    ed1 = Split(.stParmValues, Chr$(0))
    ed2 = Split(.stParms, ",")
    If UBound(ed1) = -1 Then
        ReDim ed1(UBound(ed2)) As String
    End If
    isLoading = True
    For i = 0 To UBound(ed2)
        ParmNames(i).Caption = Trim$(ed2(i))
        ParmNames(i).ToolTipText = Trim$(ed2(i))
        ParmData(i).BackColor = vbWhite
        ParmData(i).Enabled = True
        ParmData(i).Text = Trim$(ed1(i))
    Next
    isLoading = False
End With
End Sub
Private Sub GatherData()
If selIDX < 0 Then Exit Sub
If isLoading Then Exit Sub
Dim i As Long
Dim j As Long
Dim buff$
With MyLocalFunctions(selIDX)
    For i = 0 To ParmData.UBound
        If Not ParmData(i).Enabled Then Exit For
        If i = 0 Then
            buff$ = Trim$(ParmData(i))
        Else
            buff$ = buff$ & Chr$(0) & Trim$(ParmData(i))
        End If
    Next
    .stParmValues = buff$
End With
End Sub
Private Sub Command1_Click(Index As Integer)
Unload Me
End Sub

Private Sub Form_Activate()
Dim i As Long
Dim j As Long
If Not IsLoaded Then
    ClearData
    If UBound(MyLocalFunctions) < 0 Then Exit Sub
    For i = 0 To UBound(MyLocalFunctions)
        If MyLocalFunctions(i).bActive Then
        Combo1.AddItem MyLocalFunctions(i).stName
        Combo1.ItemData(Combo1.NewIndex) = i
        End If
    Next
    If Combo1.ListCount = 0 Then
        MsgBox "There are no public subroutines or functions that require parameters.", vbInformation, "Sub/Function Parameters"
        
    End If
        Combo1.ListIndex = InListBox(Combo1, Me.strInitialFuncName)
    IsLoaded = True
End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
IsLoaded = False
End Sub

Private Sub ParmData_Change(Index As Integer)
GatherData
End Sub
