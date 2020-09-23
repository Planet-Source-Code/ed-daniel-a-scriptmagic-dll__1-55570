VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCodeMain 
   Caption         =   "Script Magic"
   ClientHeight    =   6570
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCodeMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox SearchImg 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   780
      ScaleHeight     =   705
      ScaleWidth      =   9555
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   9585
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
         Left            =   5460
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   1335
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
         Left            =   1020
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   120
         Width           =   2655
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
         Left            =   3780
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   60
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Edit"
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
         Left            =   3780
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
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
         Left            =   60
         TabIndex        =   21
         Top             =   120
         Width           =   975
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
         Left            =   5520
         TabIndex        =   20
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   2715
      End
      Begin VB.Image CloseSearch 
         Appearance      =   0  'Flat
         Height          =   210
         Left            =   9300
         Picture         =   "frmCodeMain.frx":014A
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3420
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":04C8
            Key             =   "PROJECT"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":091A
            Key             =   "CODE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":0D6C
            Key             =   "BUTTON"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":0F46
            Key             =   "SUBROUTINE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":1398
            Key             =   "SUBROUTINES"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":17EA
            Key             =   "FUNCTIONS"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":1C3C
            Key             =   "CLASS"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":21CE
            Key             =   "API"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":2620
            Key             =   "TYPEDEFS"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":2A72
            Key             =   "ENUM"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":2EC4
            Key             =   "VARIABLE"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":3316
            Key             =   "ITEM"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":3768
            Key             =   "CONSTANTS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":3BBA
            Key             =   "INPUT"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":3D94
            Key             =   "FUNCTION"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3300
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4326
            Key             =   "NEW"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4438
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":454A
            Key             =   "RUN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4724
            Key             =   "CODE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":48FE
            Key             =   "FORM"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4AD8
            Key             =   "TOOLBOX"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4CB2
            Key             =   "STOP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":4E8C
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":5066
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":5178
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":528A
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":539C
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":54AE
            Key             =   "EXIT"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":55F6
            Key             =   "PROPERTIES"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCodeMain.frx":5910
            Key             =   "FIND2"
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox CodeMain 
      Height          =   4755
      Left            =   3960
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Double click a section to edit that section"
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8387
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmCodeMain.frx":5D62
   End
   Begin MSScriptControlCtl.ScriptControl SC2 
      Left            =   3300
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.PictureBox SilentWindow 
      Height          =   5775
      Left            =   180
      ScaleHeight     =   5715
      ScaleWidth      =   10275
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin RichTextLib.RichTextBox EXEResult 
         Height          =   2055
         Left            =   3600
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   3120
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmCodeMain.frx":5DE2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView ExeTree 
         Height          =   5535
         Left            =   60
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   9763
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "ImageList2"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ListBox exeList1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   3600
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   6615
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   3120
         Sorted          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label exeSilentLabel 
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
         Height          =   675
         Left            =   3600
         TabIndex        =   11
         Top             =   2280
         Width           =   6555
      End
      Begin VB.Label exeLabel1 
         Caption         =   "Debugger Output"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   2880
         Width           =   6555
      End
      Begin VB.Label exeLabel1 
         Caption         =   "Public Variables"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   9
         Top             =   60
         Width           =   6555
      End
   End
   Begin RichTextLib.RichTextBox ScratchRTF 
      Height          =   1275
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   2249
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCodeMain.frx":5E59
   End
   Begin VB.PictureBox imgSplitter 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3960
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4800
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   720
      Width           =   72
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6315
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   14975
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   3900
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "Create a new project"
            ImageKey        =   "NEW"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open an existing project"
            ImageKey        =   "OPEN"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "FIND"
            Object.ToolTipText     =   "Search"
            ImageKey        =   "FIND2"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PROPERTIES"
            Object.ToolTipText     =   "Show project properties"
            ImageKey        =   "PROPERTIES"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save project"
            ImageKey        =   "SAVE"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy selected object"
            ImageKey        =   "COPY"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste an object"
            ImageKey        =   "PASTE"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Key             =   "SCRIPT"
            Object.ToolTipText     =   "Immediate window"
            ImageKey        =   "CODE"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RUN"
            Object.ToolTipText     =   "Execute script"
            ImageKey        =   "RUN"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "STOP"
            Object.ToolTipText     =   "Stop script execution"
            ImageKey        =   "STOP"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Close window"
            ImageKey        =   "EXIT"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   5535
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnFileURL 
         Caption         =   "TestMe"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnFileNew 
         Caption         =   "&New Project"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnFileOpenProject 
         Caption         =   "&Import Project"
      End
      Begin VB.Menu mnFileSaveProject 
         Caption         =   "&Save Project"
      End
      Begin VB.Menu mnFileSaveAs 
         Caption         =   "Export &Project to Disk"
      End
      Begin VB.Menu mnFileRename 
         Caption         =   "&Rename Project"
      End
      Begin VB.Menu mnSep473254 
         Caption         =   "-"
      End
      Begin VB.Menu mnFIleEXPMain 
         Caption         =   "&Export Project to VBScript File (.vbs)"
         Begin VB.Menu mnFileExpFull 
            Caption         =   "Include Comments and Blank Lines"
         End
         Begin VB.Menu mnFileExport 
            Caption         =   "Remove Comments and Blank Lines"
         End
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnProject 
      Caption         =   "&Project"
      Begin VB.Menu mnFileTest 
         Caption         =   " &Properties"
      End
      Begin VB.Menu mn2356s 
         Caption         =   "-"
      End
      Begin VB.Menu mnEditInitialization 
         Caption         =   "Add / Edit Initialization Code"
      End
      Begin VB.Menu mnAddVars 
         Caption         =   "Add / Edit Public Variables"
      End
      Begin VB.Menu mnAddEditConstants 
         Caption         =   "Add / Edit Public Constants"
      End
      Begin VB.Menu mnSep09u8 
         Caption         =   "-"
      End
      Begin VB.Menu mnAddSub 
         Caption         =   "Add SubRoutine"
      End
      Begin VB.Menu mnAddFunction 
         Caption         =   "Add Function"
      End
      Begin VB.Menu mnAddClass 
         Caption         =   "Add Class"
      End
      Begin VB.Menu mn873D 
         Caption         =   "-"
      End
      Begin VB.Menu mnProjectFind 
         Caption         =   "&Find"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnDebugRun 
         Caption         =   "&Run Script"
      End
      Begin VB.Menu mnSet59y 
         Caption         =   "-"
      End
      Begin VB.Menu mnViewParms 
         Caption         =   "View / Edit SubRoutine and Function Parameters"
      End
      Begin VB.Menu mnSre46 
         Caption         =   "-"
      End
      Begin VB.Menu mnDebugStop 
         Caption         =   "&Stop Running"
      End
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnhelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnSep0866 
         Caption         =   "-"
      End
      Begin VB.Menu mnHelpWeb 
         Caption         =   "On the Web"
         Begin VB.Menu mnMSDN 
            Caption         =   "MSDN"
         End
         Begin VB.Menu mnMSExamples 
            Caption         =   "Microsoft Script Examples"
         End
         Begin VB.Menu mnWeb1 
            Caption         =   "www.winguides.com"
         End
         Begin VB.Menu mnWeb2 
            Caption         =   "Adaptive.net Online VBScript Reference"
         End
         Begin VB.Menu mnWeb4 
            Caption         =   "FunctionX.com"
         End
         Begin VB.Menu mnWeb3 
            Caption         =   "Online VBSTutor"
         End
         Begin VB.Menu mnWeb5 
            Caption         =   "W3Schools Tutorials and Examples"
         End
         Begin VB.Menu mnWeb6 
            Caption         =   "www.VisualBasicScript.com"
         End
         Begin VB.Menu mnWeb7 
            Caption         =   "Devguru.com"
         End
         Begin VB.Menu mnWind35 
            Caption         =   "Winscripter.com"
         End
      End
   End
End
Attribute VB_Name = "frmCodeMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private dbgMainXML As QSXML
Public strXMLToUpdate As String
Private WithEvents myScriptOBJ As ScriptEngine
Attribute myScriptOBJ.VB_VarHelpID = -1
Private SC1

Private WithEvents m_Events As SM_Event
Attribute m_Events.VB_VarHelpID = -1
Private bRunningSilent As Boolean
Private bPasswordIsRequired As Boolean
Private strPassword As String
Private myOpenPath As String
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
'My Stuff (Ed Daniel)
Private Type SM_RunErrorType
    errNumber As Long
    errDesc As String
    errLine As Long
End Type
Private MyError As SM_RunErrorType
Const sglSplitLimit = 500
Private mbMoving As Boolean
'Private dbgMainXML As QSXML
Public Function GetProjectXML() As String
    GetProjectXML = dbgMainXML.XML
End Function

Private Sub CloseSearch_Click()
SearchOff
End Sub

Private Sub CodeMain_Change()
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

Private Sub CodeMain_DblClick()
ShowSection
End Sub
Private Sub ShowSection()
Dim i As Long
Dim sectStart As Long
Dim sectEnd As Long
Dim oldSel As Long
Dim objName As String
Dim ndToEdit As Object
Dim buff$
Dim prjBuff$
With CodeMain
oldSel = .SelStart
Do While True
LOOPHERE:
    i = InStrRev(.Text, "'" & Chr$(171), oldSel)
    If i = 0 Then Exit Sub
    sectStart = i + 2
    i = InStr(sectStart, .Text, Chr$(187))
    If i = 0 Then Exit Sub
    sectEnd = i
    buff$ = Trim$(Mid$(.Text, sectStart, sectEnd - sectStart))
    If Left$(buff$, 4) = "END " Then
        oldSel = sectStart - 5
        GoTo LOOPHERE
    End If
    Exit Do
Loop
Select Case UCase$(buff$)
    Case "INITIALIZATION CODE"
        mnEditInitialization_Click
        Exit Sub
    Case "PUBLIC VARIABLES"
        mnAddVars_Click
        Exit Sub
    Case "PUBLIC CONSTANTS"
        mnAddEditConstants_Click
        Exit Sub
    Case "ALL SUBROUTINES"
        If MsgBox("Do you wish to create a new SubRoutine?", vbYesNo + vbQuestion, "Create Object") = vbYes Then
            mnAddSub_Click
            Exit Sub
        End If
        Exit Sub
    Case "ALL FUNCTIONS"
        If MsgBox("Do you wish to create a new Function?", vbYesNo + vbQuestion, "Create Object") = vbYes Then
            mnAddFunction_Click
            Exit Sub
        End If
        Exit Sub
    Case "ALL CLASSES"
        If MsgBox("Do you wish to create a new Class Object?", vbYesNo + vbQuestion, "Create Object") = vbYes Then
            mnAddClass_Click
            Exit Sub
        End If
        Exit Sub
End Select

    
If Left$(buff$, 9) = "Project: " Then
    mnFileTest_Click
    Exit Sub
End If
If InStr(buff$, "SUBROUTINE: ") > 0 Then
    objName = Trim$(Mid$(buff$, Len("SUBROUTINE:") + 1))
    Set ndToEdit = GetItemNode(objName)
    If (ndToEdit Is Nothing) Then
        MsgBox "Error retrieving object: " & objName
        Exit Sub
    End If
    EditObject ndToEdit
    Exit Sub
End If
If InStr(buff$, "FUNCTION: ") > 0 Then
    objName = Trim$(Mid$(buff$, Len("FUNCTION:") + 1))
    Set ndToEdit = GetItemNode(objName)
    If (ndToEdit Is Nothing) Then
        MsgBox "Error retrieving object: " & objName
        Exit Sub
    End If
    EditObject ndToEdit
    Exit Sub
End If
If InStr(buff$, "CLASS: ") > 0 Then
    objName = Trim$(Mid$(buff$, Len("CLASS:") + 1))
    Set ndToEdit = GetItemNode(objName)
    If (ndToEdit Is Nothing) Then
        MsgBox "Error retrieving object: " & objName
        Exit Sub
    End If
    EditObject ndToEdit
    Exit Sub
End If

End With

End Sub
Private Sub EditObject(ndObj As Object)
Dim buff$
With dbgMainXML
    buff$ = .GetAttributeValue(ndObj, "NAME")
    'SubCodeEditor.MyObjXML = ndObj.XML
    'SubCodeEditor.Show 1
    ObjectCodeEditor.MyObjName = buff$
    ObjectCodeEditor.Show 1
End With

End Sub

Private Sub CodeMain_SelChange()
With sb1
.Panels(3).Text = "Ln: " & GetCurrentLine(CodeMain)
End With
End Sub
Private Function vbProcIDX(strFuncName) As Long
Dim i As Long
With SC1
    For i = 1 To .Procedures.Count
        If .Procedures(i).Name = strFuncName Then
            vbProcIDX = i
            Exit Function
        End If
    Next
    vbProcIDX = -1
End With
End Function
Private Sub evalFunction(strFuncName As String)
Dim nd As Object
Dim buff$
Dim procIDX As Long
Dim parms As String
Dim prcString As String
ReDim ed1(0)
procIDX = vbProcIDX(strFuncName)
If procIDX < 0 Then
    MsgBox "Error no procedure defined"
    Exit Sub
End If
On Error GoTo ERRHDL
With dbgMainXML
        Set nd = Me.GetItemNode(strFuncName)
        parms = .GetAttributeValue(nd, "PARAMETERS")
        If parms = "" Then
            buff$ = nd.nodename & " " & strFuncName & "()"
            myScriptOBJ.Echo "Executing: " & buff$
            buff$ = nd.nodename & " " & strFuncName & "() Executed"
            If nd.nodename = "FUNCTION" Then
                buff$ = buff$ & "Return Value = " & CStr(SC1.Eval(strFuncName))
            Else
                SC1.ExecuteStatement strFuncName
            End If
            myScriptOBJ.Echo buff$
            EvalVariables
        Else
            If Not ValidParms(strFuncName) Then
                frmParameters.strInitialFuncName = strFuncName
                mnViewParms_Click
                Exit Sub
            End If
            parms = GetParmString(strFuncName)
            If parms = "" Then
                MsgBox "Error getting object parameters"
                Exit Sub
            End If
            prcString = strFuncName & "(" & parms & ")"
            
            buff$ = nd.nodename & " " & prcString
            myScriptOBJ.Echo "Executing: " & buff$
            buff$ = buff$ & " Executed"
            If SC1.Procedures(procIDX).HasReturnValue Then
                buff$ = buff$ & vbCrLf & "Return Value = " & CStr(SC1.Eval(prcString))
            Else
                SC1.ExecuteStatement "Call " & prcString
            End If
            myScriptOBJ.Echo buff$
            EvalVariables
            
'            MsgBox buff$
'            buff$ = nd.nodename & " " & strFuncName & " expects parameters." & vbCrLf & _
'            vbCrLf & "Parms = '" & parms & "'"
'            myscriptobj.Echo buff$
        End If
        
End With
Exit Sub
ERRHDL:
SC1_Error
Exit Sub
End Sub

Private Sub Command2_Click(Index As Integer)
If FindText(0).Text = "" Then Exit Sub
Select Case Index
    Case 0
        SMFindText FindText(0).Text
        Exit Sub
    Case 1
        CodeMain_DblClick
        Exit Sub
End Select

End Sub

Private Sub ExeTree_DblClick()
Dim buff$
On Error GoTo ERRHDL
With ExeTree
    If (.SelectedItem Is Nothing) Then Exit Sub
    If .SelectedItem.Tag <> "" Then
        Select Case .SelectedItem.Parent.Key
            Case "CONSTANTS"
                On Error Resume Next
                buff$ = .SelectedItem.Text & " = " & CStr(SC1.Eval(.SelectedItem.Text) & "")
                If Err.Number <> 0 Then
                    buff$ = "Error evaluating object {" & .SelectedItem.Text & "}." & vbCrLf & Err.Description
                    Err.Clear
                End If
                myScriptOBJ.Echo buff$
                On Error GoTo ERRHDL
            Case "VARIABLES"
                On Error Resume Next
                buff$ = .SelectedItem.Text & " = " & CStr(SC1.Eval(.SelectedItem.Text) & "")
                If Err.Number <> 0 Then
                    buff$ = "Error evaluating object {" & .SelectedItem.Text & "}." & vbCrLf & Err.Description
                    Err.Clear
                End If
                myScriptOBJ.Echo buff$
                On Error GoTo ERRHDL
            Case "FUNCTIONS"
                evalFunction .SelectedItem.Text
                If MyError.errNumber <> 0 Then GoTo ERRHDL
            Case "SUBROUTINES"
                evalFunction .SelectedItem.Text
                If MyError.errNumber <> 0 Then GoTo ERRHDL
        End Select
    End If
End With
Exit Sub
ERRHDL:
If MyError.errNumber <> 0 Then
    Err.Clear
    ProcessMyError False
Else
    MsgBox Err.Description
    Err.Clear
    mnDebugStop_Click
End If
End Sub

Private Sub SMFindText(strText2Find As String)
Dim i As Long
Dim selSt As Long
Dim lStart As Long
Dim lMax As Long
With CodeMain
    lStart = 1
    lMax = Len(.Text)
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
End Select
End If
End Sub

Private Sub Form_Activate()
If Not IsLoaded Then
    CodeMain.Font = Me.Font
    IsLoaded = True
    dbgMainXML.OpenFromString myScriptOBJ.XMLString
    LoadAProject
    'GenerateNewProject "New Project"
    isDirty = False
    'RunSilent
End If
BringWindowToTop Me.hWnd
End Sub
Public Function SetScriptDLLObject(dllObj As ScriptEngine) As Boolean
    Set myScriptOBJ = dllObj
End Function
Private Sub SearchOn()
    If bRunningSilent Then Exit Sub
    SearchImg.Visible = True
    SearchResult.Caption = ""
    FindText(0).TabStop = True
    Command2(0).TabStop = True
    Command2(1).TabStop = True
    Check1(0).TabStop = True
    
    CodeMain.TabStop = False
    FindText(0).TabIndex = 1
    Command2(0).TabIndex = 2
    Command2(1).TabIndex = 4
    Check1(0).TabIndex = 6
    Form_Resize
    FindText(0).SetFocus
End Sub
Private Sub SearchOff()
    FindText(0).TabStop = False
    Command2(0).TabStop = False
    Command2(1).TabStop = False
    Check1(0).TabStop = False
    CodeMain.TabStop = True
    SearchImg.Visible = False
    Form_Resize
    If Not bRunningSilent Then
        CodeMain.SetFocus
    End If
End Sub

Public Function CountChildren(strParent As String) As Long
Dim ret As Long
Dim i As Long
Dim nd As Object
With dbgMainXML
    Set nd = .GetChildNode(.GetRootChildren(), strParent)
    If (nd Is Nothing) Then
        CountChildren = 0
        Exit Function
    End If
    ret = CLng("0" & .GetAttributeValue(nd, "COUNT"))
    CountChildren = ret
End With
End Function
Public Function UpdateProject(strXML As String) As Boolean
With dbgMainXML
    .OpenFromString strXML
    isDirty = True
    LoadAProject
    UpdateProject = True
End With
End Function
Public Function UpdateConstants(strXML As String) As Boolean
Dim ob1 As QSXML
Dim nd As Object
Dim ndR As Object
Dim ndl As Object
Set ob1 = New QSXML
ob1.Initialize pavAUTO
ob1.OpenFromString dbgMainXML.XML

With ob1
    Set ndl = .GetRootChildren()
    Set nd = .GetChildNode(ndl, "CONSTANTS")
    If .RemoveNode(nd) Then
        Set ndR = .GetRootElement
        Set nd = .XMLAddNode(ndR, strXML)
    End If
End With
isDirty = True
With dbgMainXML
    .OpenFromString ob1.XML
    Set ob1 = Nothing
    LoadAProject "CONSTANTS"
    UpdateConstants = True
    Exit Function
End With
Exit Function
ERRHDL:
MsgBox Err.Description
Err.Clear
    UpdateConstants = False
    Exit Function

End Function
Public Function UpdateVariables(strXML As String) As Boolean
Dim ob1 As QSXML
Dim nd As Object
Dim ndR As Object
Dim ndl As Object
Set ob1 = New QSXML
ob1.Initialize pavAUTO
ob1.OpenFromString dbgMainXML.XML

With ob1
    Set ndl = .GetRootChildren()
    Set nd = .GetChildNode(ndl, "VARIABLES")
    If .RemoveNode(nd) Then
        Set ndR = .GetRootElement
        Set nd = .XMLAddNode(ndR, strXML)
    End If
End With
With dbgMainXML
    .OpenFromString ob1.XML
    Set ob1 = Nothing
    isDirty = True
    LoadAProject "VARIABLES"
    UpdateVariables = True
    Exit Function
End With
Exit Function
ERRHDL:
MsgBox Err.Description
Err.Clear
    UpdateVariables = False
    Exit Function
End Function

Private Sub Form_Load()
Dim i As Long
Set SC1 = myScriptOBJ.MSVBScriptObject()
ClearFunctions
' Loading Color
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
Set dbgMainXML = New QSXML
dbgMainXML.Initialize pavAUTO

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
    Select Case MsgBox("Save changes?", vbQuestion + vbYesNoCancel, "Save project")
        Case vbCancel
            Cancel = True
            Exit Sub
        Case vbYes
            mnFileSaveProject_Click
    End Select
End If
End Sub

Private Sub Form_Resize()
SizeControls imgSplitter.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set dbgMainXML = Nothing
IsLoaded = False

End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    
    picSplitter.Visible = True
    picSplitter.ZOrder 0
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub
Private Sub SizeSilent()
Dim i As Long, j As Long
With SilentWindow
    sb1.ZOrder 0
    .Top = tbToolbar.Height
    .Height = (Me.ScaleHeight - sb1.Height) - .Top
    .Left = 0
    .Width = Me.ScaleWidth
    .ZOrder 0
    .Visible = True
    If .ScaleWidth <= ExeTree.Width + 50 Then
        Exit Sub
    End If
    j = .ScaleHeight
    i = .ScaleWidth
End With
With ExeTree
    .Left = 0
    .Top = 0
    .Height = SilentWindow.ScaleHeight
    exeLabel1(0).Left = .Width + 20
    exeLabel1(1).Left = .Width + 20
    exeList1.Left = .Width + 20
    EXEResult.Left = .Width + 20
    exeSilentLabel.Left = .Width + 20
    exeLabel1(0).Width = i - exeLabel1(0).Left
    exeLabel1(1).Width = i - exeLabel1(1).Left
    EXEResult.Width = i - EXEResult.Left
    If j > EXEResult.Top Then
        EXEResult.Height = j - (EXEResult.Top)
    End If
    exeList1.Width = i - exeList1.Left
    exeSilentLabel.Width = i - exeSilentLabel.Left
End With

End Sub
Private Sub SizeControls(x As Single)
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If bRunningSilent Then
        If SearchImg.Visible Then
            SearchImg.Visible = False
        End If
        SizeSilent
        Exit Sub
    End If
    SilentWindow.ZOrder 1
    SilentWindow.Visible = False
    'set the width
    If x < 1500 Then x = 1500
    If x > (Me.Width - 1500) Then x = Me.Width - 1500
    tvTreeView.Width = x
    imgSplitter.Left = x
    CodeMain.Left = x + 40
    CodeMain.Width = Me.ScaleWidth - (tvTreeView.Width + 140)
    SearchImg.Left = CodeMain.Left
    SearchImg.Width = CodeMain.Width
    CloseSearch.Top = 0
    CloseSearch.Left = SearchImg.ScaleWidth - CloseSearch.Width
    'set the top
  

    If tbToolbar.Visible Then
        tvTreeView.Top = tbToolbar.Height
    Else
        tvTreeView.Top = 0
    End If
    SearchImg.Top = tvTreeView.Top
  
    

    'set the height
    If sb1.Visible Then
        tvTreeView.Height = Me.ScaleHeight - (tbToolbar.Height + sb1.Height)
    Else
        tvTreeView.Height = Me.ScaleHeight - (tbToolbar.Height)
    End If
    If SearchImg.Visible Then
        CodeMain.Top = SearchImg.Top + SearchImg.Height
        CodeMain.Height = Me.ScaleHeight - (tbToolbar.Height + sb1.Height + SearchImg.Height)
    Else
        CodeMain.Top = tvTreeView.Top
        CodeMain.Height = tvTreeView.Height
    End If

    
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height

End Sub


Private Sub m_Events_SMScriptEvent(evtName As String, evtParms As String)
If evtParms = "" Then
myScriptOBJ.Echo "Script Event: " & evtName
Else
myScriptOBJ.Echo "Script Event: " & evtName & " Parms: " & evtParms
End If
End Sub

Private Sub mnAddClass_Click()
Dim buff$
Dim strXML As String
Dim nd As Object
Dim ndl As Object
Dim i As Long
i = CountChildren("CLASSES") + 1
buff$ = Trim$(InputBox("Enter new Class name below:", "New Class", "MyClass" & i))

If buff$ = "" Then
    Exit Sub
End If
If Not isValidObjName(buff$) Then
    Exit Sub
End If
If ItemExists(buff$) Then
    MsgBox "A public object called '" & buff$ & "' already exists.", vbCritical, "Error.."
    Exit Sub
End If
strXML = "<CLASS></CLASS>"
With dbgMainXML
    Set ndl = .GetRootChildren()
    Set nd = .GetChildNode(ndl, "CLASSES")
    Set ndl = .XMLAddNode(nd, strXML)
    .SetAttribute ndl, "NAME", buff$
    .SetAttribute ndl, "SCOPE", "Public"
    buff$ = BlankClassCode()
    ndl.Text = buff$
    Set ndl = .GetChildNodeList(nd)
    .SetAttribute nd, "COUNT", CStr(ndl.Length)
    isDirty = True
    LoadAProject "CLASSES"
End With
End Sub

Private Sub mnAddEditConstants_Click()
Dim nd As Object
Dim ndl As Object
With dbgMainXML
    Set ndl = .GetRootChildren()
    Set nd = .GetChildNode(ndl, "CONSTANTS")
    frmConstants.strXML = nd.XML
    frmConstants.Show 1
    
End With
End Sub

Private Sub mnAddFunction_Click()
                EditSubDefinition.InitClassType = 1
                EditSubDefinition.Show 1
                Exit Sub
End Sub

Private Sub mnAddSub_Click()
                EditSubDefinition.InitClassType = 0
                EditSubDefinition.Show 1
                Exit Sub

End Sub

Private Sub mnAddVars_Click()
Dim nd As Object
Dim ndl As Object
With dbgMainXML
    Set ndl = .GetRootChildren()
    Set nd = .GetChildNode(ndl, "VARIABLES")
    frmVariables.strXML = nd.XML
    frmVariables.Show 1
End With

End Sub

Private Sub mnDebugRun_Click()
        RunCode
        Exit Sub

End Sub

Private Sub mnDebugStop_Click()
On Error GoTo ERRDBSTOP
Dim i As Long

Set m_Events = Nothing
With tbToolbar
    For i = 1 To .Buttons.Count
        If .Buttons(i).Key <> "STOP" Then
            .Buttons(i).Enabled = True
        End If
    Next
    Me.mnFile.Enabled = True
    Me.mnProject.Enabled = True
    Me.mnDebugRun.Enabled = True
End With
        myScriptOBJ.ClearErrors
        myScriptOBJ.ResetProject
        bRunningSilent = False
        SizeControls imgSplitter.Left
        Exit Sub
ERRDBSTOP:
MsgBox Err.Description, vbInformation, "Error.."
Err.Clear
End Sub

Private Sub mnEditInitialization_Click()
Dim nd As Object
Set nd = GetItemNode("INITIALIZATION", True)
MainCodeEditor.MyObjXML = nd.XML
MainCodeEditor.Show 1


End Sub

Private Sub mnFileExit_Click()
Unload Me
End Sub
Private Sub GenerateNewProject(newPRJName As String)
Dim buff$
Dim nd As Object
Dim ndc As Object
Dim ndp As Object
ClearFunctions
buff$ = "<SCRIPTMAGIC></SCRIPTMAGIC>"
With dbgMainXML
    .OpenFromString buff$
    Set nd = .GetRootElement()
    .SetAttribute nd, "NAME", newPRJName
    .SetAttribute nd, "HOSTID", ""
    .SetAttribute nd, "RUNMODE", "IMMEDIATE"
    .SetAttribute nd, "PASSWORD", ""
    .SetAttribute nd, "TIMEOUT", "10"
    .SetAttribute nd, "EXPLICIT", "0"
    .SetAttribute nd, "CREATED", Format$(Now, "dd mmm yyyy hh:nn AMPM")
    .SetAttribute nd, "CREATEDBY", "SYSTEM"
    .SetAttribute nd, "LASTMODIFIED", ""
    .SetAttribute nd, "LASTMODIFIEDBY", ""
    Set ndc = .AddNode(nd, "", "DESCRIPTION")
       ndc.Text = "'Script Magic Project: " & newPRJName
    Set ndc = .AddNode(nd, "", "CONSTANTS")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "VARIABLES")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "INPUT")
    .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "INITIALIZATION")
        buff$ = "'Step 1 Call the Main Processing Sub Routine" & vbLf
        buff$ = buff$ & "Call Main() " & vbLf & vbLf
        ndc.Text = buff$
    Set ndc = .AddNode(nd, "", "SUBROUTINES")
        .SetAttribute ndc, "COUNT", "1"
    Set ndp = .AddNode(ndc, "", "SUBROUTINE")
        .SetAttribute ndp, "NAME", "Main"
        .SetAttribute ndp, "PARAMETERS", ""
        .SetAttribute ndp, "SCOPE", "Public"
        ndp.Text = "MsgBox " & Chr$(34) & "Hello World" & Chr$(34)
    Set ndc = .AddNode(nd, "", "FUNCTIONS")
        .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "CLASSES")
        .SetAttribute ndc, "COUNT", "0"
    Set ndc = .AddNode(nd, "", "VBSCRIPT")
End With
isDirty = True

LoadAProject
Exit Sub
End Sub
Public Sub LoadAProject(Optional strSelOpenKey As String = "")
LoadProjectTree strSelOpenKey
LoadProjectRTF
End Sub
Private Sub LoadProjectTree(Optional strSelOpenKey As String = "")
Dim nodR As Node
Dim nodP As Node
Dim nodC As Node
Dim nodG As Node
Dim nd As Object
Dim rootNDL As Object
Dim ndc As Object
Dim ndp As Object
Dim ndl As Object
Dim buff$, i As Long
SM_FUNCTIONWORDS = SM_FUNCTIONCONST
With dbgMainXML
    Set nd = .GetRootElement()
        buff$ = .GetAttributeValue(nd, "PASSWORD")
        
        If Len(buff$) > 0 Then
            bPasswordIsRequired = True
            strPassword = sm_DecodeText(buff$)
        Else
            bPasswordIsRequired = False
            strPassword = ""
        End If
        sb1.Panels(1).Text = "Project: " & .GetAttributeValue(nd, "NAME")
        buff$ = .GetAttributeValue(nd, "FILENAME")
        If buff$ <> "" Then
            sb1.Panels(2).Text = buff$
            sb1.Panels(2).ToolTipText = buff$
        Else
            sb1.Panels(2).Text = "New Project [Unsaved]"
            sb1.Panels(2).ToolTipText = "Unsaved project"
        End If
        buff$ = ""
    If .GetAttributeValue(nd, "NAME") = "" Then
        CodeMain.Locked = False
        CodeMain.Text = ""
        CodeMain_Change
        CodeMain.Locked = True
        Exit Sub
    End If
    tvTreeView.Nodes.Clear
    Set nodR = tvTreeView.Nodes.Add(, , "PROJECT", "Project: " & .GetAttributeValue(nd, "NAME"), "PROJECT", "PROJECT")
    nodR.Expanded = True
    nodR.Bold = True
    Set rootNDL = .GetChildNodeList(nd)
            
            
            Set ndc = .GetChildNode(rootNDL, "CONSTANTS")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "CONSTANTS", "Public Constant Declarations (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "CONSTANTS", "CONSTANTS")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "ITEM", "ITEM")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                Next
            End If
            
            Set ndc = .GetChildNode(rootNDL, "VARIABLES")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "VARIABLES", "Public Variables (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "VARIABLE", "VARIABLE")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "ITEM", "ITEM")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                Next
            End If
            
            
'            Set ndc = .GetChildNode(rootNDL, "INPUT")
'            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "INPUT", "Script Input Variables (" & _
'            .GetAttributeValue(ndc, "COUNT") & ")", "INPUT", "INPUT")
'            If nodP.Key = strSelOpenKey Then
'                nodP.Expanded = True
'            End If
'            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
'            nodC.Tag = "ADD-NEW-ITEM"
'            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
'                Set ndl = .GetChildNodeList(ndc)
'                For i = 0 To ndl.Length - 1
'                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "ITEM", "ITEM")
'                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
'                Next
'            End If
            
            
            Set ndc = .GetChildNode(rootNDL, "INITIALIZATION")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "INITIALIZATION", "Initialization (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "CODE", "CODE")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to edit}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            
            Set ndc = .GetChildNode(rootNDL, "SUBROUTINES")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "SUBROUTINES", "SubRoutines (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "SUBROUTINES", "SUBROUTINES")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    If UCase$(.GetAttributeValue(ndl(i), "SCOPE")) = "PUBLIC" Then
                        AddFunctionName .GetAttributeValue(ndl(i), "NAME"), .GetAttributeValue(ndl(i), "PARAMETERS")
                    End If
                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "SUBROUTINE", "SUBROUTINE")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                    SM_FUNCTIONWORDS = RTrim(SM_FUNCTIONWORDS) & " " & nodC.Tag & " "
                Next
            End If
            
            Set ndc = .GetChildNode(rootNDL, "FUNCTIONS")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "FUNCTIONS", "Functions (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "FUNCTIONS", "FUNCTIONS")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    If UCase$(.GetAttributeValue(ndl(i), "SCOPE")) = "PUBLIC" Then
                        AddFunctionName .GetAttributeValue(ndl(i), "NAME"), .GetAttributeValue(ndl(i), "PARAMETERS")
                    End If
                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "FUNCTION", "FUNCTION")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                    SM_FUNCTIONWORDS = RTrim(SM_FUNCTIONWORDS) & " " & nodC.Tag & " "
                Next
            End If
            
            Set ndc = .GetChildNode(rootNDL, "CLASSES")
            Set nodP = tvTreeView.Nodes.Add(nodR.Key, tvwChild, "CLASSES", "Classes (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "CLASS", "CLASS")
            If nodP.Key = strSelOpenKey Then
                nodP.Expanded = True
            End If
            Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , "{Double click to create}", "BUTTON", "BUTTON")
            nodC.Tag = "ADD-NEW-ITEM"
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    Set nodC = tvTreeView.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "CLASS", "CLASS")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                    SM_FUNCTIONWORDS = RTrim(SM_FUNCTIONWORDS) & " " & nodC.Tag & " "
                Next
            End If
End With

End Sub
Public Function IndentIt(strText As String, strIndent As Long) As String
Dim i As Long
Dim ret As String
ReDim ed1(0) As String
ed1 = Split(strText, vbLf)
For i = 0 To UBound(ed1)
    If ret = "" Then
        If Left$(ed1(i), 1) <> "'" Then
        ret = Space$(strIndent) & ed1(i)
        Else
        ret = ed1(i)
        End If
    Else
        If Left$(ed1(i), 1) <> "'" Then
        ret = ret & vbLf & Space$(strIndent) & ed1(i)
        Else
        ret = ret & vbLf & ed1(i)
        End If
    End If
Next
IndentIt = ret
End Function
Public Function LoadProjectRTF() As Boolean
Dim nd As Object
Dim rootNDL As Object
Dim ndc As Object
Dim ndp As Object
Dim ndl As Object
Dim buff$, i As Long, tmp1$
With dbgMainXML
    Set nd = .GetRootElement()
    If .GetAttributeValue(nd, "NAME") = "" Then
        CodeMain.Locked = False
        CodeMain.Text = ""
        CodeMain_Change
        CodeMain.Locked = True
        Exit Function
    End If
    buff$ = "'" & Chr$(171) & "Project: " & .GetAttributeValue(nd, "NAME") & Chr$(187) & vbLf
    buff$ = buff$ & "'" & String$(50, "*") & vbLf
    buff$ = buff$ & "'Created: " & .GetAttributeValue(nd, "CREATED") & _
    " Author: " & .GetAttributeValue(nd, "AUTHOR") & vbLf
    buff$ = buff$ & "'Script Magic Run Mode: " & .GetAttributeValue(nd, "RUNMODE")
    If .GetAttributeValue(nd, "PASSWORD") = "" Then
        buff$ = buff$ & " No Password" & vbLf
    Else
        buff$ = buff$ & " Password Protected" & vbLf
    End If
    tmp1$ = .GetAttributeValue(nd, "TIMEOUT")
    If tmp1$ = "" Then tmp1$ = "10"
    If tmp1$ = "0" Then
        buff$ = buff$ & "'Script Timeout: Unlimited" & vbLf
    Else
        buff$ = buff$ & "'Script Timeout: " & tmp1$ & " seconds" & vbLf
    End If
    buff$ = buff$ & "'Project Description:" & vbLf
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
            buff$ = buff$ & GetSectionHeader("Active Global Objects") & vbLf
            buff$ = buff$ & "'SCRIPT MAGIC OBJECTS: SMDebug, SMEvent" & vbLf
            If myScriptOBJ.HasParentForm() Then
            buff$ = buff$ & "'OBJECT: ParentForm (should be a reference to the currently active window)" & vbLf
            Else
            buff$ = buff$ & "'WARNING: NO Parent form object set by calling application" & vbLf
            End If
            If myScriptOBJ.ScriptUserObjectCount = 0 Then
                buff$ = buff$ & "'USER OBJECTS: <None>" & vbLf
            Else
                buff$ = buff$ & "'USER OBJECTS: "
                For i = 1 To myScriptOBJ.ScriptUserObjectCount
                    buff$ = buff$ & myScriptOBJ.ScriptUserObjectName(i)
                Next
                buff$ = buff$ & vbLf
            End If
            buff$ = buff$ & vbLf
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
                    IndentIt(ndl(i).Text, 5) & vbLf & _
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
                    buff$ = buff$ & IndentIt(ndl(i).Text, 5) & vbLf & "End Function" & vbLf
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
                    buff$ = buff$ & IndentIt(ndl(i).Text, 5) & vbLf & "End Class" & vbLf
                    buff$ = buff$ & GetClassFooter(.GetAttributeValue(ndl(i), "NAME")) & vbLf
                Next
            End If
            buff$ = buff$ & GetSectionHeader("END OF FILE")
End With
CodeMain.Locked = False
CodeMain.Text = ""
CodeMain.Text = buff$
CodeMain.SelStart = Len(CodeMain.Text)
CodeMain_Change
CodeMain.Locked = True
CodeMain.SelStart = 0
CodeMain.SelLength = Len(CodeMain.Text) + 1
CodeMain.SelFontName = "Courier New"
CodeMain.SelFontSize = "10"
CodeMain.SelLength = 0
DoEvents
End Function

Private Sub mnFileExpFull_Click()
Dim buff$
Dim strVBS As String
On Error Resume Next
If bPasswordIsRequired Then
    If Not PromptForPassword(strPassword) Then
        Exit Sub
    End If
End If
If myOpenPath <> "" Then
CommonDialog1.InitDir = myOpenPath
End If
CommonDialog1.CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
CommonDialog1.DialogTitle = "Export to VB Script File"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "VB Script Files|*.vbs"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.ShowSave
If Err = cdlCancel Then ' 'Cancel' button was hit
    Err.Clear
    Exit Sub
End If
On Error GoTo ERRHDL
myOpenPath = GetPathFromFileName(CommonDialog1.FileName)
buff$ = CommonDialog1.FileName
myScriptOBJ.OpenProject dbgMainXML.XML, OpenString
strVBS = myScriptOBJ.GenerateProjectVBScript(False)
ScratchRTF.Text = strVBS
If Dir$(buff$) <> "" Then Kill buff$
ScratchRTF.SaveFile buff$, rtfText
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear

End Sub

Private Sub mnFileExport_Click()
Dim buff$
Dim strVBS As String
On Error Resume Next
If bPasswordIsRequired Then
    If Not PromptForPassword(strPassword) Then
        Exit Sub
    End If
End If
If myOpenPath <> "" Then
CommonDialog1.InitDir = myOpenPath
End If
CommonDialog1.CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
CommonDialog1.DialogTitle = "Export to VB Script File"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "VB Script Files|*.vbs"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.ShowSave
If Err = cdlCancel Then ' 'Cancel' button was hit
    Err.Clear
    Exit Sub
End If
On Error GoTo ERRHDL
myOpenPath = GetPathFromFileName(CommonDialog1.FileName)
buff$ = CommonDialog1.FileName
myScriptOBJ.OpenProject dbgMainXML.XML, OpenString
strVBS = myScriptOBJ.GenerateProjectVBScript(True)
ScratchRTF.Text = strVBS
If Dir$(buff$) <> "" Then Kill buff$
ScratchRTF.SaveFile buff$, rtfText
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub mnFileNew_Click()
Dim buff$
        buff$ = Trim$(InputBox("New Project Name", "New Project"))
        If buff$ = "" Then Exit Sub
        GenerateNewProject buff$
End Sub

Private Sub mnFileOpenProject_Click()
Dim tmpXML As QSXML
Dim buff$
Dim nd As Object
On Error Resume Next
If isDirty Then
    If MsgBox("You have unsaved changes are you sure you want to continue", vbYesNo + vbQuestion, "Open and lose changes") = vbNo Then
        Exit Sub
    End If
End If
If myOpenPath <> "" Then
CommonDialog1.InitDir = myOpenPath
End If
CommonDialog1.CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
CommonDialog1.DialogTitle = "Open a ScriptMagic Project"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Script Magic Files|*.mgk"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
CommonDialog1.ShowOpen
If Err = cdlCancel Then ' 'Cancel' button was hit
    Err.Clear
    Exit Sub
End If
On Error GoTo ERRHDL
myOpenPath = GetPathFromFileName(CommonDialog1.FileName)
buff$ = OpenScriptMagicFile(CommonDialog1.FileName, "", True)
If buff$ = "" Then
    Exit Sub
End If
Set tmpXML = New QSXML
With tmpXML
    .Initialize pavAUTO
    If Not .OpenFromString(buff$) Then
        Set tmpXML = Nothing
        Exit Sub
    End If
    Set nd = .GetRootElement()
    .SetAttribute nd, "FILENAME", CommonDialog1.FileName
    buff$ = .XML
End With
Set tmpXML = Nothing
With dbgMainXML
    If .OpenFromString(buff$, True) Then
        ClearFunctions
        LoadAProject
        isDirty = False
    End If
End With
isDirty = True
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear
Exit Sub
End Sub

Private Sub mnFileRename_Click()
Dim buff$
Dim nd As Object
With dbgMainXML
    Set nd = .GetRootElement
    buff$ = .GetAttributeValue(nd, "NAME")
    buff$ = Trim$(InputBox("New Project Name: ", "Rename Project", buff$))
    If buff$ = "" Then
        Exit Sub
    End If
    If Len(buff$) > 50 Then
        MsgBox "Project names are a maximum of 50 characters."
        Exit Sub
    End If
    .SetAttribute nd, "NAME", buff$
    LoadAProject
End With
End Sub

Private Sub mnFileSaveAs_Click()
Dim buff$, nd As Object
Dim nd2 As Object
On Error Resume Next
If bPasswordIsRequired Then
    If Not PromptForPassword(strPassword) Then
        Exit Sub
    End If
End If
With CommonDialog1
    If .InitDir = "" Then .InitDir = App.Path

CommonDialog1.CancelError = True ' Causes a trappable error to occur when the user hits the 'Cancel' button
CommonDialog1.DialogTitle = "Save ScriptMagic Project"
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Script Magic Files|*.mgk"
CommonDialog1.FilterIndex = 1
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.ShowSave

If Err = cdlCancel Then ' 'Cancel' button was hit
    Err.Clear
    Exit Sub
End If
On Error GoTo ERRHDL
myOpenPath = GetPathFromFileName(CommonDialog1.FileName)
buff$ = .FileName
End With
If Dir$(buff$) <> "" Then Kill buff$
With dbgMainXML
    Set nd = .GetRootElement()
    
    If Not .IsChildNode(nd, "VBSCRIPT") Then
        Set nd2 = .XMLAddNode(nd, "<VBSCRIPT></VBSCRIPT>")
    Else
        Set nd2 = .GetChildNode(.GetRootChildren(), "VBSCRIPT")
    End If
    myScriptOBJ.OpenProject dbgMainXML.XML, OpenString
    nd2.Text = myScriptOBJ.GenerateProjectVBScript(True)
    .SetAttribute nd, "FILENAME", ""
    .Save buff$
    .SetAttribute nd, "FILENAME", buff$
'    isDirty = False
    LoadAProject
        sb1.Panels(1).Text = "Project: " & .GetAttributeValue(nd, "NAME")
        buff$ = .GetAttributeValue(nd, "FILENAME")
        If buff$ <> "" Then
            sb1.Panels(2).Text = buff$
            sb1.Panels(2).ToolTipText = buff$
        Else
            sb1.Panels(2).Text = "New Project [Unsaved]"
            sb1.Panels(2).ToolTipText = "Unsaved project"
        End If
End With

Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear


End Sub

Private Sub mnFileSaveProject_Click()
Dim buff$, nd As Object
Dim nd2 As Object
On Error GoTo ERRHDL
If bPasswordIsRequired Then
    If Not PromptForPassword(strPassword) Then
        Exit Sub
    End If
End If
With dbgMainXML
    Set nd = .GetRootElement()
     If Not .IsChildNode(nd, "VBSCRIPT") Then
        Set nd2 = .XMLAddNode(nd, "<VBSCRIPT></VBSCRIPT>")
    Else
        Set nd2 = .GetChildNode(.GetRootChildren(), "VBSCRIPT")
    End If
    buff$ = .GetAttributeValue(nd, "FILENAME")
    .SetAttribute nd, "FILENAME", ""
    myScriptOBJ.OpenProject dbgMainXML.XML, OpenString
    nd2.Text = myScriptOBJ.GenerateProjectVBScript(True)
    myScriptOBJ.SaveProject
    isDirty = False
    Exit Sub
End With
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear

End Sub
Private Sub LoadSilTree()
Dim nodR As Node
Dim nodP As Node
Dim nodC As Node
Dim nodG As Node
Dim nd As Object
Dim rootNDL As Object
Dim ndc As Object
Dim ndp As Object
Dim ndl As Object
Dim buff$, i As Long
With dbgMainXML
    Set nd = .GetRootElement()
        buff$ = ""
    ExeTree.Nodes.Clear
    exeList1.Clear
    Set nodR = ExeTree.Nodes.Add(, , "PROJECT", "Project: " & .GetAttributeValue(nd, "NAME"), "PROJECT", "PROJECT")
    nodR.Expanded = True
    nodR.Bold = True
    Set rootNDL = .GetChildNodeList(nd)
            Set ndc = .GetChildNode(rootNDL, "CONSTANTS")
            Set nodP = ExeTree.Nodes.Add(nodR.Key, tvwChild, "CONSTANTS", "Public Constant Declarations (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "CONSTANTS", "CONSTANTS")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    Set nodC = ExeTree.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "ITEM", "ITEM")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                Next
            End If
            If nodP.Children = 0 Then
                ExeTree.Nodes.Remove (nodP.Index)
            End If
            Set ndc = .GetChildNode(rootNDL, "VARIABLES")
            Set nodP = ExeTree.Nodes.Add(nodR.Key, tvwChild, "VARIABLES", "Public Variables (" & _
            .GetAttributeValue(ndc, "COUNT") & ")", "VARIABLE", "VARIABLE")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    Set nodC = ExeTree.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "ITEM", "ITEM")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                Next
            End If
            If nodP.Children = 0 Then
                ExeTree.Nodes.Remove (nodP.Index)
            End If
            Set ndc = .GetChildNode(rootNDL, "SUBROUTINES")
            Set nodP = ExeTree.Nodes.Add(nodR.Key, tvwChild, "SUBROUTINES", "SubRoutines", "SUBROUTINES", "SUBROUTINES")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                    If UCase$(.GetAttributeValue(ndl(i), "SCOPE")) = "PUBLIC" Then
                        Set nodC = ExeTree.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "SUBROUTINE", "SUBROUTINE")
                        nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                    End If
                Next
            End If
            If nodP.Children = 0 Then
                ExeTree.Nodes.Remove (nodP.Index)
            End If
            Set ndc = .GetChildNode(rootNDL, "FUNCTIONS")
            Set nodP = ExeTree.Nodes.Add(nodR.Key, tvwChild, "FUNCTIONS", "Functions", "FUNCTIONS", "FUNCTIONS")
            If .GetAttributeValue(ndc, "COUNT") <> "0" Then
                Set ndl = .GetChildNodeList(ndc)
                For i = 0 To ndl.Length - 1
                  If UCase$(.GetAttributeValue(ndl(i), "SCOPE")) = "PUBLIC" Then
                    Set nodC = ExeTree.Nodes.Add(nodP.Key, tvwChild, , .GetAttributeValue(ndl(i), "NAME"), "FUNCTION", "FUNCTION")
                    nodC.Tag = .GetAttributeValue(ndl(i), "NAME")
                  End If
                Next
            End If
            If nodP.Children = 0 Then
                ExeTree.Nodes.Remove (nodP.Index)
            End If
End With
buff$ = "Your code is loaded into memory.  You can double click on any " & _
"of the public objects to the left and see the result in the immediate window." & _
"  View all public variables below."
exeSilentLabel.Caption = buff$

With tbToolbar
    For i = 1 To .Buttons.Count
        If .Buttons(i).Key <> "STOP" Then
            .Buttons(i).Enabled = False
        End If
    Next
    Me.mnFile.Enabled = False
    Me.mnProject.Enabled = False
    Me.mnDebugRun.Enabled = False
End With


End Sub
Private Sub EvalVariables()
Dim nd As Object
Dim ndl As Object
Dim rootl As Object
Dim ndc As Object
Dim obName As String
Dim buff$
Dim i As Long
List1.Clear
On Error Resume Next
With dbgMainXML
    Set rootl = .GetRootChildren
    Set nd = .GetChildNode(rootl, "VARIABLES")
    i = CLng(.GetAttributeValue(nd, "COUNT"))
    If i > 0 Then
        Set ndl = .GetChildNodeList(nd)
        For i = 0 To ndl.Length - 1
            obName = .GetAttributeValue(ndl(i), "NAME")
            buff$ = "Variable: " & obName & " = "
            buff$ = buff$ & CStr(SC1.Eval(obName) & "")
            If MyError.errNumber <> 0 Then
                buff$ = buff$ & "<ERROR>"
                MyError.errNumber = 0
            End If
            List1.AddItem buff$
        Next
    End If
    exeList1.Clear
    For i = 0 To List1.ListCount - 1
        exeList1.AddItem List1.List(i)
    Next
    List1.Clear
    Set nd = .GetChildNode(rootl, "CONSTANTS")
    i = CLng(.GetAttributeValue(nd, "COUNT"))
    If i > 0 Then
        Set ndl = .GetChildNodeList(nd)
        For i = 0 To ndl.Length - 1
            obName = .GetAttributeValue(ndl(i), "NAME")
            buff$ = "Constant: " & obName & " = "
            buff$ = buff$ & CStr(SC1.Eval(obName) & "")
            If MyError.errNumber <> 0 Then
                buff$ = buff$ & "<ERROR>"
                MyError.errNumber = 0
            End If
            List1.AddItem buff$
        Next
    End If
    For i = 0 To List1.ListCount - 1
        exeList1.AddItem List1.List(i)
    Next
    List1.Clear
End With
Err.Clear
Exit Sub
End Sub
Private Sub RunSilent()
Dim nd As Object
Dim buff$

On Error GoTo ERRHDL

MyError.errNumber = 0
LoadSilTree
bRunningSilent = True
SizeControls imgSplitter.Left
myScriptOBJ.OpenProject dbgMainXML.XML, OpenString
myScriptOBJ.ExecuteScriptDebug
DoEvents
EvalVariables
Exit Sub
ERRHDL:
SC1_Error
If MyError.errNumber = 0 Then
MsgBox Err.Description
Err.Clear
Else
    ProcessMyError False
End If

End Sub
Private Sub RunCode()
Dim nd As Object
Dim buff$
On Error GoTo ERRHDL

Set m_Events = New SM_Event
With dbgMainXML
Set nd = .GetRootElement
    buff$ = .GetAttributeValue(nd, "PASSWORD")
    If buff$ <> "" Then
        buff$ = sm_DecodeText(buff$)
        If Not PromptForPassword(buff$) Then
            Exit Sub
        End If
    End If
    RunSilent
End With
Exit Sub
ERRHDL:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub ProcessMyError(Optional bPassthrough As Boolean = True)
If bPassthrough Then
    MyError.errNumber = 0
    If MyError.errLine <> 0 Then
        MsgBox "Error at line: " & MyError.errLine & vbLf & MyError.errDesc
        If bRunningSilent Then
            mnDebugStop_Click
        End If
        DoEvents
        CodeMain.SetFocus
        GoToLineNumber CodeMain, MyError.errLine
        DoEvents
        GoToLineNumber CodeMain, MyError.errLine
        DoEvents
        SelectLine CodeMain
        DoEvents
        SelectLine CodeMain
    Else
        MsgBox "Script Error <Line Unknown> " & vbLf & MyError.errDesc
        If bRunningSilent Then
            mnDebugStop_Click
        End If
    End If
    Exit Sub
Else
    MyError.errNumber = 0
    If MyError.errLine <> 0 Then
        If MsgBox("Error at line: " & MyError.errLine & vbLf & MyError.errDesc & vbLf & vbLf & "Stop running and view Code?", vbYesNo + vbQuestion, "Show Error in Code") = vbNo Then
            Exit Sub
        End If
        If bRunningSilent Then
            mnDebugStop_Click
        End If
        DoEvents
        CodeMain.SetFocus
        GoToLineNumber CodeMain, MyError.errLine
        DoEvents
        GoToLineNumber CodeMain, MyError.errLine
        DoEvents
        SelectLine CodeMain
        DoEvents
        SelectLine CodeMain
    Else
        If MsgBox("Script Error <Unknown Line Number>" & vbLf & MyError.errDesc & vbLf & vbLf & "Stop running and view Code?", vbYesNo + vbQuestion, "Stop Script and view Code") = vbNo Then
            Exit Sub
        End If
        If bRunningSilent Then
            mnDebugStop_Click
        End If
    End If
    Exit Sub
End If

End Sub
Private Sub mnFileTest_Click()
frmProperties.MyPrjXML = dbgMainXML.XML
frmProperties.Show 1
End Sub


Private Sub mnFileURL_Click()
    ObjectCodeEditor.Show 1
End Sub

Private Sub mnMSDN_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://msdn.microsoft.com/scripting") Then
    MsgBox "Error opening: http://msdn.microsoft.com/scripting"
End If
End Sub

Private Sub mnMSExamples_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.microsoft.com/technet/scriptcenter/default.mspx") Then
    MsgBox "Error opening: http://www.microsoft.com/technet/scriptcenter/default.mspx"
End If
End Sub

Private Sub mnProjectFind_Click()
If Not bRunningSilent Then
    SearchOn
    FindText(0).SetFocus
End If
End Sub

Private Sub mnViewParms_Click()
frmParameters.Show 1
End Sub

Private Sub mnWeb1_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.winguides.com/scripting/") Then
    MsgBox "Error opening: http://www.winguides.com/scripting/"
End If

End Sub

Private Sub mnWeb2_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.hostinglogin.com/help/vbscript") Then
    MsgBox "Error opening: http://www.hostinglogin.com/help/vbscript"
End If

End Sub

Private Sub mnWeb3_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www-sbras.nsc.ru/docs/ms/vbsdoc/vbstutor.htm") Then
    MsgBox "Error opening: "
End If

End Sub

Private Sub mnWeb4_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.functionx.com/vbscript/") Then
    MsgBox "Error opening: http://www.functionx.com/vbscript/"
End If

End Sub

Private Sub mnWeb5_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.w3schools.com/vbscript/default.asp") Then
    MsgBox "Error opening: http://www.w3schools.com/vbscript/default.asp"
End If

End Sub

Private Sub mnWeb6_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.visualbasicscript.com") Then
    MsgBox "Error opening: http://www.visualbasicscript.com"
End If

End Sub

Private Sub mnWeb7_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.devguru.com/Technologies/vbscript/quickref/vbscript_intro.html") Then
    MsgBox "Error opening: http://www.devguru.com/Technologies/vbscript/quickref/vbscript_intro.html"
End If

End Sub

Private Sub mnWind35_Click()
If Not OpenAnyFileURL(Me.hWnd, "http://www.winscripter.com/WSH/default.aspx") Then
    MsgBox "Error opening: http://www.winscripter.com/WSH/default.aspx"
End If

End Sub

Private Sub SC1_Error()
Dim i As Long
On Error GoTo ERRHDL
With SC1.Error
    MyError.errNumber = .Number
    MyError.errLine = .Line
    MyError.errDesc = .Description
    .Clear
End With
Exit Sub
ERRHDL:
MyError.errNumber = Err.Number
MyError.errDesc = Err.Description
MyError.errLine = 0
Err.Clear
'SC1.Reset
End Sub

Private Sub SC1_Timeout()
MsgBox "Script has timed out"
mnDebugStop_Click

End Sub

Private Sub myScriptOBJ_SMError(smErr As SMLastErrorType)
Dim i As Long
On Error GoTo ERRHDL
With smErr
    MyError.errNumber = .ErrorNumber
    MyError.errLine = .ScriptLineNumber
    MyError.errDesc = .ErrorDescription
    MsgBox "DEBUGGER: Error: " & .ErrorNumber & " at line: " & .ScriptLineNumber & " - " & .ErrorDescription
End With
Exit Sub
ERRHDL:
MyError.errNumber = Err.Number
MyError.errDesc = Err.Description
MyError.errLine = 0
Err.Clear
End Sub

Private Sub tbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim buff$

Select Case Button.Key
    Case "FIND"
        If bRunningSilent Then Exit Sub
        If SearchImg.Visible Then
            SearchOff
        Else
            SearchOn
        End If
        Exit Sub
    Case "STOP"
       mnDebugStop_Click
        Exit Sub
    Case "PROPERTIES"
        mnFileTest_Click
        Exit Sub
    Case "SCRIPT"
        CodeMain.Locked = False
        Exit Sub
    Case "RUN"
        RunCode
        Exit Sub
    Case "EXIT"
        Unload Me
        Exit Sub
    Case "SAVE"
        mnFileSaveProject_Click
        Exit Sub
    Case "OPEN"
        mnFileOpenProject_Click
        Exit Sub
    Case "NEW"
        buff$ = Trim$(InputBox("New Project Name", "New Project"))
        If buff$ = "" Then Exit Sub
        GenerateNewProject buff$
    End Select

End Sub

Private Sub tvTreeView_DblClick()
Dim nd As Object
With tvTreeView
    If (.SelectedItem Is Nothing) Then Exit Sub
    If .SelectedItem.Tag = "" Then Exit Sub
    If .SelectedItem.Tag = "ADD-NEW-ITEM" Then
        Select Case .SelectedItem.Parent.Key
            Case "CLASSES"
                mnAddClass_Click
                Exit Sub
            Case "INITIALIZATION"
                mnEditInitialization_Click
                Exit Sub
            Case "VARIABLES"
                mnAddVars_Click
                Exit Sub
            Case "SUBROUTINES"
                mnAddSub_Click
                Exit Sub
            Case "FUNCTIONS"
                mnAddFunction_Click
                Exit Sub
            Case "CONSTANTS"
                mnAddEditConstants_Click
                Exit Sub
            Case Else
                MsgBox "Add " & .SelectedItem.Parent.Key
                Exit Sub
        End Select
    End If
    Select Case .SelectedItem.Parent.Key
        Case "CONSTANTS"
                mnAddEditConstants_Click
                Exit Sub
        Case "VARIABLES"
                mnAddVars_Click
                Exit Sub
    End Select
    Set nd = GetItemNode(.SelectedItem.Tag)
    If (nd Is Nothing) Then
        Exit Sub
    End If
    EditObject nd
End With
End Sub
Public Function DeleteItemNode(strNodeName As String, Optional bIsRoot As Boolean = False) As Boolean
Dim nd As Object
Dim ndp As Object
Dim ret As Boolean
Dim buff$
Dim i As Long

Set nd = GetItemNode(strNodeName, bIsRoot)
If (nd Is Nothing) Then
    DeleteItemNode = False
    Exit Function
End If

With dbgMainXML
    RemoveFunction .GetAttributeValue(nd, "NAME")
    If Not bIsRoot Then
        Set ndp = nd.ParentNode
        buff$ = ndp.nodename
        i = ndp.childNodes.Length
        ret = .RemoveNode(nd)
        If ret And i > 0 Then
            .SetAttribute ndp, "COUNT", CStr(i - 1)
        End If
    Else
        ret = .RemoveNode(nd)
    End If
End With
If ret Then LoadAProject buff$
DeleteItemNode = ret
Exit Function
End Function

Public Function GetItemNode(itmName As String, Optional isRootItem As Boolean = False) As Object
Dim nd As Object
Dim ndl As Object
Dim ret As Object
Dim i As Long, j As Long
Dim rootl As Object
With dbgMainXML
Set rootl = .GetRootChildren()
If itmName = "" Then
    MsgBox "Invalid item name", vbCritical, "GetItemNode()"
    Set GetItemNode = ret
    Exit Function
End If
If isRootItem Then
    For i = 0 To rootl.Length - 1
        If UCase$(rootl(i).nodename) = UCase$(itmName) Then
            Set GetItemNode = rootl(i)
            Exit Function
        End If
    Next
    Set GetItemNode = Nothing
    Exit Function
End If
For i = 0 To rootl.Length - 1
    If CLng("0" & .GetAttributeValue(rootl(i), "COUNT")) > 0 Then
        Set ndl = .GetChildNodeList(rootl(i))
        For j = 0 To ndl.Length - 1
            If UCase$(.GetAttributeValue(ndl(j), "NAME")) = UCase$(itmName) Then
            Set GetItemNode = ndl(j)
            Exit Function
            End If
        Next
    End If
Next
End With
    Set GetItemNode = ret
    Exit Function
End Function
Public Function ItemExists(itmName As String) As Boolean
Dim nd As Object
Dim ndl As Object
Dim i As Long, j As Long
Dim rootl As Object
If InStr(SM_BUILTINOBJECTS, " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is the name of a built in ScriptMagic object", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
If InStr(UCase$(SM_RESERVEDWORDS), " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is a vbscript reserved word", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
If InStr(UCase$(SM_FUNCTIONCONST), " " & UCase$(itmName) & " ") > 0 Then
    MsgBox itmName & " is the name of a vbscript function", vbCritical, "Error.."
    ItemExists = True
    Exit Function
End If
With dbgMainXML
Set rootl = .GetRootChildren()
If itmName = "" Then
    MsgBox "Invalid item name", vbCritical, "ItemExists()"
    ItemExists = True
    Exit Function
End If
For i = 0 To rootl.Length - 1
    If CLng("0" & .GetAttributeValue(rootl(i), "COUNT")) > 0 Then
        Set ndl = .GetChildNodeList(rootl(i))
        For j = 0 To ndl.Length - 1
            If UCase$(.GetAttributeValue(ndl(j), "NAME")) = UCase$(itmName) Then
            ItemExists = True
            Exit Function
            End If
        Next
    End If
Next
End With
ItemExists = False
Exit Function
'Set nd = .GetChildNode(ndl, clsName)
'If (nd Is Nothing) Then
'    MsgBox clsName & " is invalid"
'    ItemExists = True
'    Exit Function
'End If
'If .GetAttributeValue(nd, "COUNT") = "0" Then
'    ItemExists = True
'    Exit Function
'End If
'    Set ndl = .GetChildNodeList(nd)
'    For i = 0 To ndl.length - 1
'        If UCase$(.GetAttributeValue(ndl(i), "NAME")) = UCase$(itmName) Then
'            ItemExists = True
'            Exit Function
'        End If
'    Next
'    ItemExists = False
'    Exit Function
'End With
End Function
Public Function AddProjectItem(clsParent As String, itmXML As String) As Boolean
Dim nd As Object
Dim ndl As Object
Dim i As Long
With dbgMainXML
Set ndl = .GetRootChildren()
Set nd = .GetChildNode(ndl, clsParent)
If (nd Is Nothing) Then
    MsgBox clsParent & " is invalid"
    AddProjectItem = False
    Exit Function
End If
.XMLAddNode nd, itmXML
Set ndl = .GetChildNodeList(nd)
.SetAttribute nd, "COUNT", CStr(ndl.Length)
End With
AddProjectItem = True
isDirty = True
LoadAProject clsParent
End Function
Public Function UpdateItem(itmName As String, strXML As String) As Boolean
Dim nd As Object
Dim ndp As Object
Dim xTmp As New QSXML
Dim xND As Object
xTmp.Initialize pavAUTO
xTmp.OpenFromString strXML
Set xND = xTmp.GetRootElement
With dbgMainXML
    Set nd = Me.GetItemNode(itmName)
    If (nd Is Nothing) Then
        UpdateItem = False
        Exit Function
    End If
    Set ndp = nd.ParentNode
    nd.Text = xND.Text
    .SetAttribute nd, "PARAMETERS", xTmp.GetAttributeValue(xND, "PARAMETERS")
    .SetAttribute nd, "NAME", xTmp.GetAttributeValue(xND, "NAME")
    .SetAttribute nd, "SCOPE", xTmp.GetAttributeValue(xND, "SCOPE")
End With
Set xTmp = Nothing
isDirty = True
LoadAProject ndp.nodename
UpdateItem = True
End Function
Public Function UpdateInitialization(itmName As String, strXML As String) As Boolean
Dim nd As Object
Dim xTmp As New QSXML
Dim xND As Object
xTmp.Initialize pavAUTO
xTmp.OpenFromString strXML
Set xND = xTmp.GetRootElement

With dbgMainXML
    Set nd = GetItemNode(itmName, True)
    If (nd Is Nothing) Then
        UpdateInitialization = False
        Exit Function
    End If
    nd.Text = xND.Text
End With
isDirty = True
Set xTmp = Nothing
LoadAProject "INITIALIZATION"
UpdateInitialization = True
End Function
Private Sub SelFindArea(strArea As String)
Dim stStart$
Dim stEnd$
Dim buff$, i As Long
stStart$ = "'" & Chr$(171)
stEnd$ = Chr$(187)
Select Case strArea
    Case "VARIABLES"
        buff$ = stStart$ & " Public Variables " & stEnd$
    Case "FUNCTIONS"
        buff$ = stStart$ & " ALL Functions " & stEnd$
    Case "SUBROUTINES"
        buff$ = stStart$ & " ALL SubRoutines " & stEnd$
    Case "CONSTANTS"
        buff$ = stStart$ & " Public Constants " & stEnd$
    Case "INITIALIZATION"
        buff$ = stStart$ & " Initialization Code " & stEnd$
    Case "INPUT"
        buff$ = stStart$ & " Script Input Variables " & stEnd$
    Case "CLASSES"
        buff$ = stStart$ & " ALL Classes " & stEnd$
    Case Else
        MsgBox strArea
        Exit Sub
End Select

With CodeMain
i = InStr(.Text, buff$)
If i > 0 Then
    .SelStart = i + 5
    DoEvents
    SelectLine CodeMain
    DoEvents
    Exit Sub
End If
End With
End Sub
Private Sub SelFindItem(strHeader As String, strArea As String)
Dim stStart$
Dim stEnd$
Dim buff$, i As Long
stStart$ = "'" & Chr$(171)
stEnd$ = Chr$(187)
Select Case strHeader
    Case "VARIABLES"
        buff$ = "Public " & strArea
    Case "FUNCTIONS"
        buff$ = " Function " & strArea & "("
    Case "SUBROUTINES"
        buff$ = " Sub " & strArea & "("
    Case "CONSTANTS"
        buff$ = "Public Const " & strArea & " = "
    Case "INITIALIZATION"
        buff$ = stStart$ & " Initialization Code " & stEnd$
    Case "INPUT"
        buff$ = stStart$ & " Script Input Variables " & stEnd$
    Case "CLASSES"
        buff$ = "Class " & strArea
    Case Else
        MsgBox strArea
        Exit Sub
End Select

With CodeMain
i = InStr(.Text, buff$)
If i = 0 Then Exit Sub
    .SelStart = i + 5
    DoEvents
    SelectLine CodeMain
    DoEvents
    Exit Sub
End With

End Sub
Private Sub tvTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim nd As Object
With tvTreeView
    
    If (.SelectedItem Is Nothing) Then Exit Sub
    If .SelectedItem.Index = 1 Then
        CodeMain.SelStart = 0
        Exit Sub
    End If
    If .SelectedItem.Tag = "" And .SelectedItem.Key = "" Then Exit Sub
    If .SelectedItem.Tag = "ADD-NEW-ITEM" Then Exit Sub
    Select Case .SelectedItem.Key
        Case "CONSTANTS"
                SelFindArea .SelectedItem.Key
                Exit Sub
        Case "VARIABLES"
                SelFindArea .SelectedItem.Key
                Exit Sub
        Case ""
            'donothing
        Case Else
                SelFindArea .SelectedItem.Key
                Exit Sub
    End Select
    If .SelectedItem.Tag = "" Then Exit Sub
    Set nd = GetItemNode(.SelectedItem.Tag)
    If (nd Is Nothing) Then
        Exit Sub
    End If
    SelFindItem .SelectedItem.Parent.Key, .SelectedItem.Tag
    Set nd = Nothing
End With

End Sub
Public Function OpenScriptMagicFile(strFilename As String, Optional strPassword As String = "", Optional verbose As Boolean = True) As String
Dim x As QSXML
Dim y As Object
Dim buff$
On Error GoTo ERRHDL
Set x = New QSXML
x.Initialize pavAUTO
If Not x.OpenFromFile(strFilename, verbose) Then
    OpenScriptMagicFile = ""
    Set x = Nothing
    Exit Function
End If
With x
    Set y = .GetRootElement()
    If UCase$(y.nodename) <> "SCRIPTMAGIC" Then
        If verbose Then
            MsgBox "Invalid file format.", vbCritical, "Error.."
        End If
        Set x = Nothing
        OpenScriptMagicFile = ""
        Exit Function
    End If
        buff$ = .GetAttributeValue(y, "PASSWORD")
        If buff$ <> "" Then
            buff$ = sm_DecodeText(buff$)
            If strPassword <> "" Then
                If UCase$(buff$) <> UCase$(strPassword) Then
                    If verbose Then
                        MsgBox "Invalid password for this project"
                    End If
                End If
            Else
                If Not PromptForPassword(buff$) Then
                    Set x = Nothing
                    OpenScriptMagicFile = ""
                    Exit Function
                End If
            End If
        End If
        OpenScriptMagicFile = .XML
End With
Set x = Nothing
Exit Function
ERRHDL:
If verbose Then
    MsgBox Err.Description, vbCritical, "OpenScriptMagicFile"
End If
    Err.Clear
OpenScriptMagicFile = ""
End Function
