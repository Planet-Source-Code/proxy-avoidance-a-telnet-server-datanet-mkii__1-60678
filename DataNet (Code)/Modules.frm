VERSION 5.00
Begin VB.Form Modules 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modules"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Add Module"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Standard.FS"
      Top             =   240
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Select the plugins that you would like to be open to users that use this domain."
      Top             =   1440
      Width           =   2775
   End
End
Attribute VB_Name = "Modules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Amount As Integer
Dim APPdir As String

Private Sub Command1_Click()
CreatePath (APPdir & "data")
fWriteValue APPdir & "Data\Config.ini", "INFO", "MODULES", "S", (Amount + 1)
fWriteValue APPdir & "Data\Config.ini", "MODULES", "MOD" & Amount, "S", Text7.Text

Call Form_Load
End Sub

Private Sub Command2_Click()
CreatePath (APPdir & "data")
fWriteValue APPdir & "Data\Config.ini", "MODULES", "MOD" & Amount, "S", Text7.Text
End Sub

Private Sub Form_Load()
Dim TempInt As Integer
Dim TempStr As String
ShadeIt Me, Light
ShadeIt List1, Dark
APPdir = App.Path & slashval(App.Path)
CreatePath (APPdir & "data")
fReadValue APPdir & "Data\Config.ini", "INFO", "MODULES", "S", "0", Amount
List1.Clear
    Do While TempInt < Amount
    fReadValue APPdir & "Data\Config.ini", "MODULES", "MOD" & TempInt, "S", "", TempStr
    List1.AddItem TempStr
    TempInt = TempInt + 1
    DoEvents
    Loop
End Sub
