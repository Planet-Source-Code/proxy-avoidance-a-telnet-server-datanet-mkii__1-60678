VERSION 5.00
Begin VB.Form CMDres 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Command Restrictions"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Unlo 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
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
      Left            =   5880
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
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
      Left            =   5760
      TabIndex        =   5
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Keyword List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   4335
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.ListBox List1 
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
         Height          =   2970
         ItemData        =   "CMDres.frx":0000
         Left            =   120
         List            =   "CMDres.frx":0010
         TabIndex        =   2
         ToolTipText     =   "Select a user from this list to edit/delete."
         Top             =   720
         Width           =   2655
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Select a domain to look in."
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "CMDres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim Dat() As String
Dat = Split(Combo1.Text, "]")

If List1.ListIndex = 0 Then Master.Domain(Dat(0)).CMDrest4 = Text3.Text
If List1.ListIndex = 1 Then Master.Domain(Dat(0)).CMDrest3 = Text3.Text
If List1.ListIndex = 2 Then Master.Domain(Dat(0)).CMDrest2 = Text3.Text
If List1.ListIndex = 3 Then Master.Domain(Dat(0)).CMDrest1 = Text3.Text

End Sub

Private Sub Form_Load()
Dim TempInt As Integer
TempInt = 1
    Do While TempInt <= NoDomains
        If Master.Domain(TempInt).Status <> aDeleted Then
        Combo1.AddItem TempInt & "] " & Master.Domain(TempInt).Description
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
    
    If Combo1.ListCount = 0 Then
    Unlo.Enabled = True
    Say "You must create a domain before you setup command restrictions."
    Else
    Combo1.ListIndex = 0
    End If



ShadeIt Me, Light
ShadeIt Frame1, Light
ShadeIt Frame2, Light
ShadeIt List1, Dark

End Sub

Private Sub List1_Click()
Dim Dat() As String
Dat = Split(Combo1.Text, "]")

If List1.ListIndex = 0 Then Text3.Text = Master.Domain(Dat(0)).CMDrest4
If List1.ListIndex = 1 Then Text3.Text = Master.Domain(Dat(0)).CMDrest3
If List1.ListIndex = 2 Then Text3.Text = Master.Domain(Dat(0)).CMDrest2
If List1.ListIndex = 3 Then Text3.Text = Master.Domain(Dat(0)).CMDrest1
End Sub

Private Sub Unlo_Timer()
Unload Me
End Sub
