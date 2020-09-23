VERSION 5.00
Begin VB.Form LozMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LozWare Alert"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   Icon            =   "LozMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5520
   Begin VB.CommandButton Command2 
      Caption         =   "Process Alert"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "LozMes.frx":23D2
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "LozMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Follow As String
Public Codey As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Dim Fringo As Integer
Dim StringIn As String
    If Len(Trim(Follow)) > 0 Then
    Shell "explorer.exe " & Follow, vbMaximizedFocus
    End If
Fringo = FreeFile
Open App.path & slashval(App.path) & "Processed.dat" For Binary As Fringo
StringIn = Codey & "|"
Put Fringo, LOF(Fringo) + 1, CStr(StringIn)
Close Fringo
Unload Me
End Sub

Private Sub Form_Load()
ShadeIt Text1, Dark
ShadeIt Me, Light
End Sub
