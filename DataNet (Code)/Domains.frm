VERSION 5.00
Begin VB.Form Domains 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Domains"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5280
   Icon            =   "Domains.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Add/Modify/Delete Domain"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      Begin VB.CommandButton Command3 
         Caption         =   "Modify Domain"
         Enabled         =   0   'False
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
         Left            =   3000
         TabIndex        =   4
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Remove Domain"
         Enabled         =   0   'False
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
         Left            =   3000
         TabIndex        =   3
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Domain"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   2040
         Width           =   1815
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
         Height          =   3180
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Domains"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim TempInt As Integer
Load CreateDom
CreateDom.Tag = "CREATE"
TempInt = 0
    Do While TempInt < CreateDom.List1.ListCount
    CreateDom.List1.Selected(TempInt) = True
    TempInt = TempInt + 1
    DoEvents
    Loop
CreateDom.Text9.Text = Round(Rnd * 20000, 0) + 1
CreateDom.Show
Unload Me
End Sub

Private Sub Command2_Click()
On Error GoTo endi
Dim Datas() As String
Datas = Split(List1.Text, "]")
Master.Domain(Datas(0)).Status = aDeleted
List1.RemoveItem List1.ListIndex

Say "It is recommended that you restart DataNet after removing a domain."
endi:
End Sub

Private Sub Command3_Click()
Dim Datas() As String
Dim TempInt As Integer
Dim TempInt2 As Integer
Dim TempStr As String
Dim MaxMod As Integer
Datas = Split(List1.Text, "]")

Load CreateDom
CreateDom.Tag = "MODIFY" & Datas(0)
CreateDom.Text2.Text = Master.Domain(Datas(0)).Greetings
CreateDom.Text3.Text = Master.Domain(Datas(0)).Script
CreateDom.Text7.Text = Master.Domain(Datas(0)).Description
CreateDom.Text8.Text = Master.Domain(Datas(0)).Refer
CreateDom.Text9.Text = Master.Domain(Datas(0)).Port

    If Master.Domain(Datas(0)).Status = aAlive Then
    CreateDom.Option1.Value = True
    CreateDom.Option2.Value = False
    End If
    
    If Master.Domain(Datas(0)).Status = aDisabled Then
    CreateDom.Option1.Value = False
    CreateDom.Option2.Value = True
    End If

TempInt = 0
CreatePath (Master.Domain(Datas(0)).Home & slashval(Master.Domain(Datas(0)).Home) & "Data")
fReadValue Master.Domain(Datas(0)).Home & slashval(Master.Domain(Datas(0)).Home) & "Data\modules.ini", "INFO", "MODULES", "S", "0", MaxMod
    Do While TempInt < MaxMod
    fReadValue Master.Domain(Datas(0)).Home & slashval(Master.Domain(Datas(0)).Home) & "Data\modules.ini", "MODULES", "MOD" & TempInt, "S", "0", TempStr
    TempInt2 = 0
        Do While TempInt2 < CreateDom.List1.ListCount
            If CreateDom.List1.List(TempInt2) = TempStr Then
            CreateDom.List1.Selected(TempInt2) = True
            End If
        TempInt2 = TempInt2 + 1
        DoEvents
        Loop
    TempInt = TempInt + 1
    DoEvents
    Loop

CreateDom.Show
Unload Me
End Sub

Private Sub Form_Load()
Dim TempInt As Integer
TempInt = 1
ShadeIt Me, Light
ShadeIt Frame1, Light
ShadeIt List1, Dark
    Do While TempInt <= NoDomains
        If Master.Domain(TempInt).Status <> aDeleted Then
        List1.AddItem TempInt & "] " & Master.Domain(TempInt).Description
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Sub

Private Sub List1_Click()
    If List1.ListIndex >= 0 Then
    Command2.Enabled = True
    Command3.Enabled = True
    End If
End Sub
