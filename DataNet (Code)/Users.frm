VERSION 5.00
Begin VB.Form Users 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Users"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Unlo 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3000
      Top             =   120
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
      Height          =   4335
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   2895
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
         TabIndex        =   13
         ToolTipText     =   "Select a domain to look in."
         Top             =   360
         Width           =   2655
      End
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
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Select a user from this list to edit/delete."
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete User"
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
         Left            =   600
         TabIndex        =   11
         Top             =   3840
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "User Details"
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
      Height          =   4335
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Disabled"
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
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   3240
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   1320
         ScaleHeight     =   1545
         ScaleWidth      =   2145
         TabIndex        =   14
         Top             =   1560
         Width           =   2175
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Guest"
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
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   2055
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Normal"
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
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Service Admin"
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
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "System Admin"
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
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Real Name:"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text2 
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
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "UserName:"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text4 
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
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Password:"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text6 
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
         IMEMode         =   3  'DISABLE
         Left            =   1320
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "Authority:"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New"
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
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Replace Selected"
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
         TabIndex        =   1
         Top             =   3840
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Users"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ArchiveUser(0 To 1000) As UserInfo
Dim UserNo As Integer
Dim Pathy As String

Private Sub Combo1_Change()
Dim TempInt As Integer
Dim Datas() As String
Dim Fringo As Integer
Datas = Split(Combo1.Text, "]")
CreatePath (Master.Domain(Datas(0)).Home & slashval(Master.Domain(Datas(0)).Home) & "Data")
Pathy = Master.Domain(Datas(0)).Home & slashval(Master.Domain(Datas(0)).Home) & "Data\users.ini"
Fringo = FreeFile
CreatePath FolderFromPath(Pathy)
Open Pathy For Binary As Fringo
Close Fringo
Me.Caption = "Users"

fReadValue Pathy, "SETUP", "USERSNO", "S", "0", UserNo
List1.Clear
    Do While TempInt < UserNo
    fReadValue Pathy, "USER" & TempInt, "USER", "S", "Guest", ArchiveUser(TempInt).UserName
    fReadValue Pathy, "USER" & TempInt, "NAME", "S", "Guest", ArchiveUser(TempInt).Realname
    fReadValue Pathy, "USER" & TempInt, "PASS", "S", "", ArchiveUser(TempInt).Password
    fReadValue Pathy, "USER" & TempInt, "RIGHTS", "S", "100", ArchiveUser(TempInt).Rights
    fReadValue Pathy, "USER" & TempInt, "STATUS", "S", "200", ArchiveUser(TempInt).Status
    List1.AddItem TempInt & "] " & ArchiveUser(TempInt).UserName
    ArchiveUser(TempInt).Stage = Clear
    TempInt = TempInt + 1
    DoEvents
    Loop
End Sub

Private Sub Combo1_Click()
Call Combo1_Change
End Sub

Private Sub Command1_Click()
Dim Datas2() As String
Datas2 = Split(Combo1.Text, "]")
    If Master.Domain(Datas2(0)).UserExists(Replace(Text4.Text, " ", "_")) = True Then
    Me.Caption = "Users - User already exists"
    Exit Sub
    End If
Me.Caption = "Users"
fWriteValue Pathy, "USER" & UserNo, "USER", "S", Replace(Text4.Text, " ", "_")
fWriteValue Pathy, "USER" & UserNo, "NAME", "S", Text2.Text
fWriteValue Pathy, "USER" & UserNo, "PASS", "S", Text6.Text
    If Check1.Value = 1 Then
    fWriteValue Pathy, "USER" & UserNo, "STATUS", "S", 200
    Else
    fWriteValue Pathy, "USER" & UserNo, "STATUS", "S", 100
    End If
If Option1.Value = True Then fWriteValue Pathy, "USER" & UserNo, "RIGHTS", "S", Accy.aSystemAdmin
If Option2.Value = True Then fWriteValue Pathy, "USER" & UserNo, "RIGHTS", "S", Accy.aServiceAdmin
If Option3.Value = True Then fWriteValue Pathy, "USER" & UserNo, "RIGHTS", "S", Accy.aStandard
If Option4.Value = True Then fWriteValue Pathy, "USER" & UserNo, "RIGHTS", "S", Accy.aGuest
fWriteValue Pathy, "SETUP", "USERSNO", "S", (UserNo + 1)
DoEvents

Dim Dat() As String
Dat = Split(Combo1.Text, "]")
Master.Domain(Dat(0)).IniUser

Call Combo1_Change
End Sub

Private Sub Command2_Click()
Dim Datas() As String
Dim Datas2() As String
Datas = Split(List1.Text, "]")
Datas2 = Split(Combo1.Text, "]")
    If List1.ListIndex < 0 Then Exit Sub
    If (UCase(Trim(Text4.Text)) = UCase(Trim(Datas(1)))) = False Then
        If Master.Domain(Datas2(0)).UserExists(Replace(Text4.Text, " ", "_")) = True Then
        Me.Caption = "Users - User already exists"
        Exit Sub
        End If
    End If
Me.Caption = "Users"

fWriteValue Pathy, "USER" & Datas(0), "USER", "S", Replace(Text4.Text, " ", "_")
fWriteValue Pathy, "USER" & Datas(0), "NAME", "S", Text2.Text
fWriteValue Pathy, "USER" & Datas(0), "PASS", "S", Text6.Text
    If Check1.Value = 1 Then
    fWriteValue Pathy, "USER" & Datas(0), "STATUS", "S", 200
    Else
    fWriteValue Pathy, "USER" & Datas(0), "STATUS", "S", 100
    End If
If Option1.Value = True Then fWriteValue Pathy, "USER" & Datas(0), "RIGHTS", "S", Accy.aSystemAdmin
If Option2.Value = True Then fWriteValue Pathy, "USER" & Datas(0), "RIGHTS", "S", Accy.aServiceAdmin
If Option3.Value = True Then fWriteValue Pathy, "USER" & Datas(0), "RIGHTS", "S", Accy.aStandard
If Option4.Value = True Then fWriteValue Pathy, "USER" & Datas(0), "RIGHTS", "S", Accy.aGuest

DoEvents

Dim Dat() As String
Dat = Split(Combo1.Text, "]")
Master.Domain(Dat(0)).Initialise

Call Combo1_Change
End Sub

Private Sub Command4_Click()
Dim Datas() As String
Dim TempInt As Integer
Dim TempInt2 As Integer
Dim TempSel As Integer
TempSel = List1.ListIndex

    If List1.ListIndex >= 0 Then
    Datas = Split(List1.Text, "]")
    fWriteValue Pathy, "USER" & Datas(0), "STATUS", "S", 300
    
    Datas = Split(Combo1.Text, "]")
    Master.Domain(Datas(0)).IniUser
    
    Call Combo1_Change
    DoEvents
    End If

    If TempSel < List1.ListCount Then
    List1.ListIndex = TempSel
    Else
        If List1.ListCount > 0 Then
        List1.ListIndex = List1.ListCount - 1
        End If
    End If
End Sub

Private Sub Form_Load()
Dim TempInt As Integer
TempInt = 1

ShadeIt Me, Light
ShadeIt Frame1, Light
ShadeIt Frame2, Light
ShadeIt Check1, Light

ShadeIt Text1, Dark
ShadeIt Text3, Dark
ShadeIt Text5, Dark
ShadeIt Text7, Dark
ShadeIt Picture1, Dark
ShadeIt Option1, Dark
ShadeIt Option2, Dark
ShadeIt Option3, Dark
ShadeIt Option4, Dark


    Do While TempInt <= NoDomains
        If Master.Domain(TempInt).Status <> aDeleted Then
        Combo1.AddItem TempInt & "] " & Master.Domain(TempInt).Description
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
    
    If Combo1.ListCount = 0 Then
    Unlo.Enabled = True
    Say "You must create a domain before you setup user accounts."
    Else
    Combo1.ListIndex = 0
    End If
    
End Sub

Private Sub List1_Click()
Me.Caption = "Users"
    If List1.ListIndex >= 0 Then
    Text2.Text = ArchiveUser(List1.ListIndex).Realname
    Text4.Text = ArchiveUser(List1.ListIndex).UserName
    Text6.Text = ArchiveUser(List1.ListIndex).Password
    If ArchiveUser(List1.ListIndex).Rights = aGuest Then Option4.Value = True
    If ArchiveUser(List1.ListIndex).Rights = aStandard Then Option3.Value = True
    If ArchiveUser(List1.ListIndex).Rights = aServiceAdmin Then Option2.Value = True
    If ArchiveUser(List1.ListIndex).Rights = aSystemAdmin Then Option1.Value = True
        If ArchiveUser(List1.ListIndex).Status = aAlive Then
        Check1.Value = 0
        Else
        Check1.Value = 1
        End If
    Else
    Text2.Text = ""
    Text4.Text = ""
    Text6.Text = ""
    Option4.Value = True
    Check1.Value = 0
    End If
End Sub

Private Sub Unlo_Timer()
Unload Me
End Sub
