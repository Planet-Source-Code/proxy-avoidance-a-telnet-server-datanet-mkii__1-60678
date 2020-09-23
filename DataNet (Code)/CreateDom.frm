VERSION 5.00
Begin VB.Form CreateDom 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Domain"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "CreateDom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin DataNet.PortStatus PortStatus 
      Height          =   495
      Left            =   2400
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _extentx        =   3413
      _extenty        =   873
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   6120
      Tag             =   "0"
      Top             =   6720
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C00000&
      Caption         =   "Script"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   "Configure Modules"
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
         TabIndex        =   19
         Top             =   4080
         Width           =   1935
      End
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
         Height          =   3405
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Text            =   "CreateDom.frx":23D2
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "<Inserts a script that will configure the Standard modules>"
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
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   5295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   7200
      TabIndex        =   8
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8880
      TabIndex        =   7
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C00000&
      Caption         =   "Display Message"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   10215
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "CreateDom.frx":242B
         ToolTipText     =   "A short greeting, normally indicates the domain name."
         Top             =   360
         Width           =   9975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "General"
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
      Height          =   2415
      Left            =   5760
      TabIndex        =   2
      Top             =   2040
      Width           =   4575
      Begin VB.TextBox Text9 
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
         Left            =   1800
         TabIndex        =   18
         Text            =   "23"
         ToolTipText     =   "The port number must be unique to this domain, else Data Net will crash."
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Internal Port:"
         Top             =   1080
         Width           =   1695
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
         Left            =   1800
         TabIndex        =   14
         Text            =   "New Domain"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Domain Name:"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Text            =   "Status:"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1800
         ScaleHeight     =   825
         ScaleWidth      =   2625
         TabIndex        =   9
         Top             =   1440
         Width           =   2655
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            Caption         =   "Enabled"
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
            TabIndex        =   11
            Top             =   120
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
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
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Reference:"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text8 
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
         Left            =   1800
         TabIndex        =   3
         Text            =   "NEW"
         ToolTipText     =   "The port number must be unique to this domain, else Data Net will crash."
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      Caption         =   "Modules"
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
      Height          =   1935
      Left            =   5760
      TabIndex        =   0
      Top             =   4560
      Width           =   4575
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
         Height          =   1470
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         ToolTipText     =   "Select the plugins that you would like to be open to users that use this domain."
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "CreateDom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ery As String
Private Sub Command1_Click()
On Error GoTo endi:
Dim NewDomain As Integer
Dim Fringo As Integer
Dim TempInt As Integer
Dim TempInt2 As Integer
Dim FSO As New FileSystemObject

    If Me.Tag = "CREATE" Then
            If FSO.FolderExists(Trim((App.path & slashval(App.path) & "Domains\" & Text7.Text))) = True Then
            Timer1.Tag = 0
            Ery = "Must Choose a Different Domain Name!"
            Timer1.Enabled = True
            Exit Sub
            End If
            If Master.RefExists(Trim(Text8.Text)) = True Then
            Timer1.Tag = 0
            Ery = "Must Choose a Different Reference Name!"
            Timer1.Enabled = True
            Exit Sub
            End If
    NoDomains = NoDomains + 1
    NewDomain = NoDomains
    Load Master.Domain(NewDomain)
    End If

    If Left(Me.Tag, 6) = "MODIFY" Then
    NewDomain = Right(Me.Tag, Len(Me.Tag) - 6)
        If Not UCase(Trim(Master.Domain(NewDomain).Home)) = UCase(Trim((App.path & slashval(App.path) & "Domains\" & Text7.Text))) Then
            If FSO.FolderExists(Trim((App.path & slashval(App.path) & "Domains\" & Text7.Text))) = True Then
            Timer1.Tag = 0
            Ery = "Must Choose a Different Domain Name!"
            Timer1.Enabled = True
            Exit Sub
            End If
        CreatePath (App.path & slashval(App.path) & "Domains\" & Text7.Text)
            If FSO.FolderExists(Master.Domain(NewDomain).Home) = True Then
            FSO.CopyFolder Master.Domain(NewDomain).Home, App.path & slashval(App.path) & "Domains\" & Text7.Text
            End If
        End If
        If Not UCase(Trim(Master.Domain(NewDomain).Refer)) = UCase(Trim(Text8.Text)) Then
            If Master.RefExists(Trim(Text8.Text)) = True Then
            Timer1.Tag = 0
            Ery = "Must Choose a Different Reference Name!"
            Timer1.Enabled = True
            Exit Sub
            End If
        End If
    End If

Master.Domain(NewDomain).Home = App.path & slashval(App.path) & "Domains\" & Text7.Text
CreatePath (Master.Domain(NewDomain).Home & slashval(Master.Domain(NewDomain).Home) & "Users")
If Option1.Value = True Then Master.Domain(NewDomain).Status = aAlive
If Option2.Value = True Then Master.Domain(NewDomain).Status = aDisabled
Master.Domain(NewDomain).Description = Text7.Text

If Number(Text9.Text) = True Then
    If (Text9.Text > 0) And (Text9.Text <= 30000) Then
        If (Len(Trim(PortStatus.ApplicationUsingPort(Text9.Text, TCP))) > 0) And (Master.Domain(NewDomain).Port <> Text9.Text) Then
        Timer1.Tag = 0
        Ery = "Specified Port In Use!!!"
        Timer1.Enabled = True
            If Me.Tag = "CREATE" Then
            Unload Master.Domain(NewDomain)
            NoDomains = NoDomains - 1
            End If
        Exit Sub
        End If
    Master.Domain(NewDomain).Port = Text9.Text
    End If
End If

Master.Domain(NewDomain).Greetings = Text2.Text
Master.Domain(NewDomain).Script = Text3.Text
Master.Domain(NewDomain).Refer = Text8.Text

Fringo = FreeFile
CreatePath (Master.Domain(NewDomain).Home & slashval(Master.Domain(NewDomain).Home) & "Data")
Open Master.Domain(NewDomain).Home & slashval(Master.Domain(NewDomain).Home) & "Data\modules.ini" For Binary As Fringo
Close Fringo
TempInt = 0
TempInt2 = 0
    Do While TempInt < List1.ListCount
        If List1.Selected(TempInt) = True Then
        fWriteValue Master.Domain(NewDomain).Home & slashval(Master.Domain(NewDomain).Home) & "Data\modules.ini", "MODULES", "MOD" & TempInt2, "S", List1.List(TempInt)
        TempInt2 = TempInt2 + 1
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
fWriteValue Master.Domain(NewDomain).Home & slashval(Master.Domain(NewDomain).Home) & "Data\modules.ini", "INFO", "MODULES", "S", TempInt2

Master.Domain(NewDomain).Initialise
Master.Domain(NewDomain).Action aListen
Master.Domain(NewDomain).IniMods
Load Domains
Domains.Show
Unload Me
Exit Sub
endi:
    If Me.Tag = "CREATE" Then
    Unload Master.Domain(NewDomain)
    NoDomains = NoDomains - 1
    End If
Timer1.Tag = 0
Ery = "Error when creating/modifying domain!!!"
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
Load Domains
Domains.Show
Unload Me
End Sub

Private Sub Command3_Click()
Text3.Text = Text3.Text & vbCrLf
Text3.Text = Text3.Text & "Sys.Init" & vbCrLf
Text3.Text = Text3.Text & "Sys.Force 1" & vbCrLf
Text3.Text = Text3.Text & "Txt.Wrap 1" & vbCrLf
Text3.Text = Text3.Text & "Txt.Pnt 1" & vbCrLf

End Sub

Private Sub Form_Load()
Dim TempInt As Integer
Dim TempStr As String
APPdir = App.path & slashval(App.path)
CreatePath (APPdir & "data")
fReadValue APPdir & "Data\Config.ini", "INFO", "MODULES", "S", "0", Amount

    Do While TempInt < Amount
    fReadValue APPdir & "Data\Config.ini", "MODULES", "MOD" & TempInt, "S", "", TempStr
    List1.AddItem TempStr
    TempInt = TempInt + 1
    DoEvents
    Loop
ShadeIt Me, Light
ShadeIt Frame1, Light
ShadeIt Frame2, Light
ShadeIt Frame3, Light
ShadeIt Frame4, Light

ShadeIt Text1, Dark
ShadeIt Text4, Dark
ShadeIt Text5, Dark
ShadeIt Text6, Dark
ShadeIt Picture1, Dark
ShadeIt Option1, Dark
ShadeIt Option2, Dark
ShadeIt List1, Dark

End Sub

Private Sub Timer1_Timer()
Timer1.Tag = Timer1.Tag + 1
    
    If Me.Caption = "Create Domain" Then
    Me.Caption = "Create Domain - " & Ery
    Else
    Me.Caption = "Create Domain"
    End If
    
    If Timer1.Tag >= 10 Then
    Timer1.Tag = 0
    Timer1.Enabled = False
    Me.Caption = "Create Domain"
    End If
End Sub

