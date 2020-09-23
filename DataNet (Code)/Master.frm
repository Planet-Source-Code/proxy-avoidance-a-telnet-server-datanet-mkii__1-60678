VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm Master 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DataNet MKII"
   ClientHeight    =   4440
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8715
   Icon            =   "Master.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Master.frx":23D2
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   335
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   8655
      TabIndex        =   2
      Top             =   0
      Width           =   8715
      Begin VB.TextBox SayTxt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   4110
      Left            =   0
      ScaleHeight     =   4050
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   330
      Width           =   3135
      Begin DataNet.PortStatus PortStatus 
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   1920
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   873
      End
      Begin VB.Timer LozAlert 
         Interval        =   10000
         Left            =   1200
         Top             =   840
      End
      Begin MSWinsockLib.Winsock LiveUp 
         Left            =   720
         Top             =   840
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin DataNet.Domain Domain 
         Height          =   465
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   1320
         Visible         =   0   'False
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   820
      End
      Begin MSWinsockLib.Winsock WanListen 
         Left            =   1200
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock WAN 
         Index           =   0
         Left            =   240
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock LAN 
         Index           =   0
         Left            =   720
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer ListRefresh 
         Interval        =   100
         Left            =   1680
         Top             =   2520
      End
      Begin VB.Timer Speed 
         Interval        =   1000
         Left            =   1200
         Top             =   2520
      End
      Begin VB.Timer NoticeTim 
         Interval        =   1000
         Left            =   720
         Top             =   2520
      End
      Begin MSComctlLib.ImageList MasterList 
         Left            =   360
         Top             =   3000
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Master.frx":46F4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ClientList 
         Height          =   4095
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   7223
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "MasterList"
         ForeColor       =   0
         BackColor       =   12582912
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Menu Fi 
      Caption         =   "Server"
      Begin VB.Menu Sta 
         Caption         =   "Listening"
         Checked         =   -1  'True
      End
      Begin VB.Menu Ref 
         Caption         =   "Refresh Domains"
      End
      Begin VB.Menu term 
         Caption         =   "Terminate Connections"
      End
      Begin VB.Menu Sepy 
         Caption         =   "-"
      End
      Begin VB.Menu ex 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu ser 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu shy 
         Caption         =   "Show"
      End
      Begin VB.Menu bty 
         Caption         =   "About"
      End
      Begin VB.Menu sre 
         Caption         =   "-"
      End
      Begin VB.Menu exy 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Ma 
      Caption         =   "Management"
      Begin VB.Menu Dom 
         Caption         =   "Domains"
      End
      Begin VB.Menu Use 
         Caption         =   "Users"
      End
      Begin VB.Menu Mod 
         Caption         =   "Modules"
      End
      Begin VB.Menu CMDresy 
         Caption         =   "Command Restrictions"
      End
      Begin VB.Menu Sep 
         Caption         =   "-"
      End
      Begin VB.Menu Cent 
         Caption         =   "General"
      End
   End
   Begin VB.Menu To 
      Caption         =   "Tools"
      Begin VB.Menu Traf 
         Caption         =   "Traffic Moniter"
      End
      Begin VB.Menu Reg 
         Caption         =   "Register DLL"
      End
      Begin VB.Menu era 
         Caption         =   "-"
      End
      Begin VB.Menu lozw 
         Caption         =   "LozWare Alerts"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu hlpy 
         Caption         =   "Help"
      End
      Begin VB.Menu ab 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Master"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_TOOLTIP As Integer = 64
Const NIF_ICON = &H2
Const NIF_MESSAGE = &H1
Const NIF_TIP = &H4
Const NIM_ADD = &H0
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const WM_MOUSEMOVE = &H200
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_LBUTTONDBLCLK = &H203
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_RBUTTONDBLCLK = &H206
Private Type NOTIFYICONDATA
    cbSize           As Long
    hWnd             As Long
    uID              As Long
    uFlags           As Long
    uCallbackMessage As Long
    hIcon            As Long
    szTip            As String * MAX_TOOLTIP
End Type
Private nfIconData As NOTIFYICONDATA
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public NOTICEexpirey As Integer
Dim SessInfo(0 To 1000) As ConInfo
Dim LogStr As String

Private Sub ab_Click()
Load About
About.Show
End Sub

Private Sub bty_Click()
Load About
About.Show

End Sub

Private Sub Cent_Click()
Load Central
Central.Show
End Sub

Private Sub ClientList_DblClick()
Dim Indy As Integer
    If ItemSelected(ClientList) = False Then
    Exit Sub
    End If
Indy = Right(ClientList.SelectedItem.Key, Len(ClientList.SelectedItem.Key) - 1)

    If isLoaded(LogBox(Indy)) = False Then
    Set LogBox(Indy) = New frmTelnet
    Load LogBox(Indy)
    LogBox(Indy).Tag = Indy
    LogBox(Indy).Initialise
    LogBox(Indy).Caption = WAN(Indy).RemoteHostIP & " - View Session"
    LogBox(Indy).Show
    End If
End Sub

Private Sub CMDresy_Click()
Load CMDres
CMDres.Show
End Sub

Private Sub Dom_Click()
Load Domains
Domains.Show
End Sub

Private Sub Domain_Message(Index As Integer, Message As String)
Say Domain(Index).Description & ": " & Message
End Sub

Private Sub Ex_Click()
Master.Tag = "CLOSE"
Unload Master
End Sub

Private Sub exy_Click()
Master.Tag = "CLOSE"
Unload Master
End Sub

Private Sub hlpy_Click()
Load Help
Help.Show
End Sub

Private Sub LAN_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
LAN(Index).GetData Data

    If Len(Replace(Data, ">" & Chr(231) & Chr(12), ">")) < Len(Data) Then
    SendDat Replace(Data, ">" & Chr(231) & Chr(12), ">" & SessInfo(Index).Data), Index
    Else
    SendDat Data, Index
    End If
End Sub

Private Sub ListRefresh_Timer()
Dim TempInt As Integer
Dim Indy As Integer
    Do While TempInt < ClientList.ListItems.Count
    TempInt = TempInt + 1
    Indy = Right(ClientList.ListItems.Item(TempInt).Key, Len(ClientList.ListItems.Item(TempInt).Key) - 1)
        If WAN(Indy).State <> sckConnected Then
        WAN(Indy).Close
        LAN(Indy).Close
        ClientList.ListItems.Remove (TempInt)
        End If
    DoEvents
    Loop
End Sub

Private Sub LiveUp_Connect()
LiveUp.SendData UserName & "|" & Versy & "|0"
DoEvents
End Sub

Private Sub LiveUp_DataArrival(ByVal bytesTotal As Long)
On Error GoTo endi
Dim Data As String
Dim Contents As String
Dim Fringo As Integer
Dim SplitUp() As String
Dim Codes() As String
Dim TempInt As Integer
LiveUp.GetData Data
SplitUp = Split(Data, "|")
'0 = Code
'1 = Follow Path
'2 = Message
'3 = Index
Fringo = FreeFile
Open App.path & slashval(App.path) & "Processed.dat" For Binary As Fringo
Contents = String$(LOF(Fringo), " ")
Get Fringo, , Contents
Close Fringo
Codes = Split(Contents, "|")
TempInt = 0

    Do While TempInt <= UBound(Codes)
        If (Number(Trim(Codes(TempInt))) = True) And (Number(SplitUp(0)) = True) Then
            If CInt(SplitUp(0)) = CInt(Trim(Codes(TempInt))) Then
            LiveUp.SendData "---|" & Versy & "|" & (CInt(SplitUp(3)) + 1)
            DoEvents
            Exit Sub
            End If
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
Set Alert(CInt(SplitUp(3))) = New LozMes
Load Alert(CInt(SplitUp(3)))
Alert(CInt(SplitUp(3))).Text1.Text = SplitUp(2)
Alert(CInt(SplitUp(3))).Follow = SplitUp(1)
Alert(CInt(SplitUp(3))).Codey = Trim(Replace(SplitUp(0), vbCrLf, ""))
Alert(CInt(SplitUp(3))).Show
LiveUp.SendData "---|" & Versy & "|" & (CInt(SplitUp(3)) + 1)
DoEvents
endi:
End Sub

Private Sub LozAlert_Timer()
    If LiveUp.State <> sckConnected Then
    LiveUp.Close
    LiveUp.Connect Servy, 500
    DoEvents
    End If
End Sub

Private Sub lozw_Click()
Load LozCon
LozCon.Show
End Sub

Private Sub MDIForm_Load()
'On Error GoTo endi
Dim Fringo As Integer
Dim Viewy As Integer
Dim TmpString As String
Dim TempInt As Integer
Dim HDCm As Long
Dim BltBit As Long
Dim TempData(0 To 6) As Variant
    If App.PrevInstance = True Then
    MsgBox "DataNet MKII is already running!", vbInformation, "Program Running"
    End
    End If
    
Versy = "1.30"
dColour = RGB(173, 205, 242)
fColour = RGB(205, 222, 242)
dColourFor = RGB(0, 0, 0)
fColourFor = RGB(0, 0, 0)

ShadeIt SayTxt, Light
ShadeIt ClientList, Light
    
APPdir = App.path & slashval(App.path)
CreatePath (APPdir & "data")
Fringo = FreeFile
Open APPdir & "Data\Config.ini" For Binary As Fringo
Close Fringo
    If Exist(APPdir & "Data\Config.ini") = False Then
    WANport = 23
    AccessCode = 1234
    NoDomains = 0
    Viewy = 0
    BandMoniter = False
    Greet = vbCrLf & "DataNet MKII Server" & vbCrLf & "Framework Version " & Versy
    
    Else
    

    fReadValue APPdir & "Data\Config.ini", "SETUP", "BAND", "S", "0", TempInt
    BandMoniter = CBool(TempInt)
    fReadValue APPdir & "Data\Config.ini", "SETUP", "PORT", "S", "23", WANport
    fReadValue APPdir & "Data\Config.ini", "SETUP", "ACCESS", "S", "", AccessCode
    fReadValue APPdir & "Data\Config.ini", "SETUP", "DOMAINS", "S", "0", NoDomains
    fReadValue APPdir & "Data\Config.ini", "SETUP", "VIEW", "S", "0", Viewy
    fReadValue APPdir & "Data\Config.ini", "SETUP", "LOZWARE", "S", "LOZWARE.ZAPTO.ORG", Servy
    ClientList.View = Viewy
    
    
    Fringo = FreeFile
    CreatePath (APPdir & slashval(APPdir) & "Data")
    Open APPdir & slashval(APPdir) & "Data\Display.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Greet = TmpString
    End If

TempInt = 1
    Do While TempInt <= NoDomains

    fReadValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt, "REFER", "S", "ERROR", TempData(6)
    fReadValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt, "PORT", "S", "1", TempData(0)
    fReadValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt, "HOME", "S", "C:\", TempData(1)
    fReadValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt, "NAME", "S", "ERROR", TempData(2)
    fReadValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt, "STATUS", "S", "200", TempData(3)
    Load Domain(TempInt)
    CreatePath (Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data")
    Domain(TempInt).Port = TempData(0)
        If Left(TempData(1), Len(App.path & slashval(App.path) & "Domains\")) = App.path & slashval(App.path) & "Domains\" Then
        Domain(TempInt).Home = TempData(1)
        Else
        Domain(TempInt).Home = App.path & slashval(App.path) & "Domains\" & TempData(1)
        End If
    
    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Display.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).Greetings = TmpString
    
    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest1.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).CMDrest1 = TmpString

    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest2.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).CMDrest2 = TmpString

    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest3.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).CMDrest3 = TmpString

    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest4.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).CMDrest4 = TmpString

    
    Fringo = FreeFile
    Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Script.dat" For Binary As Fringo
    TmpString = String$(LOF(Fringo), " ")
    Get Fringo, , TmpString
    Close Fringo
    Domain(TempInt).Script = TmpString
    Domain(TempInt).Refer = Replace(TempData(6), " ", "_")
    Domain(TempInt).Description = TempData(2)
    Domain(TempInt).Status = TempData(3)
    Domain(TempInt).Initialise
    Domain(TempInt).Action aListen
    TempInt = TempInt + 1
    DoEvents
    Loop

    If WanListen.State <> sckListening Then
    WanListen.Close
    WanListen.LocalPort = WANport
        If Len(Trim(PortStatus.ApplicationUsingPort(WANport, TCP))) > 0 Then
        Say Trim(PortStatus.ApplicationUsingPort(WANport, TCP)) & " is using a port required by DataNet"
        Exit Sub
        End If
    WanListen.Listen
    End If


With nfIconData
    .hWnd = Me.hWnd
    .uID = Me.Icon
    .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon.Handle
    .szTip = "DataNet MKII" & vbNullChar
    .cbSize = Len(nfIconData)
End With
Shell_NotifyIcon NIM_ADD, nfIconData

LiveUp.Close
LiveUp.Connect Servy, 500
Randomize

fWriteValue "HKCU", "Software\Microsoft\Windows\CurrentVersion\Run", "DataNet", "S", App.path & slashval(App.path) & App.EXEName & ".exe HIDE"
Exit Sub
endi:
Say "There was an error loading DataNet!"
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single

lMsg = X / Screen.TwipsPerPixelX
Select Case lMsg
    Case WM_RBUTTONUP
    Me.PopupMenu ser
    Case WM_LBUTTONDBLCLK
    Master.Visible = True
    Master.Show
    Case Else
End Select
End Sub

Private Sub MDIForm_Resize()
On Error GoTo endi:
    If Master.Width < 6900 Then
    Master.Width = 6900
    End If

    If Master.Height < 4400 Then
    Master.Height = 4400
    End If
SayTxt.Width = Master.Width
ClientList.Height = Master.Height - 1200
endi:
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo endo
    If Master.Tag = "CLOSE" Then
    Dim Fringo As Integer
    Dim TempInt As Integer
    Dim TempInt2 As Integer
    Dim Tempy As String
    
    Fringo = FreeFile
    CreatePath (APPdir & "data")
    Open APPdir & "Data\Config.ini" For Binary As Fringo
    Close Fringo
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "PORT", "S", WANport
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "ACCESS", "S", AccessCode
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "VIEW", "S", ClientList.View
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "LOZWARE", "S", Servy

    If BandMoniter = True Then
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "BAND", "S", "1"
    Else
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "BAND", "S", "0"
    End If

    Tempy = LogStr
    LogStr = ""
    Fringo = FreeFile
    Open App.path & slashval(App.path) & "Log.txt" For Binary As Fringo
    Put Fringo, LOF(Fringo) + 1, Tempy
    Close Fringo
    
    APPdir = App.path & slashval(App.path)
    Fringo = FreeFile
    CreatePath (APPdir & slashval(APPdir) & "Data")
        If Exist(APPdir & slashval(APPdir) & "Data\Display.dat") Then
        Kill APPdir & slashval(APPdir) & "Data\Display.dat"
        DoEvents
        End If
    Open APPdir & slashval(APPdir) & "Data\Display.dat" For Binary As Fringo
    Put Fringo, , Greet
    Close Fringo
    
    
    TempInt2 = 0
    TempInt = 1
        Do While TempInt <= NoDomains
            If Domain(TempInt).Status <> aDeleted Then
            TempInt2 = TempInt2 + 1
            fWriteValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt2, "NAME", "S", Domain(TempInt).Description
            fWriteValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt2, "PORT", "S", Domain(TempInt).Port
            fWriteValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt2, "HOME", "S", Right(Domain(TempInt).Home, Len(Domain(TempInt).Home) - Len(App.path & slashval(App.path) & "Domains\"))
            fWriteValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt2, "STATUS", "S", Domain(TempInt).Status
            fWriteValue APPdir & "Data\Config.ini", "DOMAIN" & TempInt2, "REFER", "S", Domain(TempInt).Refer

            CreatePath (Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data")
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Display.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Display.dat"
                DoEvents
                End If
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Script.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Script.dat"
                DoEvents
                End If
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest1.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest1.dat"
                DoEvents
                End If
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest2.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest2.dat"
                DoEvents
                End If
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest3.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest3.dat"
                DoEvents
                End If
                If Exist(Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest4.dat") = True Then
                Kill Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest4.dat"
                DoEvents
                End If
                
            Fringo = FreeFile
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Display.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).Greetings
            Close Fringo
            Fringo = FreeFile
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Script.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).Script
            Close Fringo
            Fringo = FreeFile
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest1.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).CMDrest1
            Close Fringo
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest2.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).CMDrest2
            Close Fringo
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest3.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).CMDrest3
            Close Fringo
            Open Domain(TempInt).Home & slashval(Domain(TempInt).Home) & "Data\Rest4.dat" For Binary As Fringo
            Put Fringo, , Domain(TempInt).CMDrest4
            Close Fringo

            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
    DoEvents
    fWriteValue APPdir & "Data\Config.ini", "SETUP", "DOMAINS", "S", TempInt2
    DoEvents
    End
    Else
    
    Cancel = True
    Master.Visible = False
    End If
endo:
End Sub


Private Sub Mod_Click()
Load Modules
Modules.Show
End Sub

Private Sub NoticeTim_Timer()
NOTICEexpirey = NOTICEexpirey + 1
SayTxt.FontBold = False
    If NOTICEexpirey >= 6 Then
    SayTxt.Text = ""
    NoticeTim.Enabled = False
    End If
End Sub

Private Sub Ref_Click()
Dim TempInt As Integer
TempInt = 1
    Do While TempInt <= NoDomains
    Domain(TempInt).Initialise
    TempInt = TempInt + 1
    DoEvents
    Loop
    
End Sub

Private Sub Reg_Click()
Load Register
Register.Show
End Sub

Private Sub shy_Click()
Master.Visible = True
Master.Show
End Sub

Private Sub Speed_Timer()
CurUp = 0
CurDown = 0
End Sub

Private Sub Sta_Click()
On Error GoTo endi
Dim TempInt As Integer
    If Sta.Checked = True Then
    WanListen.Close
    Say "Listening socket closed."
    Sta.Checked = False
    Else
        If WanListen.State <> sckListening Then
        WanListen.Close
        WanListen.LocalPort = WANport
        WanListen.Listen
        End If
    Say "Listening socket opened on port " & WANport & "."
    Sta.Checked = True
    End If
Exit Sub
endi:
MsgBox "There was an error listing for connections, the most likely reason may be a port conflict. Please double check that all of your domains are using unique ports, and that there are no other programs on your computer using any of these ports.", vbExclamation, "Error"

End Sub

Private Sub term_Click()
Dim TempInt As Integer
    Do While TempInt < WAN.Count
        If WAN(TempInt).State = sckConnected Then
        SendDat vbCrLf & ">SERVER INITIATED MASS TERMINATION.", TempInt
        End If
    WAN(TempInt).Close
    LAN(TempInt).Close
    TempInt = TempInt + 1
    DoEvents
    Loop
Sta.Checked = False
Say "All sockets were successfully closed."
End Sub

Private Sub Traf_Click()
Load Traffic
Traffic.Show
End Sub

Private Sub Use_Click()
Load Users
Users.Show
End Sub

Private Sub WAN_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim TmpStr As String
Dim InyStr As String
Dim TempInt As Integer
    If WAN(Index).State <> sckConnected Then
    Exit Sub
    End If
WAN(Index).GetData Data
    
    If BandMoniter = True Then
    CurDown = CurDown + Len(Data)
    TotDown = TotDown + Len(Data)
    End If
    
    If Left(Data, 1) = Chr(255) Then
    SendDat HandleCommands(Right(Data, 2)), Index
    DoEvents
    Exit Sub
    End If

    If (SessInfo(Index).Stage = Clear) And (LAN(Index).State <> sckConnected) Then
    Exit Sub
    End If
    
    If Len(Data) > 0 Then
        If Asc(Data) = 96 Then
            If Len(SessInfo(Index).PrevData) > Len(SessInfo(Index).Data) Then
            Data = Right(Left(SessInfo(Index).PrevData, Len(SessInfo(Index).Data) + 1), 1)
            Else
            Data = ""
            End If
        End If
    End If
    
    If Len(Data) > 0 Then
        If (Asc(Data) <> 8) And (Asc(Data) <> 127) Then
        SessInfo(Index).Data = SessInfo(Index).Data & Data
        End If
    Else
    Exit Sub
    End If


    If (Right(SessInfo(Index).Data, 1) = vbCr) Or (Right(SessInfo(Index).Data, 1) = vbLf) Then

        If (Right(SessInfo(Index).Data, 1) = vbLf) Then
        SessInfo(Index).Data = Left(SessInfo(Index).Data, Len(SessInfo(Index).Data) - 1)
        End If
        If (Right(SessInfo(Index).Data, 1) = vbCr) Then
        SessInfo(Index).Data = Left(SessInfo(Index).Data, Len(SessInfo(Index).Data) - 1)
        End If
    SessInfo(Index).PrevData = SessInfo(Index).Data
    InyStr = SessInfo(Index).Data
    SessInfo(Index).Data = ""

        
        If SessInfo(Index).Stage = Clear Then
            If LAN(Index).State = sckConnected Then
                If Len(InyStr) > 0 Then
                LAN(Index).SendData InyStr
                DoEvents
                Else
                LAN(Index).SendData vbCr
                DoEvents
                End If
            Else
            SendDat vbCrLf & "Not connected to domain." & vbCrLf, Index
            DoEvents
            End If
        End If
        
        If SessInfo(Index).Stage = Password Then
            If Number(InyStr) = True Then
                If (InyStr > 0) And (InyStr <= NoDomains) Then
                    If Domain(InyStr).Status = aAlive Then
                    SendDat vbCrLf & "Connecting:" & Domain(InyStr).Port & "..." & vbCrLf & vbCrLf, Index
                    DoEvents
                    LAN(Index).Connect "localhost", Domain(InyStr).Port
                    DoEvents
                    SessInfo(Index).Stage = Clear
                    Else
                    SendDat vbCrLf & "Domain: ", Index
                    DoEvents
                    End If
                Else
                SendDat vbCrLf & "Domain: ", Index
                DoEvents
                End If
            Else
            SendDat vbCrLf & "Domain: ", Index
            DoEvents
            End If
        End If
        
        If SessInfo(Index).Stage = Login Then
            If InyStr = AccessCode Then
            SendDat vbCrLf & "Access Granted." & vbCrLf, Index
            SendDat vbCrLf & "Select a domain...", Index
            TempInt = 1
                Do While TempInt <= NoDomains
                    If Domain(TempInt).Status = aAlive Then
                    SendDat vbCrLf & TempInt & "] " & Domain(TempInt).Description, Index
                    Else
                    SendDat vbCrLf & TempInt & "] " & Domain(TempInt).Description & " [DISABLED]", Index
                    End If
                TempInt = TempInt + 1
                DoEvents
                Loop
            SendDat vbCrLf & "Domain: ", Index
            DoEvents
            SessInfo(Index).Stage = Password
            Else
            SendDat vbCrLf & "Access Denied." & vbCrLf, Index
            WAN(Index).Close
            LAN(Index).Close
            End If
        End If
    
    Else
        If Len(Data) > 0 Then
            If (Asc(Data) = 8) Or (Asc(Data) = 127) Then
                If Len(SessInfo(Index).Data) > 0 Then
                SessInfo(Index).Data = Left(SessInfo(Index).Data, Len(SessInfo(Index).Data) - 1)
                SendDat Chr(8), Index
                DoEvents
                End If
            Else
                If SessInfo(Index).Stage = Login Then
                SendDat "*", Index
                DoEvents
                Else
                SendDat Data, Index
                DoEvents
                End If
            End If
        End If
    End If
End Sub

Private Sub WAN_Speed(Index As Integer, UPstream As Long, DOWNstream As Long)
CurUp = CurUp + UPstream
CurDown = CurDown + DOWNstream
    If TotDown < 1000000000 Then
    TotDown = TotDown + DOWNstream
    Else
    TotBreached = True
    TotDown = DOWNstream
    End If
    
    If TotUp < 1000000000 Then
    TotUp = TotUp + UPstream
    Else
    TotBreached = True
    TotUp = UPstream
    End If
End Sub

Private Sub WANlisten_ConnectionRequest(ByVal requestID As Long)
Dim Success As Boolean
Dim Indy As Integer

    If ClientList.ListItems.Count < 200 Then
    

    FreeWAN Indy, Success
        If Success = False Then
        Indy = WAN.Count
        Load WAN(Indy)
        Load LAN(Indy)
        End If
    LAN(Indy).Close
    WAN(Indy).Close
    WAN(Indy).Accept requestID
    DoEvents
        If WAN(Indy).State = sckConnected Then
        
        ClientList.ListItems.Add , "K" & Indy, WAN(Indy).RemoteHostIP, 1
        LogStr = LogStr & WAN(Indy).RemoteHostIP & " at " & Date & " [" & Time & "]" & vbCrLf
            If Len(LogStr) > 100 Then
            Dim Tempy As String
            Dim Fringo As Integer
            Tempy = LogStr
            LogStr = ""
            Fringo = FreeFile
            Open App.path & slashval(App.path) & "Log.txt" For Binary As Fringo
            Put Fringo, LOF(Fringo) + 1, Tempy
            Close Fringo
            End If
            
        SessInfo(Indy).IP = WAN(Indy).RemoteHostIP
        SessInfo(Indy).SignOn = Data & " [" & Time & "]"
        SessInfo(Indy).Data = ""
        
        SendComs Indy
        
            If Len(AccessCode) > 0 Then
            SendDat Greet & vbCrLf & "Access Code: ", Indy
            DoEvents
            SessInfo(Indy).Stage = Login
            Else
            Dim TempInt As Integer
            SendDat Greet & vbCrLf & vbCrLf & "Select a domain...", Indy
            DoEvents
            TempInt = 1
                Do While TempInt <= NoDomains
                    If Domain(TempInt).Status = aAlive Then
                    SendDat vbCrLf & TempInt & "] " & Domain(TempInt).Description, Indy
                    Else
                    SendDat vbCrLf & TempInt & "] " & Domain(TempInt).Description & " [DISABLED]", Indy
                    End If
                TempInt = TempInt + 1
                DoEvents
                Loop
            SendDat vbCrLf & "Domain: ", Indy
            DoEvents
            SessInfo(Indy).Stage = Password
            End If
            
            If TotConnections < 1000000000 Then
            TotConnections = TotConnections + 1
            Else
            TotBreached = True
            TotConnections = 1
            End If
        End If
    End If
End Sub

Private Function FreeWAN(ReturnIndex As Integer, Successfull As Boolean)
Dim TempInt As Integer
TempInt = 0
ReturnIndex = -1
    Do While TempInt < Master.WAN.Count
        If Master.WAN(TempInt).State <> sckConnected Then
        Master.WAN(TempInt).Close
        ReturnIndex = TempInt
        GoTo done
        End If
    TempInt = TempInt + 1
    Loop
done:
    If ReturnIndex >= 0 Then
    Successfull = True
    Else
    Successfull = False
    End If
End Function

Public Function SendDat(Data As String, Index As Integer)
On Error GoTo endi
    If (isLoaded(LogBox(Index)) = True) And (Len(Data) > 0) Then
    LogBox(Index).AddText Data
    End If
WAN(Index).SendData Data
DoEvents
    If BandMoniter = True Then
    CurUp = CurUp + Len(Data)
    TotUp = TotUp + Len(Data)
    End If
endi:
End Function
Public Function RefExists(Reference As String, Optional Index As Integer) As Boolean
Dim TempInt As Integer
RefExists = False
TempInt = 0
    Do While TempInt <= NoDomains
        If UCase(Trim(Reference)) = UCase(Trim(Master.Domain(TempInt).Refer)) Then
        RefExists = True
        Index = TempInt
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function

