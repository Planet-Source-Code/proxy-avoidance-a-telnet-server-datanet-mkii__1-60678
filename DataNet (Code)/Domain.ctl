VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl Domain 
   BackColor       =   &H00C00000&
   ClientHeight    =   3360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   3360
   ScaleWidth      =   3180
   Begin DataNet.PortStatus PortStatus 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
   End
   Begin MSWinsockLib.Winsock WAN 
      Index           =   0
      Left            =   840
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WanListen 
      Left            =   1800
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Domain"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Domain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Port = 0
Const m_def_CMDrest = ""
Const m_def_TotalCon = 0
Const m_def_SockInitiated = 0
Const m_def_Greetings = "/DataNet MKII Server"
Const m_def_Status = 200
Const m_def_Home = "C:\"
Const m_def_Script = ""
Const m_def_Refer = ""
Const m_def_Description = "DataNet Standard Domain"
'Property Variables:
Dim m_Port As Integer
Dim m_TotalCon As Long
Dim m_Greetings As String
Dim m_Home As String
Dim m_Script As String
Dim m_Refer As String
Dim m_CMDrest1 As String
Dim m_CMDrest2 As String
Dim m_CMDrest3 As String
Dim m_CMDrest4 As String
Dim m_Description As String
Dim m_Status As eStatus
Public Enum aStatus
    aListen = 100
    aClose = 200
    aTerminate = 300
End Enum
Public Enum eStatus
    aAlive = 100
    aDisabled = 200
    aDeleted = 300
End Enum
Dim AliveUser(0 To 1000) As UserInfoy
Dim ArchiveUser(0 To 1000) As UserInfoy
Dim LimboUser(0 To 1000) As UserInfoy
Dim UserNo As Integer
Private Type Mody
    Namey As String
    Obj As Object
    Description As String
End Type
Private Type UserInfoy
    UserName As String
    Realname As String
    Password As String
    Rights As Righty
    Stage As Stagey
    Status As eStatus
End Type
Private Enum Stagey
    Login = 100
    Password = 200
    Clear = 300
End Enum
Private Enum Righty
    aSystemAdmin = 400
    aServiceAdmin = 300
    aStandard = 200
    aGuest = 100
End Enum

Dim LogFi(0 To 1000) As Integer
Public Event NewConnection(IP As String)
Public Event Message(Message As String)
Dim CMDline(0 To 1000) As String

Dim UsrPath(0 To 1000) As String
Dim Mody(0 To 10000) As Mody
Dim ModsLoaded As Integer
Dim adoptCom(0 To 1000) As String
Dim adoptArg(0 To 1000) As String
Dim AccountLink(0 To 1000) As Integer

Public Property Get Port() As Integer
    Port = m_Port
End Property
Public Property Let Port(ByVal New_Port As Integer)
    m_Port = New_Port
    PropertyChanged "Port"
End Property
Public Property Get Refer() As String
    Refer = m_Refer
End Property
Public Property Let Refer(ByVal New_Refer As String)
    m_Refer = New_Refer
    PropertyChanged "Refer"
End Property
Public Property Get CMDrest1() As String
    CMDrest1 = m_CMDrest1
End Property
Public Property Let CMDrest1(ByVal New_CMDrest1 As String)
    m_CMDrest1 = New_CMDrest1
    PropertyChanged "CMDrest1"
End Property
Public Property Get CMDrest2() As String
    CMDrest2 = m_CMDrest2
End Property
Public Property Let CMDrest2(ByVal New_CMDrest2 As String)
    m_CMDrest2 = New_CMDrest2
    PropertyChanged "CMDrest2"
End Property
Public Property Get CMDrest3() As String
    CMDrest3 = m_CMDrest3
End Property
Public Property Let CMDrest3(ByVal New_CMDrest3 As String)
    m_CMDrest3 = New_CMDrest3
    PropertyChanged "CMDrest3"
End Property
Public Property Get CMDrest4() As String
    CMDrest4 = m_CMDrest4
End Property
Public Property Let CMDrest4(ByVal New_CMDrest4 As String)
    m_CMDrest4 = New_CMDrest4
    PropertyChanged "CMDrest4"
End Property
Public Property Get Script() As String
    Script = m_Script
End Property
Public Property Let Script(ByVal New_Script As String)
    m_Script = New_Script
    PropertyChanged "Script"
End Property
Public Property Get Status() As eStatus
    Status = m_Status
End Property
Public Property Let Status(ByVal New_Status As eStatus)
    m_Status = New_Status
    PropertyChanged "Status"
End Property
Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Let Description(ByVal New_Description As String)
    m_Description = New_Description
    PropertyChanged "Description"
End Property
Public Property Get Home() As String
    Home = m_Home
End Property
Public Property Let Home(ByVal New_Home As String)
    m_Home = New_Home
    PropertyChanged "Home"
End Property
Public Property Get TotalCon() As Long
    TotalCon = m_TotalCon
End Property
Public Property Get SockInitiated() As Long
    SockInitiated = WAN.Count
End Property
Public Property Get Greetings() As String
    Greetings = m_Greetings
End Property
Public Property Let Greetings(ByVal New_Greetings As String)
    m_Greetings = New_Greetings
    PropertyChanged "Greetings"
End Property
Private Sub UserControl_InitProperties()
    m_Port = m_def_Port
    m_Greetings = m_def_Greetings
    m_TotalCon = m_def_TotalCon
    m_Home = m_def_Home
    m_Description = m_def_Description
    m_Status = m_def_Status
    m_Script = m_def_Script
    m_Refer = m_def_Refer
    m_CMDrest1 = m_def_CMDrest
    m_CMDrest2 = m_def_CMDrest
    m_CMDrest3 = m_def_CMDrest
    m_CMDrest4 = m_def_CMDrest
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Port = PropBag.ReadProperty("Port", m_def_Port)
    m_Greetings = PropBag.ReadProperty("Greetings", m_def_Greetings)
    m_Home = PropBag.ReadProperty("Home", m_def_Home)
    m_Description = PropBag.ReadProperty("Description", m_def_Description)
    m_Status = PropBag.ReadProperty("Status", m_def_Status)
    m_Script = PropBag.ReadProperty("Script", m_def_Script)
    m_Refer = PropBag.ReadProperty("Refer", m_def_Refer)
    m_CMDrest1 = PropBag.ReadProperty("CMDrest1", m_def_CMDrest)
    m_CMDrest2 = PropBag.ReadProperty("CMDrest2", m_def_CMDrest)
    m_CMDrest3 = PropBag.ReadProperty("CMDrest3", m_def_CMDrest)
    m_CMDrest4 = PropBag.ReadProperty("CMDrest4", m_def_CMDrest)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Port", m_Port, m_def_Port)
    Call PropBag.WriteProperty("Greetings", m_Greetings, m_def_Greetings)
    Call PropBag.WriteProperty("Home", m_Home, m_def_Home)
    Call PropBag.WriteProperty("Description", m_Description, m_def_Description)
    Call PropBag.WriteProperty("Status", m_Status, m_def_Status)
    Call PropBag.WriteProperty("Script", m_Script, m_def_Script)
    Call PropBag.WriteProperty("Refer", m_Refer, m_def_Refer)
    Call PropBag.WriteProperty("CMDrest1", m_CMDrest1, m_def_CMDrest)
    Call PropBag.WriteProperty("CMDrest2", m_CMDrest2, m_def_CMDrest)
    Call PropBag.WriteProperty("CMDrest3", m_CMDrest3, m_def_CMDrest)
    Call PropBag.WriteProperty("CMDrest4", m_CMDrest4, m_def_CMDrest)
End Sub
Private Function FreeWAN(ReturnIndex As Integer, Successfull As Boolean)
Dim TempInt As Integer
TempInt = 0
ReturnIndex = -1
    Do While TempInt < WAN.Count
        If WAN(TempInt).State <> sckConnected Then
        WAN(TempInt).Close
        ReturnIndex = TempInt
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
    
    If ReturnIndex >= 0 Then
    Successfull = True
    Else
    Successfull = False
    End If
End Function
Private Function Exist(strPath As String) As Boolean
Dim TempInt As Long
    TempInt = (Dir(strPath) = "")
    
If TempInt = 0 Then
Exist = True
Else
Exist = False
End If
End Function
Private Function slashval(path As String) As String
    If Right(path, 1) = "\" Then
    slashval = ""
    Else
    slashval = "\"
    End If
End Function
Private Function slashvalB(path As String) As String
    If Left(path, 1) = "\" Then
    slashvalB = ""
    Else
    slashvalB = "\"
    End If
End Function
Private Function DELslash(path As String) As String
    If Right(path, 1) = "\" Then
    DELslash = Left(path, Len(path) - 1)
    Else
    DELslash = path
    End If
End Function
Private Function DELslashB(path As String) As String
    If Left(path, 1) = "\" Then
    DELslashB = Right(path, Len(path) - 1)
    Else
    DELslashB = path
    End If
End Function
Private Function IfExist(path As String) As Boolean
Temp = CurDir         '
On Error GoTo endi    '
                      '
ChDir path            '
IfExist = True        '
ChDir Temp            '  Checks to see if the specified path exists,
Exit Function         '  and returns it as a boolean (true or false)
                      '
endi:                 '
IfExist = False       '
ChDir Temp            '
End Function
Private Function pFormatSize(ByVal dSize As Double) As String

' 1024  b = 1 kb: 1024 kb = 1 mb

    If dSize < 1024 Then
        pFormatSize = dSize & " bytes"
    Else
        dSize = dSize / 1024
        If dSize < 1000 Then
            pFormatSize = Format$(dSize, "#,##0.0") & " kb"
        Else
            pFormatSize = Format$(dSize / 1024, "#,##0.0") & " mb"
        End If
    End If
    
End Function
Private Function FolderFromPath(FilePath As String) As String
Dim Tempyo As String
Dim Lengy As Long
Lengy = 0
    Do While (Left(Tempyo, 1) = "\") = False
    Lengy = Lengy + 1
    Tempyo = Right(FilePath, Lengy)
        If Lengy > 100 Then
        FolderFromPath = "/NO FOLDER>"
        Exit Function
        End If
    DoEvents
    Loop

FolderFromPath = Left(FilePath, Len(FilePath) - Len(Tempyo))
End Function
Private Function FileFromPath(FilePath As String) As String
Dim Tempyo As String
Dim Lengy As Long
Lengy = 0
    Do While (Left(Tempyo, 1) = "\") = False
    If Lengy > 100 Then GoTo endio
    Lengy = Lengy + 1
    Tempyo = Right(FilePath, Lengy)
    DoEvents
    Loop
FileFromPath = Right(Tempyo, Len(Tempyo) - 1)

Exit Function
endio:
FileFromPath = ""
End Function
Private Sub UserControl_Resize()
UserControl.Width = 990
UserControl.Height = 465
End Sub
Private Sub UserControl_Initialize()
Call UserControl_Resize
End Sub

'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
'--------------------------------
Public Sub IniMods()
CreatePath (m_Home & slashval(m_Home) & "Data")
fReadValue m_Home & slashval(m_Home) & "Data\modules.ini", "INFO", "MODULES", "S", "0", ModsLoaded

TempInt = 0
    Do While TempInt < ModsLoaded
    LoadMod (TempInt)
    TempInt = TempInt + 1
    DoEvents
    Loop
End Sub
Public Sub IniUser()
Dim TempInt As Integer
Dim TempInt2 As Integer
Dim Fringo As Integer
Dim Pathy As String
Dim tmpPathy As String
CreatePath (m_Home & slashval(m_Home) & "Data")
tmpPathy = m_Home & slashval(m_Home) & "Data\Users.tmp"
Pathy = m_Home & slashval(m_Home) & "Data\users.ini"
fReadValue Pathy, "SETUP", "USERSNO", "S", "0", UserNo
TempInt = 0
    Do While TempInt < UserNo
    fReadValue Pathy, "USER" & TempInt, "USER", "S", "Guest", ArchiveUser(TempInt).UserName
    fReadValue Pathy, "USER" & TempInt, "NAME", "S", "Guest", ArchiveUser(TempInt).Realname
    fReadValue Pathy, "USER" & TempInt, "PASS", "S", "", ArchiveUser(TempInt).Password
    fReadValue Pathy, "USER" & TempInt, "RIGHTS", "S", "100", ArchiveUser(TempInt).Rights
    fReadValue Pathy, "USER" & TempInt, "STATUS", "S", "200", ArchiveUser(TempInt).Status
    ArchiveUser(TempInt).Stage = Clear
    TempInt = TempInt + 1
    DoEvents
    Loop

    If Exist(tmpPathy) = True Then
    Kill tmpPathy
    DoEvents
    End If
Fringo = FreeFile
Open tmpPathy For Binary As Fringo
Close Fringo
DoEvents
TempInt2 = 0
TempInt = 0
    Do While TempInt < UserNo
        If ArchiveUser(TempInt).Status <> aDeleted Then
        fWriteValue tmpPathy, "USER" & TempInt2, "USER", "S", ArchiveUser(TempInt).UserName
        fWriteValue tmpPathy, "USER" & TempInt2, "NAME", "S", ArchiveUser(TempInt).Realname
        fWriteValue tmpPathy, "USER" & TempInt2, "PASS", "S", ArchiveUser(TempInt).Password
        fWriteValue tmpPathy, "USER" & TempInt2, "STATUS", "S", ArchiveUser(TempInt).Status
        fWriteValue tmpPathy, "USER" & TempInt2, "RIGHTS", "S", ArchiveUser(TempInt).Rights
        TempInt2 = TempInt2 + 1
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
fWriteValue tmpPathy, "SETUP", "USERSNO", "S", TempInt2
    If Exist(tmpPathy) = True Then
    DoEvents
        If Exist(Pathy) = True Then
        Kill Pathy
        DoEvents
        End If
    FileCopy tmpPathy, Pathy
    DoEvents
    End If

fReadValue Pathy, "SETUP", "USERSNO", "S", "0", UserNo
TempInt = 0
    Do While TempInt < UserNo
    fReadValue Pathy, "USER" & TempInt, "USER", "S", "Guest", ArchiveUser(TempInt).UserName
    fReadValue Pathy, "USER" & TempInt, "NAME", "S", "Guest", ArchiveUser(TempInt).Realname
    fReadValue Pathy, "USER" & TempInt, "PASS", "S", "", ArchiveUser(TempInt).Password
    fReadValue Pathy, "USER" & TempInt, "RIGHTS", "S", "100", ArchiveUser(TempInt).Rights
    fReadValue Pathy, "USER" & TempInt, "STATUS", "S", "200", ArchiveUser(TempInt).Status
    ArchiveUser(TempInt).Stage = Clear
    TempInt = TempInt + 1
    DoEvents
    Loop
End Sub

Public Sub Initialise()
Call IniMods
Call IniUser
End Sub


Public Sub Action(Action As aStatus)
Dim TempInt As Integer
On Error GoTo endi
    If Action = aListen Then
    Dim Succy As Boolean
        If WanListen.State <> sckListening Then
        WanListen.Close
        DoEvents
            If Len(Trim(PortStatus.ApplicationUsingPort(m_Port, TCP))) > 0 Then
            RaiseEvent Message(Trim(PortStatus.ApplicationUsingPort(m_Port, TCP)) & " is using a port required by DataNet")
            Exit Sub
            End If
        WanListen.Bind m_Port, "localhost"
        WanListen.Listen
        End If
    RaiseEvent Message("Listening.")
    End If
    
    If Action = aClose Then
    WanListen.Close
    RaiseEvent Message("Listening socket closed.")
    End If

    If Action = aTerminate Then
        Do While TempInt < WAN.Count
        SendIt vbCrLf & ">SERVER INITIATED MASS TERMINATION.", TempInt
        DoEvents
        WAN(TempInt).Close
        TempInt = TempInt + 1
        DoEvents
        Loop
    RaiseEvent Message("All sockets were successfully closed.")
    End If
Exit Sub
endi:
RaiseEvent Message("ERROR")
End Sub

Private Sub WAN_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String
Dim TempInt As Integer
Dim Coms() As String
Dim TimeDec As String
WAN(Index).GetData Data
    If Data = vbCr Then
    Data = ""
    End If
    
    If LogFi(Index) >= 0 Then Put LogFi(Index), LOF(LogFi(Index)) + 1, Data
    
    If LimboUser(Index).Stage = Login Then
    LimboUser(Index).UserName = Data
    SendIt vbCrLf & "Password: ", Index
    DoEvents
    LimboUser(Index).Stage = Password
    Exit Sub
    End If

    If LimboUser(Index).Stage = Password Then
    Dim Indexy As Integer
    LimboUser(Index).Password = Data
        If HasAccess(LimboUser(Index), Indexy) = True Then
        AliveUser(Index) = ArchiveUser(Indexy)
        SendIt vbCrLf & "Access Granted." & vbCrLf, Index
        DoEvents

        CreatePath m_Home & slashval(m_Home) & "logs\" & AliveUser(Index).UserName
        TimeDec = Day(Date) & "-" & Month(Date) & "-" & Right(Year(Date), 2)
        TimeDec = TimeDec & " [" & Hour(Time) & Minute(Time) & "]"
        On Error GoTo cont1
        Close LogFi(Index)
cont1:

        LogFi(Index) = FreeFile
        Open m_Home & slashval(m_Home) & "logs\" & AliveUser(Index).UserName & "\" & TimeDec & "_" & Round(Rnd * 1000, 0) & ".txt" For Binary As LogFi(Index)
        UsrPath(Index) = m_Home & slashval(m_Home) & "Users\" & AliveUser(Index).UserName
        AccountLink(Index) = Indexy
        CMDline(Index) = FormatPath(AliveUser(Index), UsrPath(Index))
        Coms() = Split(m_Script, vbCrLf)
            Do While TempInt <= UBound(Coms())
                If Len(Coms(TempInt)) > 0 Then
                SendIt cmd(Coms(TempInt), Index), Index
                Else
                    If Len(adoptArg(Index)) > 0 Then
                    SendIt cmd(Coms(TempInt), Index), Index
                    End If
                End If
            TempInt = TempInt + 1
            DoEvents
            Loop
            
        RaiseEvent Message(UCase(AliveUser(Index).UserName) & " logged in.")
        LimboUser(Index).Stage = Clear
        SendLine Index, True
        Else
        SendIt vbCrLf & "Access Denied." & vbCrLf, Index
        WAN(Index).Close
        DoEvents
        End If
    Exit Sub
    End If

    If LimboUser(Index).Stage = Clear Then
    Dim Executed As Boolean
        If Len(Data) > 0 Then
        Executed = True
        SendIt cmd(Data, Index), Index
        Else
            If Len(adoptArg(Index)) > 0 Then
            Executed = True
            SendIt cmd(Data, Index), Index
            End If
        End If
        
    SendLine Index, Executed
    Exit Sub
    End If
End Sub


Private Function HasAccess(User As UserInfoy, Index As Integer) As Boolean
Dim TempInt As Integer
HasAccess = False
    Do While TempInt < UserNo
        If (UCase(User.UserName) = UCase(ArchiveUser(TempInt).UserName)) And (User.Password = ArchiveUser(TempInt).Password) Then
            If ArchiveUser(TempInt).Status = aAlive Then
            Index = TempInt
            HasAccess = True
            Else
            HasAccess = False
            End If
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function

Public Sub SendLine(Index As Integer, Executed As Boolean)
    If Executed = True Then
    SendIt vbCrLf & vbCrLf & CMDline(Index) & ">" & Chr(231) & Chr(12), Index
    Else
    SendIt vbCrLf & CMDline(Index) & ">" & Chr(231) & Chr(12), Index
    End If
DoEvents
End Sub
Private Function FullPath(path As String) As String
FullPath = m_Home & slashval(m_Home) & "Users\" & path
End Function

Public Function cmd(Command As String, Index As Integer) As String
Dim Splity() As String
Dim Part() As String
Dim Comm As String
Dim Argu As String
Dim TempInt As Integer
Dim Restr() As String
Dim TEMPcom As String
If AliveUser(Index).Rights = aSystemAdmin Then Restr = Split(m_CMDrest4, vbCrLf)
If AliveUser(Index).Rights = aServiceAdmin Then Restr = Split(m_CMDrest3, vbCrLf)
If AliveUser(Index).Rights = aStandard Then Restr = Split(m_CMDrest2, vbCrLf)
If AliveUser(Index).Rights = aGuest Then Restr = Split(m_CMDrest1, vbCrLf)
    If Len(adoptArg(Index)) > 0 Then
    TEMPcom = Trim(adoptArg(Index)) & " " & Command
    Else
        If Len(adoptCom(Index)) > 0 Then
        TEMPcom = Trim(adoptCom(Index)) & "." & Command
        Else
        TEMPcom = Command
        End If
    End If
    
TempInt = 0
    Do While TempInt <= UBound(Restr)
        If Len(Replace(Trim(UCase(TEMPcom)), Trim(UCase(Restr(TempInt))), "")) <> Len(Trim(UCase(TEMPcom))) Then
        cmd = vbCrLf & "Command restricted."
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
    
    If ((Trim(UCase(Command)) = "HELP") And (Len(adoptCom(Index)) = 0) And (Len(adoptArg(Index)) = 0)) Then
    cmd = vbCrLf
    cmd = cmd & vbCrLf
    cmd = cmd & "DataNet MKII Telnet Server Help" & vbCrLf
    cmd = cmd & vbCrLf
    cmd = cmd & " Long Command Syntax: MODULE.COMMAND ARGUMENT" & vbCrLf
    cmd = cmd & "Short Command Syntax: COMMAND ARGUMENT" & vbCrLf
    cmd = cmd & "NOTE: To use the short command syntax, you must first 'adopt'" & vbCrLf
    cmd = cmd & "      a specific module, to adopt a module, you must use the" & vbCrLf
    cmd = cmd & "      command syntax: MOD:MODULE" & vbCrLf
    cmd = cmd & vbCrLf
    cmd = cmd & "EXAMPLE:" & vbCrLf
    cmd = cmd & "test>SCR.CLS        <- Long syntax for a 'scr' command" & vbCrLf
    cmd = cmd & "test>MOD:SCR        <- Adopts the 'scr' module, now the short" & vbCrLf
    cmd = cmd & "                       syntax can be used with this module" & vbCrLf
    cmd = cmd & "test>CLS            <- Short syntax for a 'scr' command" & vbCrLf
    cmd = cmd & "test>ECHO Hello!    <- Short syntax for a 'scr' command" & vbCrLf
    cmd = cmd & "                       (including an argument)" & vbCrLf
    cmd = cmd & vbCrLf
    cmd = cmd & "To view the list of modules that this domain has to offer, enter" & vbCrLf
    cmd = cmd & "the command MODULES"
    Exit Function
    End If

    If (Trim(UCase(Command)) = "MODULES") And (Len(adoptArg(Index)) = 0) Then
    cmd = vbCrLf
    cmd = cmd & vbCrLf
    cmd = cmd & "Loaded module listing..." & vbCrLf
    cmd = cmd & vbCrLf
    TempInt = 0
        Do While TempInt < ModsLoaded
        cmd = cmd & Mody(TempInt).Namey & ":" & vbCrLf
        cmd = cmd & Mody(TempInt).Description & "---" & vbCrLf
        TempInt = TempInt + 1
        DoEvents
        Loop
    cmd = cmd & "Module listing complete."
    Exit Function
    End If

    If Trim(UCase(Left(Command, 4))) = "MOD:" Then
        If Len(Trim(Right(Command, Len(Command) - 4))) > 0 Then
        adoptCom(Index) = Trim(Right(Command, Len(Command) - 4))
        cmd = vbCrLf & "'" & Right(Command, Len(Command) - 4) & "' module adopted."
        Else
        adoptCom(Index) = ""
        cmd = vbCrLf & "No module adopted."
        End If
    Exit Function
    End If
    If Trim(UCase(Left(Command, 4))) = "COM:" Then
        If Len(Trim(Right(Command, Len(Command) - 4))) > 0 Then
        adoptArg(Index) = Trim(Right(Command, Len(Command) - 4))
        cmd = vbCrLf & "'" & Right(Command, Len(Command) - 4) & "' command adopted."
        Else
        adoptArg(Index) = ""
        cmd = vbCrLf & "No command adopted."
        End If
    Exit Function
    End If
    
Command = Command & " "
    If ComValid(Command, Index) = True Then
        Do While TempInt <= Len(Command)
            If Right(Left(Command, TempInt), 1) = " " Then
            Comm = Trim(Left(Command, TempInt - 1))
            Argu = Right(Command, Len(Command) - TempInt)
            GoTo found
            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
found:
If Right(Argu, 1) = " " Then Argu = Left(Argu, Len(Argu) - 1)
    Part = Split(Comm, ".")
    TempInt = 0
        Do While TempInt < ModsLoaded
            If UCase(Part(0)) = UCase(Mody(TempInt).Namey) Then
            On Error GoTo end1
            Mody(TempInt).Obj.RunCMD Right(Comm, Len(Comm) - (Len(Part(0)) + 1)), Argu, Index
            Exit Function
end1:
            cmd = vbCrLf & "Module not responding."
            Exit Function
            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
    cmd = vbCrLf & "Unknown command."
    Else
    cmd = vbCrLf & "Invalid command."
    End If


End Function

Private Function ComValid(Command As String, Index As Integer) As Boolean
Dim Stops As Integer
Dim Colons As Integer
Dim TempInt As Integer
ComValid = False
    
    If Len(Trim(adoptArg(Index))) > 0 Then
    Command = Trim(adoptArg(Index)) & " " & Command
    End If
    
    Do While TempInt <= Len(Command)
        If Right(Left(Command, TempInt), 1) = "." Then
        Stops = Stops + 1
        End If
        If Right(Left(Command, TempInt), 1) = " " Then
            If Stops = 0 Then
                If Len(adoptCom(Index)) > 0 Then
                Command = adoptCom(Index) & "." & Command
                ComValid = True
                Else
                ComValid = False
                End If
            Exit Function
            End If
            
            If Stops = 1 Then
            ComValid = True
            Exit Function
            End If
            
            If Stops > 1 Then
                If Len(adoptCom(Index)) > 0 Then
                Command = adoptCom(Index) & "." & Command
                ComValid = True
                Else
                ComValid = False
                End If
            Exit Function
            End If
        End If
    
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function

Private Sub LoadMod(Index As Integer)
On Error GoTo endi
Dim TempStr As String
CreatePath (m_Home & slashval(m_Home) & "Data")
fReadValue m_Home & slashval(m_Home) & "Data\modules.ini", "MODULES", "MOD" & Index, "S", "", TempStr
Set Mody(Index).Obj = CreateObject(TempStr)
Mody(Index).Description = Mody(Index).Obj.Description
Mody(Index).Namey = Mody(Index).Obj.Title
Mody(Index).Obj.LocateMaster Me
Mody(Index).Obj.LocateMain Master
endi:
DoEvents
End Sub

'Functions aimed at modules usage

Public Sub Reply(Index As Integer, Message As String)
SendIt Message, Index
DoEvents
End Sub
Public Sub ChLine(Index As Integer, Line As String)
CMDline(Index) = Line
DoEvents
End Sub
Public Sub gAliveUser(Index As Integer, UserName As String, Realname As String, Password As String, Rights As Integer, Stage As Integer, Status As Integer)
UserName = AliveUser(Index).UserName
Realname = AliveUser(Index).Realname
Password = AliveUser(Index).Password
Rights = AliveUser(Index).Rights
Stage = AliveUser(Index).Stage
Status = AliveUser(Index).Status
DoEvents
End Sub
Public Sub sAliveUser(Index As Integer, UserName As String, Realname As String, Password As String, Rights As Integer, Stage As Integer, Status As Integer)
AliveUser(Index).UserName = UserName
AliveUser(Index).Realname = Realname
AliveUser(Index).Password = Password
AliveUser(Index).Rights = Rights
AliveUser(Index).Stage = Stage
AliveUser(Index).Status = Status
DoEvents
End Sub

Public Sub gArchiveUser(Index As Integer, UserName As String, Realname As String, Password As String, Rights As Integer, Stage As Integer, Status As Integer)
UserName = ArchiveUser(Index).UserName
Realname = ArchiveUser(Index).Realname
Password = ArchiveUser(Index).Password
Rights = ArchiveUser(Index).Rights
Stage = ArchiveUser(Index).Stage
Status = ArchiveUser(Index).Status
DoEvents
End Sub
Public Function gAccountLink(Index As Integer) As Integer
gAccountLink = AccountLink(Index)
End Function
Public Function gUserNo() As Integer
gUserNo = UserNo
End Function
Public Function CD(Index As Integer, Optional NewPath As String = "") As String
    If Len(NewPath) > 0 Then
    UsrPath(Index) = NewPath
    CD = UsrPath(Index)
    Else
    CD = UsrPath(Index)
    End If
End Function
Public Function gAdopted(Index As Integer, Extend As Boolean, Optional NewAdopt As String = "") As String
    If Len(NewAdopt) > 0 Then
        If Extend = True Then
        adoptArg(Index) = NewAdopt
        gAdopted = adoptArg(Index)
        Else
        adoptCom(Index) = NewAdopt
        gAdopted = adoptCom(Index)
        End If
    Else
        If Extend = True Then
        gAdopted = adoptArg(Index)
        Else
        gAdopted = adoptCom(Index)
        End If
    End If
End Function

Public Function gSocket(Index As Integer) As Winsock
Set gSocket = WAN(Index)
End Function
Public Function gSocketCount() As Integer
gSocketCount = WAN.Count
End Function

Public Function UserExists(UserName As String) As Boolean
Dim TempInt As Integer
UserExists = False
TempInt = 0
    Do While TempInt < UserNo
        If UCase(Trim(UserName)) = UCase(Trim(ArchiveUser(TempInt).UserName)) Then
        UserExists = True
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function
Private Sub SendIt(Data As String, Index As Integer)
On Error GoTo endi
If LogFi(Index) >= 0 Then Put LogFi(Index), LOF(LogFi(Index)) + 1, Data
    If WAN(Index).State = sckConnected Then
    WAN(Index).SendData Data
    DoEvents
    End If
endi:
End Sub

Private Sub WANlisten_ConnectionRequest(ByVal requestID As Long)
Dim Success As Boolean
Dim Indy As Integer
Dim TempInt As Integer
FreeWAN Indy, Success
    If Success = False Then
    Indy = WAN.Count
    Load WAN(Indy)
    End If
WAN(Indy).Close
WAN(Indy).Accept requestID

LogFi(Indy) = -1
LimboUser(Indy).Stage = Login
adoptCom(Indy) = ""
adoptArg(Indy) = ""
TempInt = 0
    Do While TempInt < ModsLoaded
    ClearMod TempInt
    TempInt = TempInt + 1
    DoEvents
    Loop

RaiseEvent NewConnection(WAN(Indy).RemoteHostIP)

SendIt Chr(27) & "[2J", Indy
DoEvents
SendIt Chr(27) & "[H", Indy
DoEvents
SendIt m_Greetings & vbCrLf & "Login: ", Indy
DoEvents

    If m_TotalCon < 1000000000 Then
    m_TotalCon = m_TotalCon + 1
    Else
    m_TotalCon = 1
    End If
End Sub

Private Function ClearMod(Index As Integer)
On Error GoTo endi
Mody(Index).Obj.Clear Indy
endi:
End Function

Private Function FormatPath(cUser As UserInfoy, path As String) As String
Dim Starter As String

    If cUser.Rights = aSystemAdmin Then
    FormatPath = UCase(DELslash(path))
    Exit Function
    End If
    
    If cUser.Rights = aServiceAdmin Then
    Starter = FileFromPath(slashvalB(DELslash(m_Home)) & DELslash(m_Home))
    FormatPath = UCase(DELslash(Starter & slashval(Starter) & DELslashB(Right(path, Len(path) - Len(DELslash(m_Home))))))
    Exit Function
    End If
    
    If cUser.Rights <= aStandard Then
    Starter = cUser.UserName
    FormatPath = UCase(DELslash(Starter & slashval(Starter) & DELslashB(Right(path, Len(path) - Len(DELslash(m_Home & slashval(m_Home) & "Users\" & cUser.UserName))))))
    End If
End Function
