VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OVR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object
Dim OverRide(1 To 600) As OverW
Dim Secs(1 To 600) As Integer
Dim OpenOver(0 To 600) As Integer
Dim isOpen(0 To 600) As Boolean
Private WithEvents Timera As CTimers
Attribute Timera.VB_VarHelpID = -1
Public Sub LocateMaster(Domainy As Object)
Set Domain = Domainy
End Sub
Private Sub SendReply(Message As String, Index As Integer)
Call Domain.reply(Index, Message)
End Sub

'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
'
'
'
'            MODIFY BELOW THIS LINE, NOT ABOVE!
'  (except for any variable arrays you may want to use)
'
'
'
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
Public Function Title() As String
'Enter a modules name here (preferably 3 letters long):
Title = "OVR"
Set Timera = New CTimers
    Do While Timera.Count < 600
    Timera.Add
    DoEvents
    Loop
End Function
Public Function Description() As String
'Enter a description of the module here:
Description = "Member of the Standard control library; the override (ovr)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to create override codes, and initiate them." & vbCrLf
Description = Description & "Command list: ovr.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
Dim TempInt As Integer
TempIn = 0
    Do While TempInt <= 600
    OverRide(TempInt).Alive(Index) = False
    DoEvents
    Loop
isOpen(Index) = False
OpenOver(Index) = -1
End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
    Dim Indexo As Integer

Dim TmpUser As UserInfoy
Dim Infor As BaseData     '  Creates a BaseData variable, and fills
Infor.Argument = Argument '  it with the appropraite data, BaseData is
Infor.CMD = CMD           '  required when you use some commands from the
Set Infor.Domain = Domain '  BASE module
Infor.Index = Index       '

Dim cUser As UserInfoy                '  Creates a variable called cUser that
cUser = GetUser(Infor, Index, eAlive) '  holds all of the data on your current user

Argument = Trim(Argument)  ' Trims the arguemnt, you may want to delete this

    'Help command
    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for OVR module...", Index
    SendReply vbCrLf & "INI [Code]         -Initiates an override sequence.", Index
    SendReply vbCrLf & "OPEN [Code]        -Opens an existing sequence, or if you specify", Index
    SendReply vbCrLf & "                    a new code, a new sequence wil be created.", Index
    SendReply vbCrLf & "EXP [Mins]         -Sets the opened sequence's expirey time in mins.", Index
    SendReply vbCrLf & "                    Activated once sequence is created/modified.", Index
    SendReply vbCrLf & "                    (-1 for unlimited)", Index
    SendReply vbCrLf & "LEVL [1][2][3][4]  -Sets the opened sequence's override level.", Index
    SendReply vbCrLf & "                    1=Guest", Index
    SendReply vbCrLf & "                    2=Standard", Index
    SendReply vbCrLf & "                    3=Service Admin", Index
    SendReply vbCrLf & "                    4=System Admin", Index
    SendReply vbCrLf & "QUAN [Amount]      -Sets the amount of times the opened sequence can", Index
    SendReply vbCrLf & "                    be initated before it expires (-1 for unlimited).", Index
    SendReply vbCrLf & "LIST               -Lists all outstanding override codes created", Index
    SendReply vbCrLf & "                    by you.", Index
    SendReply vbCrLf & "DEL                -Deletes the opened sequence.", Index
    Exit Function
    End If

    If UCase(CMD) = "INI" Then
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "You must specify an override code.", Index
        Exit Function
        End If
    Indexo = CodeExist(Argument)
        If Indexo >= 0 Then
            If ((OverRide(Indexo).Expirey = -1) Or (OverRide(Indexo).Expirey > 0)) And ((OverRide(Indexo).Remain > 0) Or (OverRide(Indexo).Remain = -1)) Then
                If cUser.Rights < OverRide(Indexo).Levely Then
                    If OverRide(Indexo).Remain > 0 Then
                    OverRide(Indexo).Remain = OverRide(Indexo).Remain - 1
                    End If
                TmpUser = cUser
                TmpUser.Rights = OverRide(Indexo).Levely
                SetAliveUser Infor, Index, TmpUser
                    If OverRide(Indexo).Expirey >= 0 Then
                    SendReply vbCrLf & vbCrLf & "Time: " & Time & ".", Index
                    If TmpUser.Rights = aGuest Then SendReply vbCrLf & "Your authority level has been overrided to Guest,", Index
                    If TmpUser.Rights = aStandard Then SendReply vbCrLf & "Your authority level has been overrided to Standard,", Index
                    If TmpUser.Rights = aServiceAdmin Then SendReply vbCrLf & "Your authority level has been overrided to Service Admin,", Index
                    If TmpUser.Rights = aSystemAdmin Then SendReply vbCrLf & "Your authority level has been overrided to System Admin,", Index
                    SendReply vbCrLf & "this will expire in " & OverRide(Indexo).Expirey & " minutes.", Index
                    Else
                    If TmpUser.Rights = aGuest Then SendReply vbCrLf & vbCrLf & "Your authority level has been overrided to Guest.", Index
                    If TmpUser.Rights = aStandard Then SendReply vbCrLf & vbCrLf & "Your authority level has been overrided to Standard.", Index
                    If TmpUser.Rights = aServiceAdmin Then SendReply vbCrLf & vbCrLf & "Your authority level has been overrided to Service Admin.", Index
                    If TmpUser.Rights = aSystemAdmin Then SendReply vbCrLf & vbCrLf & "Your authority level has been overrided to System Admin.", Index
                    End If
                OverRide(Indexo).Alive(Index) = True
                OverRide(Indexo).ResetUser(Index) = cUser.Rights
                    If OverRide(Indexo).Remain = 0 Then
                    OverRide(Indexo).Code = ""
                    End If
                Else
                SendReply vbCrLf & "You would not benefit from this override.", Index
                End If
            Else
            GoTo dead
            End If
        Else
        GoTo dead
        End If
    Exit Function
dead:
    SendReply vbCrLf & "Specified override sequence does not exist.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "DEL" Then
        If isOpen(Index) = False Then
        SendReply vbCrLf & "You must open a sequence first.", Index
        Exit Function
        End If
    OverRide(OpenOver(Index)).Code = ""
    OverRide(OpenOver(Index)).Expirey = 0
    OverRide(OpenOver(Index)).Remain = 0
    Timera(OpenOver(Index)).Disable
    SendReply vbCrLf & "Opened sequence deleted.", Index
    Exit Function
    End If

    
    If UCase(CMD) = "LIST" Then
    SendReply vbCrLf & vbCrLf & "Outstanding override codes...", Index
    SendReply vbCrLf & "CODE           : QUANTITY    : EXPIREY      : LEVEL", Index
    TempInt = 1
        Do While TempInt <= 600
            If (UCase(Trim(OverRide(TempInt).Owner)) = UCase(Trim(cUser.Username))) And (Len(OverRide(TempInt).Code) > 0) Then
            SendReply vbCrLf & OverRide(TempInt).Code & Space(14 - Len(CStr(OverRide(TempInt).Code))) & " : " & OverRide(TempInt).Remain & Space(11 - Len(CStr(OverRide(TempInt).Remain))) & " : " & OverRide(TempInt).Expirey & Space(12 - Len(CStr(OverRide(TempInt).Expirey))) & " : ", Index
            If OverRide(TempInt).Levely = Righty.aGuest Then SendReply "Guest", Index
            If OverRide(TempInt).Levely = Righty.aStandard Then SendReply "Standard", Index
            If OverRide(TempInt).Levely = Righty.aServiceAdmin Then SendReply "Service Admin", Index
            If OverRide(TempInt).Levely = Righty.aSystemAdmin Then SendReply "System Admin", Index
            End If
        DoEvents
        TempInt = TempInt + 1
        Loop
    Exit Function
    End If
    
    If UCase(CMD) = "LEVL" Then
        If isOpen(Index) = False Then
        SendReply vbCrLf & "You must open a sequence first.", Index
        Exit Function
        End If
        If Number(Argument) = False Then
        SendReply vbCrLf & "You must specify a level.", Index
        Exit Function
        End If
    TempInt = Round(Argument, 0)
        If (TempInt >= 1) And (TempInt <= 4) Then
            If (TempInt * 100) <= cUser.Rights Then
            OverRide(OpenOver(Index)).Levely = (TempInt * 100)
            SendReply vbCrLf & "Override level updated.", Index
            Else
            SendReply vbCrLf & "Access Denied.", Index
            End If
        Else
        SendReply vbCrLf & "Between 1-4.", Index
        End If
    Exit Function
    End If

    If UCase(CMD) = "QUAN" Then
        If isOpen(Index) = False Then
        SendReply vbCrLf & "You must open a sequence first.", Index
        Exit Function
        End If
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "You must specify a quantity.", Index
        Exit Function
        End If
        If Number(Argument) = False Then
        SendReply vbCrLf & "You must specify a quantity.", Index
        Exit Function
        End If
    Argument = Round(CInt(Argument), 0)
        If (CInt(Argument) < -1) Or (CInt(Argument) > 4000) Then
        SendReply vbCrLf & "Between -1 - 4000.", Index
        Exit Function
        End If
    OverRide(OpenOver(Index)).Remain = CInt(Argument)
    SendReply vbCrLf & "Override quantity updated.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "EXP" Then
        If isOpen(Index) = False Then
        SendReply vbCrLf & "You must open a sequence first.", Index
        Exit Function
        End If
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "You must specify a duration (in minutes).", Index
        Exit Function
        End If
        If Number(Argument) = False Then
        SendReply vbCrLf & "You must specify a duration (in minutes).", Index
        Exit Function
        End If
    Argument = Round(CInt(Argument), 0)
        If (CInt(Argument) < -1) Or (CInt(Argument) > 4320) Then
        SendReply vbCrLf & "Between -1 - 4320.", Index
        Exit Function
        End If
    OverRide(OpenOver(Index)).Expirey = CInt(Argument)
    SendReply vbCrLf & "Override expirey updated.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "OPEN" Then
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "You must specify an override code.", Index
        Exit Function
        End If
    Indexo = CodeExist(Argument)
        If Indexo = -1 Then
            If CountMe(Index) >= 3 Then
            SendReply vbCrLf & "You cannot create more than 3 override sequences.", Index
            Exit Function
            End If
        Indexo = FreeCode
            If Indexo >= 0 Then
            isOpen(Index) = True
            OpenOver(Index) = Indexo
            OverRide(Indexo).Code = Argument
            OverRide(Indexo).Levely = cUser.Rights
            OverRide(Indexo).Owner = cUser.Username
            Timera(Indexo).Interval = 10000
            Timera(Indexo).Enable
            Else
            SendReply vbCrLf & "Override limit reached.", Index
            Exit Function
            End If
        Else
            If UCase(Trim(OverRide(Indexo).Owner)) = UCase(Trim(cUser.Username)) Then
            isOpen(Index) = True
            OpenOver(Index) = Indexo
            Else
            SendReply vbCrLf & "Access Denied.", Index
            Exit Function
            End If
        End If
        
        If OverRide(Indexo).Expirey = 0 Then
        OverRide(Indexo).Expirey = 5
        End If
        If OverRide(Indexo).Remain = 0 Then
        OverRide(Indexo).Remain = 5
        End If
    SendReply vbCrLf & "Override sequence opened.", Index
    Exit Function
    End If
    
'SendReply sends a reply to the user!
'SendReply MESSAGE, INDEX

'Returns a message if the command could not be found:
SendReply vbCrLf & "Unknown command.", Index
End Function

Private Function CodeExist(Code As String) As Integer
Dim TempInt As Integer
CodeExist = -1
TempInt = 1
    Do While TempInt <= 600
        If Code = OverRide(TempInt).Code Then
        CodeExist = TempInt
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function
Private Function FreeCode() As Integer
Dim TempInt As Integer
FreeCode = -1
TempInt = 1
    Do While TempInt <= 600
        If Len(OverRide(TempInt).Code) = 0 Then
        FreeCode = TempInt
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function
Private Function CountMe(Index As Integer) As Integer
Dim TempInt As Integer
Dim cUser As UserInfoy
Dim Infor As BaseData
Infor.Argument = ""
Infor.CMD = ""
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, eAlive)

CountMe = 0
TempInt = 1
    Do While TempInt <= 600
        If (Len(OverRide(TempInt).Code) > 0) And (UCase(Trim(OverRide(TempInt).Owner)) = UCase(Trim(cUser.Username))) Then
        CountMe = CountMe + 1
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function

Private Sub Timera_Timer(ByVal Index As Integer)
    If OverRide(Index).Expirey > 0 Then
    Secs(Index) = Secs(Index) + 10
        If Secs(Index) >= 60 Then
        OverRide(Index).Expirey = OverRide(Index).Expirey - 1
            If OverRide(Index).Expirey = 0 Then
            Dim TempInt As Integer
            Dim TmpUser As UserInfoy
            Dim Infor As BaseData
            Infor.Argument = ""
            Infor.CMD = ""
            Set Infor.Domain = Domain
            Infor.Index = 0
            TempInt = 0
                Do While TempInt <= 600
                    If OverRide(Index).Alive(TempInt) = True Then
                    TmpUser = GetUser(Infor, TempInt, eAlive)
                    TmpUser.Rights = OverRide(Index).ResetUser(TempInt)
                    Domain.CD TempInt, Domain.Home & slashval(Domain.Home) & "Users\" & TmpUser.Username
                    SetAliveUser Infor, TempInt, TmpUser
                    Call Domain.reply(TempInt, vbCrLf & vbCrLf & "Override has expired.")
                    Call Domain.sendline(TempInt, True)
                    End If
                OverRide(Index).Alive(TempInt) = False
                TempInt = TempInt + 1
                DoEvents
                Loop
            OverRide(Index).Code = ""
            Timera(Index).Disable
            End If
        Secs(Index) = 0
        End If
    End If
End Sub
