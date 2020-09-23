VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "USR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object

Private OpenUser(0 To 1000) As UserInfoy
Private OpenUserInt(0 To 1000) As Integer

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
'
'
'
'
'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
Public Function Title() As String
Title = "USR"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the user (usr)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to manipulate user accounts within the domain." & vbCrLf
Description = Description & "Command list: usr.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
OpenUser(Index).Password = ""
OpenUser(Index).Realname = ""
OpenUser(Index).Rights = aGuest
OpenUser(Index).Stage = Login
OpenUser(Index).Status = aDisabled
OpenUser(Index).Username = ""
OpenUserInt(Index) = 0
End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim TmpUser As UserInfoy
Dim cUser As UserInfoy
Dim Tempint As Integer
Dim MaxUser As Integer
Dim Infor As BaseData
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
MaxUser = Domain.gUserNo
cUser = GetUser(Infor, Index, ealive)
Argument = Trim(Argument)

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for USR module...", Index
    SendReply vbCrLf & "LIST               -Displays user list.", Index
    SendReply vbCrLf & "MYPW [New PW]      -Changes current user password.", Index
    SendReply vbCrLf & "MYINFO             -Displays current user information.", Index
    SendReply vbCrLf & "NEW [New Name]     -Creates a new user.", Index
    SendReply vbCrLf & "OPEN [Index]       -Prepares the specified account for modification.", Index
    SendReply vbCrLf & "VIEW               -Returns opened account's information.", Index
    SendReply vbCrLf & "USER [New Name]    -Changes opened account's username.", Index
    SendReply vbCrLf & "REAL [New Name]    -Changes opened account's realname.", Index
    SendReply vbCrLf & "PASS [New PW]      -Changes opened account's password.", Index
    SendReply vbCrLf & "STAT [1][2][3]     -Changes opened account's status.", Index
    SendReply vbCrLf & "                    1=Enabled", Index
    SendReply vbCrLf & "                    2=Disabled", Index
    SendReply vbCrLf & "                    3=Deleted", Index
    SendReply vbCrLf & "LEVL [1][2][3][4]  -Changes opened account's authority level.", Index
    SendReply vbCrLf & "                    1=Guest", Index
    SendReply vbCrLf & "                    2=Standard", Index
    SendReply vbCrLf & "                    3=Service Admin", Index
    SendReply vbCrLf & "                    4=System Admin", Index
    Exit Function
    End If

    
        If UCase(CMD) = "OPEN" Then
            If cUser.Rights < aServiceAdmin Then
            SendReply vbCrLf & "Access Denied.", Index
            Exit Function
            End If
            If Number(Argument) = True Then
                If (CInt(Argument) < MaxUser) And (CInt(Argument) >= 0) Then
                OpenUser(Index) = GetUser(Infor, CInt(Argument), eArchive)
                OpenUserInt(Index) = CInt(Argument)
                    If (cUser.Rights < aSystemAdmin) And (OpenUser(Index).Rights = aSystemAdmin) Then
                    SendReply vbCrLf & "Access Denied.", Index
                    ClearUser Index
                    Exit Function
                    End If
                SendReply vbCrLf & "Account opened.", Index
                Else
                SendReply vbCrLf & "Index non-existent.", Index
                End If
            Else
            SendReply vbCrLf & "Index required.", Index
            End If
        Exit Function
        End If

        If UCase(CMD) = "NEW" Then
            If cUser.Rights < aServiceAdmin Then
            SendReply vbCrLf & "Access Denied.", Index
            Exit Function
            End If
            If Len(Argument) = 0 Then
            SendReply vbCrLf & "Must specify a new username.", Index
            Exit Function
            End If
            If Domain.UserExists(Replace(Argument, " ", "_")) = True Then
            SendReply vbCrLf & "Username already exists.", Index
            Exit Function
            End If
        Dim Pathy As String
        Dim NewUse As Integer
        NewUse = Domain.gUserNo
        Pathy = Domain.Home & slashval(Domain.Home) & "Data\Users.ini"
        fWriteValue Pathy, "USER" & NewUse, "USER", "S", Replace(Argument, " ", "_")
        fWriteValue Pathy, "USER" & NewUse, "NAME", "S", "New User"
        fWriteValue Pathy, "USER" & NewUse, "PASS", "S", "password"
        fWriteValue Pathy, "USER" & NewUse, "STATUS", "S", 200
        fWriteValue Pathy, "USER" & NewUse, "RIGHTS", "S", Righty.aStandard
        fWriteValue Pathy, "SETUP", "USERSNO", "S", (NewUse + 1)
        DoEvents
        Domain.IniUser
        DoEvents
        SendReply vbCrLf & "User created.", Index
        Exit Function
        End If

        If UCase(CMD) = "MYPW" Then
        fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & Domain.gAccountLink(Index), "PASS", "S", Argument
        DoEvents
        Domain.IniUser
        DoEvents
        SendReply vbCrLf & "Password changed.", Index
        Exit Function
        End If

        If UCase(CMD) = "STAT" Then
            If Len(OpenUser(Index).Username) > 0 Then
                If Number(Argument) = True Then
                Tempint = Round(Argument, 0)
                    If (Tempint >= 1) And (Tempint <= 3) Then
                    fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & OpenUserInt(Index), "STATUS", "S", (Tempint * 100)
                    Domain.IniUser
                    DoEvents
                        If Argument <> 3 Then
                        OpenUser(Index) = GetUser(Infor, OpenUserInt(Index), eArchive)
                        Else
                        ClearUser Index
                        End If
                    SendReply vbCrLf & "Status changed.", Index
                    Else
                    SendReply vbCrLf & "Between 1-3.", Index
                    End If
                Else
                SendReply vbCrLf & "Number required.", Index
                End If
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        If UCase(CMD) = "LEVL" Then
            If Len(OpenUser(Index).Username) > 0 Then
                If Number(Argument) = True Then
                Tempint = Round(Argument, 0)
                    If (Tempint >= 1) And (Tempint <= 4) Then
                        If (Tempint * 100) <= cUser.Rights Then
                        fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & OpenUserInt(Index), "RIGHTS", "S", (Tempint * 100)
                        Domain.IniUser
                        DoEvents
                        OpenUser(Index) = GetUser(Infor, OpenUserInt(Index), eArchive)
                        SendReply vbCrLf & "Level changed.", Index
                        Else
                        SendReply vbCrLf & "Access Denied.", Index
                        End If
                    Else
                    SendReply vbCrLf & "Between 1-4.", Index
                    End If
                Else
                SendReply vbCrLf & "Number required.", Index
                End If
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        If UCase(CMD) = "USER" Then
            If Len(OpenUser(Index).Username) > 0 Then
                If Domain.UserExists(Replace(Argument, " ", "_")) = True Then
                SendReply vbCrLf & "Username exists.", Index
                Exit Function
                End If
            fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & OpenUserInt(Index), "USER", "S", Replace(Argument, " ", "_")
            Domain.IniUser
            DoEvents
            OpenUser(Index) = GetUser(Infor, OpenUserInt(Index), eArchive)
            SendReply vbCrLf & "Username changed.", Index
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        If UCase(CMD) = "REAL" Then
            If Len(OpenUser(Index).Username) > 0 Then
            fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & OpenUserInt(Index), "NAME", "S", Argument
            Domain.IniUser
            DoEvents
            OpenUser(Index) = GetUser(Infor, OpenUserInt(Index), eArchive)
            SendReply vbCrLf & "Realname changed.", Index
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        If UCase(CMD) = "PASS" Then
            If Len(OpenUser(Index).Username) > 0 Then
            fWriteValue Domain.Home & slashval(Domain.Home) & "Data\Users.ini", "USER" & OpenUserInt(Index), "PASS", "S", Argument
            Domain.IniUser
            DoEvents
            OpenUser(Index) = GetUser(Infor, OpenUserInt(Index), eArchive)
            SendReply vbCrLf & "Password changed.", Index
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        
        If UCase(CMD) = "VIEW" Then
            If Len(OpenUser(Index).Username) > 0 Then
            SendReply vbCrLf & vbCrLf & "Opened account information...", Index
            SendReply vbCrLf & "Archive Index: " & OpenUserInt(Index), Index
            SendReply vbCrLf & "Username: " & OpenUser(Index).Username, Index
            SendReply vbCrLf & "Realname: " & OpenUser(Index).Realname, Index
            SendReply vbCrLf & "Password: " & OpenUser(Index).Password, Index
            If OpenUser(Index).Rights = aGuest Then SendReply vbCrLf & "Level: Guest", Index
            If OpenUser(Index).Rights = aStandard Then SendReply vbCrLf & "Level: Standard", Index
            If OpenUser(Index).Rights = aServiceAdmin Then SendReply vbCrLf & "Level: Service Admin", Index
            If OpenUser(Index).Rights = aSystemAdmin Then SendReply vbCrLf & "Level: System Admin", Index
            If OpenUser(Index).Status = aAlive Then SendReply vbCrLf & "Status: Enabled", Index
            If OpenUser(Index).Status = aDisabled Then SendReply vbCrLf & "Status: Disabled", Index
            If OpenUser(Index).Status = aDeleted Then SendReply vbCrLf & "Status: Deleted", Index
            Else
            SendReply vbCrLf & "No account opened.", Index
            End If
        Exit Function
        End If
        
        If UCase(CMD) = "MYINFO" Then
        Tempint = Domain.gAccountLink(Index)
        TmpUser = GetUser(Infor, Tempint, eArchive)
        SendReply vbCrLf & vbCrLf & "Current user information...", Index
        SendReply vbCrLf & "Archive Index: " & Tempint, Index
        SendReply vbCrLf & "Username: " & TmpUser.Username, Index
        SendReply vbCrLf & "Realname: " & TmpUser.Realname, Index
        If TmpUser.Rights = aGuest Then SendReply vbCrLf & "Level: Guest", Index
        If TmpUser.Rights = aStandard Then SendReply vbCrLf & "Level: Standard", Index
        If TmpUser.Rights = aServiceAdmin Then SendReply vbCrLf & "Level: Service Admin", Index
        If TmpUser.Rights = aSystemAdmin Then SendReply vbCrLf & "Level: System Admin", Index
        If TmpUser.Status = aAlive Then SendReply vbCrLf & "Status: Enabled", Index
        If TmpUser.Status = aDisabled Then SendReply vbCrLf & "Status: Disabled", Index
        If TmpUser.Status = aDeleted Then SendReply vbCrLf & "Status: Deleted", Index
        Exit Function
        End If
        
        If UCase(CMD) = "LIST" Then
            If cUser.Rights < aStandard Then
            SendReply vbCrLf & "Access Denied.", Index
            Exit Function
            End If
        Tempint = 0
        SendReply vbCrLf & vbCrLf & "User account list...", Index
        SendReply vbCrLf & "INDEX      : USERNAME           : REALNAME             : LEVEL", Index
            Do While Tempint < MaxUser
            TmpUser = GetUser(Infor, Tempint, eArchive)
            SendReply vbCrLf & Tempint & Space(10 - Len(CStr(Tempint))) & " : " & TmpUser.Username & Space(18 - Len(TmpUser.Username)) & " : " & TmpUser.Realname & Space(20 - Len(TmpUser.Realname)) & " : ", Index
            If TmpUser.Rights = Righty.aGuest Then SendReply "Guest", Index
            If TmpUser.Rights = Righty.aStandard Then SendReply "Standard", Index
            If TmpUser.Rights = Righty.aServiceAdmin Then SendReply "Service Admin", Index
            If TmpUser.Rights = Righty.aSystemAdmin Then SendReply "System Admin", Index
            Tempint = Tempint + 1
            DoEvents
            Loop
        Exit Function
        End If
    
SendReply vbCrLf & "Unknown command.", Index
End Function

Private Sub ClearUser(Index As Integer)
OpenUserInt(Index) = 0
OpenUser(Index).Password = ""
OpenUser(Index).Realname = ""
OpenUser(Index).Rights = 100
OpenUser(Index).Status = 200
OpenUser(Index).Username = ""
End Sub
