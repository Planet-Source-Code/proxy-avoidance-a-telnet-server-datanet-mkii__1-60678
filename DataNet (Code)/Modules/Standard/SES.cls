VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object

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
Title = "SES"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the session (ses)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to manipulate current sessions within the domain." & vbCrLf
Description = Description & "Command list: ses.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)

End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim TmpUser As UserInfoy
Dim cUser As UserInfoy
Dim Tempint As Integer
Dim Infor As BaseData
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, ealive)
Argument = Trim(Argument)

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for SES module...", Index
    SendReply vbCrLf & "MYINFO             -Displays current session data.", Index
    SendReply vbCrLf & "LIST               -Displays list of current sessions.", Index
    SendReply vbCrLf & "TERM [Index]       -Terminates specified session.", Index
    SendReply vbCrLf & "EXIT               -Terminates your session.", Index
    Exit Function
    End If

    If UCase(CMD) = "MYINFO" Then
    SendReply vbCrLf & vbCrLf & "Current session information...", Index
    SendReply vbCrLf & "Alive Index = " & Index, Index
    Exit Function
    End If

    If UCase(CMD) = "EXIT" Then
    SendReply vbCrLf & "Termination initiated.", Index
    Domain.gSocket(Index).SendData vbCrLf & "Session terminated." & vbCrLf
    DoEvents
    Domain.gSocket(Index).Close
    Exit Function
    End If

    If UCase(CMD) = "TERM" Then
        If cUser.Rights < aServiceAdmin Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
        If Number(Argument) = True Then
            If (CInt(Argument) >= 0) And (CInt(Argument) < Domain.gSocketCount) Then
                If (Domain.gSocket(CInt(Argument)).State = 7) Then
                    If (GetUser(Infor, CInt(Argument), ealive).Rights < aSystemAdmin) Then
                    SendReply vbCrLf & "Termination initiated.", Index
                    Domain.gSocket(CInt(Argument)).SendData vbCrLf & "Session terminated." & vbCrLf
                    DoEvents
                    Domain.gSocket(CInt(Argument)).Close
                    Else
                    SendReply vbCrLf & "Access Denied.", Index
                    End If
                Else
                SendReply vbCrLf & "Session dead.", Index
                End If
            Else
            SendReply vbCrLf & "Between 0-" & (Domain.gSocketCount - 1) & ".", Index
            End If
        Else
        SendReply vbCrLf & "Index required.", Index
        End If
    Exit Function
    End If

    If UCase(CMD) = "LIST" Then
        If cUser.Rights < aStandard Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Tempint = 0
    SendReply vbCrLf & vbCrLf & "Current session list...", Index
    SendReply vbCrLf & "INDEX      : USERNAME           : REALNAME             : LEVEL", Index
        Do While Tempint < Domain.gSocketCount
            If Domain.gSocket(Tempint).State = 7 Then
            TmpUser = GetUser(Infor, Tempint, ealive)
            SendReply vbCrLf & Tempint & Space(10 - Len(CStr(Tempint))) & " : " & TmpUser.Username & Space(18 - Len(TmpUser.Username)) & " : " & TmpUser.Realname & Space(20 - Len(TmpUser.Realname)) & " : ", Index
            If TmpUser.Rights = Righty.aGuest Then SendReply "Guest", Index
            If TmpUser.Rights = Righty.aStandard Then SendReply "Standard", Index
            If TmpUser.Rights = Righty.aServiceAdmin Then SendReply "Service Admin", Index
            If TmpUser.Rights = Righty.aSystemAdmin Then SendReply "System Admin", Index
            End If
        Tempint = Tempint + 1
        DoEvents
        Loop
    Exit Function
    End If
    
SendReply vbCrLf & "Unknown command.", Index
End Function

