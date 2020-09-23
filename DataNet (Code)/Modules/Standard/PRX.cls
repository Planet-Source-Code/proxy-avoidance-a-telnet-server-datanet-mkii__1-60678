VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object
Private RoomMess As String
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
Title = "COM"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the communications" & vbCrLf
Description = Description & "(com) module holds a range of commands that enables users" & vbCrLf
Description = Description & "to communicate with one another." & vbCrLf
Description = Description & "Command list: com.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)

End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim cUser As UserInfoy
Dim mUser As UserInfoy
Dim TempInt As Integer
Dim Infor As BaseData
Dim Lines() As String
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, eAlive)
    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for COM module...", Index
    SendReply vbCrLf & "ROOM               -Initialises chatroom mode.", Index
    SendReply vbCrLf & "SEND [User] [Text] -Sends a message to a logged-on user.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "SEND" Then
    Dim Party() As String
    Dim dMessage As String
    Dim dUser As String
    Dim Sentit As Boolean
    Argument = Trim(Argument)
    Party = Split(Argument, " ")
        If (UBound(Party) + 1) < 2 Then
        SendReply vbCrLf & "Message not sent - missing parameters.", Index
        Exit Function
        End If
    dUser = Party(0)
    dMessage = ""
    TempInt = 1
        Do While TempInt <= UBound(Party)
        dMessage = dMessage & Party(TempInt) & " "
        TempInt = TempInt + 1
        DoEvents
        Loop
    dMessage = Trim(dMessage)
    dMessage = "-Message from " & StrConv(cUser.Username, vbProperCase) & ":" & vbCrLf & dMessage
    TempInt = 0
        Do While TempInt < Domain.gSocketCount
        mUser = GetUser(Infor, TempInt, eAlive)
            If Trim(UCase(dUser)) = Trim(UCase(mUser.Username)) Then
            Sentit = True
            Call Domain.reply(TempInt, vbCrLf & vbCrLf & dMessage)
            Call Domain.sendline(TempInt, True)
            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
        If Sentit = True Then
        SendReply vbCrLf & "Message sent.", Index
        Else
        SendReply vbCrLf & "User not logged on.", Index
        End If
    Exit Function
    End If
    
    If UCase(CMD) = "ROOM" Then
    SendReply Chr(27) & "[2J", Index
    SendReply Chr(27) & "[H", Index
    SendReply " <:::::::::::::::::::DOMAIN CHATROOM - TYPE 'EXIT' TO LEAVE:::::::::::::::::::>", Index
    SendReply vbCrLf, Index
    Domain.CMD "Com:COM./", Index
    Domain.ChLine Index, "CHAT"
    Lines = Split(RoomMess, vbCrLf)
    TempInt = UBound(Lines) - 15
    If TempInt < 0 Then TempInt = 0
        Do While TempInt < UBound(Lines)
        SendReply Lines(TempInt) & vbCrLf, Index
        TempInt = TempInt + 1
        DoEvents
        Loop
    SendReply Chr(27) & "[22;1f", Index
    Exit Function
    End If

    If UCase(CMD) = "/" Then
    Dim TooBig As Boolean
        If Len(RoomMess) > 1500 Then
        RoomMess = Right(RoomMess, 750)
        End If
        If Len(Argument) > 0 Then
            If Len(Argument) <= 100 Then
            RoomMess = RoomMess & vbCrLf & cUser.Username & "(" & Index & "):" & vbCrLf & Argument & vbCrLf
            TempInt = 0
                Do While TempInt < Domain.gSocketCount
                    If (UCase(Trim(Domain.gAdopted(TempInt, True))) = "COM./") And (TempInt <> Index) Then
                    SendReply Chr(27) & "[23;1f" & "-New Message", TempInt
                    Domain.sendline TempInt, False
                    End If
                TempInt = TempInt + 1
                DoEvents
                Loop
            Else
            TooBig = True
            End If
        End If
        If UCase(Argument) = "EXIT" Then
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        SendReply vbCrLf & "Chatroom mode disabled.", Index
        Exit Function
        End If
    SendReply Chr(27) & "[2J", Index
    SendReply Chr(27) & "[H", Index
    SendReply " <:::::::::::::::::::DOMAIN CHATROOM - TYPE 'EXIT' TO LEAVE:::::::::::::::::::>", Index
    SendReply vbCrLf, Index
    Domain.CMD "Com:COM./", Index
    Domain.ChLine Index, "CHAT"
    Lines = Split(RoomMess, vbCrLf)
    TempInt = UBound(Lines) - 15
    If TempInt < 0 Then TempInt = 0
        Do While TempInt < UBound(Lines)
        SendReply Lines(TempInt) & vbCrLf, Index
        TempInt = TempInt + 1
        DoEvents
        Loop
        If TooBig = True Then
        SendReply Chr(27) & "[23;1f" & "-Message too big", TempInt
        End If
    SendReply Chr(27) & "[22;1f", Index
    Exit Function
    End If
SendReply vbCrLf & "Unknown command.", Index
End Function

Private Sub ClearHistory()
Dim Lines() As String
Dim TempInt As Integer
    RoomMess = ""
    Lines = Split(RoomMess, vbCrLf)
    TempInt = UBound(Lines) - 15
    If TempInt < 0 Then TempInt = 0
        Do While TempInt < UBound(Lines)
        RoomMess = RoomMess & Lines(TempInt) & vbCrLf
        TempInt = TempInt + 1
        DoEvents
        Loop
End Sub

