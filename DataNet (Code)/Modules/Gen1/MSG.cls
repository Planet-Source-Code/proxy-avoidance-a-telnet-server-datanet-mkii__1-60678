VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object
Private Main As Object

Private Type Message
    Text As String
    Users(0 To 100) As String
    FilePath(0 To 100) As String
    InboxList(0 To 100) As String
    InboxCount As Integer
    UserCount As Integer
    FileCount As Integer
    Subject As String
End Type


Dim Msg(0 To 1000) As Message

Public Sub LocateMain(Mainy As Object)
Set Main = Mainy
End Sub
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
Title = "MSG"
Randomize
End Function
Public Function Description() As String
'Enter a description of the module here:
Description = "Member of the Generation 1 control library; the message (msg)" & vbCrLf
Description = Description & "module enables users to send mail messages to one another." & vbCrLf
Description = Description & "Command list: msg.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
'Enter the list of variables that should be reset when a new client connects:
Dim Blanko As Message
Msg(Index) = Blanko
End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim Fringo As Integer
Dim Infor As BaseData     '  Creates a BaseData variable, and fills
Infor.Argument = Argument '  it with the appropraite data, BaseData is
Infor.CMD = CMD           '  required when you use some commands from the
Set Infor.Domain = Domain '  BASE module
Infor.Index = Index       '
Dim TempInt As Integer
Dim cUser As UserInfoy                '  Creates a variable called cUser that
cUser = GetUser(Infor, Index, eAlive) '  holds all of the data on your current user
Dim fol As Folder
Dim sub1 As Folder
Dim f As String
Dim Patho As String
Dim Contents As String
Dim FiParts() As String
Dim Indexy As Integer
Dim Splits() As String
Dim FSO As New FileSystemObject
Patho = Domain.Home & slashval(Domain.Home) & "Data\Mail\" & cUser.Username

'These are all of the commands that your module contains,
'this function works by supplying you with the user's command (CMD)
'and the command argument (Argument). So from there it is easy to
'construct your own commands:

    'Help command
    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for MSG module...", Index
    SendReply vbCrLf & "NEW [Subject]      -Creates a new email.", Index
    SendReply vbCrLf & "LIST               -Lists all emails in your inbox.", Index
    SendReply vbCrLf & "READ [Index]       -Reads specified email.", Index
    SendReply vbCrLf & "GET [Index]        -Retreives attachments from specified email.", Index
    SendReply vbCrLf & "DEL [Index]        -Deletes specified email.", Index
    SendReply vbCrLf & "MYINFO             -Displays your email address.", Index
    Exit Function
    End If

    If UCase(CMD) = "LIST" Then
    On Error GoTo endo15
    CreatePath Patho
    Set fol = FSO.GetFolder(Patho)
    f = Dir(fol & "\*.*", vbReadOnly + vbHidden + vbNormal + vbSystem)
    SendReply vbCrLf & vbCrLf & "Retrieving emails in inbox..." & vbCrLf, Index
    TempInt = 0
        Do While (Len(f) > 0) And (TempInt < 100)
        Fringo = FreeFile
        Open Patho & slashval(Patho) & f For Binary As Fringo
        Contents = String$(LOF(Fringo), " ")
        Get Fringo, , Contents
        Close Fringo
        FiParts = Split(Contents, Chr(11) & Chr(22) & Chr(1) & Chr(2))
            If (UBound(FiParts) + 1) <> 5 Then
            GoTo endo1
            End If
            If UCase(Right(f, 4)) = ".NEW" Then
            SendReply vbCrLf & TempInt & "] " & FiParts(0) & Space(25 - (Len(CStr(FiParts(0))) + Len(CStr(TempInt)) + 1)) & "From: " & FiParts(1) & Space(55 - Len(TempInt & "] " & FiParts(0) & Space(25 - (Len(CStr(FiParts(0))) + Len(CStr(TempInt)) + 1)) & "From: " & FiParts(1))) & "(NEW)", Index
            Msg(Index).InboxCount = TempInt
            Msg(Index).InboxList(TempInt) = Patho & slashval(Patho) & f
            TempInt = TempInt + 1
            End If
            If UCase(Right(f, 4)) = ".OLD" Then
            SendReply vbCrLf & TempInt & "] " & FiParts(0) & Space(25 - (Len(CStr(FiParts(0))) + Len(CStr(TempInt)) + 1)) & "From: " & FiParts(1), Index
            Msg(Index).InboxList(TempInt) = Patho & slashval(Patho) & f
            Msg(Index).InboxCount = TempInt
            TempInt = TempInt + 1
            End If
endo1:
        f = Dir
        Close Fringo
        DoEvents
        Loop
endo15:
    Exit Function
    End If

    If UCase(CMD) = "DEL" Then
    On Error GoTo endo23
        If Number(Argument) = False Then
        SendReply vbCrLf & "Command requires an email index.", Index
        Exit Function
        End If
        If (CInt(Argument) < 0) Or (CInt(Argument) > Msg(Index).InboxCount) Then
        SendReply vbCrLf & "Invalid index, type MSG.LIST.", Index
        Exit Function
        End If
    FSO.DeleteFile Msg(Index).InboxList(CInt(Argument))
        If IfExist(Left(Msg(Index).InboxList(CInt(Argument)), Len(Msg(Index).InboxList(CInt(Argument))) - 4)) = True Then
        FSO.DeleteFolder Left(Msg(Index).InboxList(CInt(Argument)), Len(Msg(Index).InboxList(CInt(Argument))) - 4), True
        End If
    Set fol = FSO.GetFolder(Patho)
    f = Dir(fol & "\*.*", vbReadOnly + vbHidden + vbNormal + vbSystem)
    TempInt = 0
        Do While (Len(f) > 0) And (TempInt < 100)
        Fringo = FreeFile
        Open Patho & slashval(Patho) & f For Binary As Fringo
        Contents = String$(LOF(Fringo), " ")
        Get Fringo, , Contents
        Close Fringo
        FiParts = Split(Contents, Chr(11) & Chr(22) & Chr(1) & Chr(2))
            If (UBound(FiParts) + 1) <> 5 Then
            GoTo endo123
            End If
            If UCase(Right(f, 4)) = ".NEW" Then
            Msg(Index).InboxCount = TempInt
            Msg(Index).InboxList(TempInt) = Patho & slashval(Patho) & f
            TempInt = TempInt + 1
            End If
            If UCase(Right(f, 4)) = ".OLD" Then
            Msg(Index).InboxList(TempInt) = Patho & slashval(Patho) & f
            Msg(Index).InboxCount = TempInt
            TempInt = TempInt + 1
            End If
endo123:
        f = Dir
        Close Fringo
        DoEvents
        Loop
    SendReply vbCrLf & "Email deleted.", Index
    Exit Function
endo23:
    SendReply vbCrLf & "Could not delete email.", Index
    Exit Function
    End If


    If UCase(CMD) = "GET" Then
    On Error GoTo endo11
    Dim FolderSrc As String
    Dim FolderDes As String
        If Number(Argument) = False Then
        SendReply vbCrLf & "Command requires an email index.", Index
        Exit Function
        End If
        If (CInt(Argument) < 0) Or (CInt(Argument) > Msg(Index).InboxCount) Then
        SendReply vbCrLf & "Invalid index, type MSG.LIST.", Index
        Exit Function
        End If
    FolderSrc = Left(Msg(Index).InboxList(CInt(Argument)), Len(Msg(Index).InboxList(CInt(Argument))) - 4)
        If IfExist(FolderSrc) = False Then
        SendReply vbCrLf & "Email does not contain any attachments.", Index
        Exit Function
        End If
    Fringo = FreeFile
    Open Msg(Index).InboxList(CInt(Argument)) For Binary As Fringo
    Contents = String$(LOF(Fringo), " ")
    Get Fringo, , Contents
    Close Fringo
    FiParts = Split(Contents, Chr(11) & Chr(22) & Chr(1) & Chr(2))
        If (UBound(FiParts) + 1) <> 5 Then
        GoTo endo11
        End If
Retryit:
    FolderDes = Domain.Home & slashval(Domain.Home) & "Users\" & cUser.Username & slashval(cUser.Username) & "MSG ATTACH\" & FiParts(0) & "_" & Round(Rnd * 1000, 0)
        If IfExist(FolderDes) = True Then
        GoTo Retryit
        End If
    CreatePath FolderDes
    Set fol = FSO.GetFolder(FolderSrc)
    f = Dir(fol & "\*.*", vbReadOnly + vbHidden + vbNormal + vbSystem)
    SendReply vbCrLf & vbCrLf & "Retrieving attachments...", Index
    TempInt = 0
        Do While Len(f) > 0
        FileCopy FolderSrc & slashval(FolderSrc) & f, FolderDes & slashval(FolderDes) & f
        TempInt = TempInt + 1
        f = Dir
        DoEvents
        Loop
    SendReply vbCrLf & TempInt & " attachment(s) retrieved.", Index
    Exit Function
endo11:
    SendReply vbCrLf & "Error retrieving attachments.", Index
    Exit Function
    End If

    If UCase(CMD) = "READ" Then
        If Number(Argument) = False Then
        SendReply vbCrLf & "Command requires an email index.", Index
        Exit Function
        End If
        If (CInt(Argument) < 0) Or (CInt(Argument) > Msg(Index).InboxCount) Then
        SendReply vbCrLf & "Invalid index, type MSG.LIST.", Index
        Exit Function
        End If
    Fringo = FreeFile
    On Error GoTo endo2
    Open Msg(Index).InboxList(CInt(Argument)) For Binary As Fringo
    Contents = String$(LOF(Fringo), " ")
    Get Fringo, , Contents
    Close Fringo
        If (Right(Msg(Index).InboxList(CInt(Argument)), 4) = ".NEW") Then
        DoEvents
        FSO.MoveFile Msg(Index).InboxList(CInt(Argument)), Left(Msg(Index).InboxList(CInt(Argument)), Len(Msg(Index).InboxList(CInt(Argument))) - 4) & ".OLD"
        End If
    FiParts = Split(Contents, Chr(11) & Chr(22) & Chr(1) & Chr(2))
    SendReply vbCrLf, Index
    SendReply vbCrLf & "-MESSAGE (CONFIDENTIAL)----------------------------------------------------", Index
    SendReply vbCrLf & "NO. OF ATTACHMENTS: " & FiParts(4), Index
    SendReply vbCrLf & "         DATE SENT: " & FiParts(2), Index
    SendReply vbCrLf & "              FROM: " & FiParts(1), Index
    SendReply vbCrLf, Index
    SendReply vbCrLf & "           SUBJECT: " & FiParts(0), Index
    SendReply vbCrLf & "===========================================================================", Index
    SendReply vbCrLf & FiParts(3), Index
    Exit Function
endo2:
    SendReply vbCrLf & "Could not read email.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "MYINFO" Then
    SendReply vbCrLf & "This is your email address: " & UCase(cUser.Username) & "@" & UCase(Domain.refer), Index
    Exit Function
    End If

    If UCase(CMD) = "NEW" Then
        If Len(Trim(Argument)) = 0 Then
        SendReply vbCrLf & "Your message must have a subject.", Index
        Exit Function
        End If
        If UCase(Argument) = "EXIT" Then
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        SendReply vbCrLf & "Message abandoned.", Index
        Exit Function
        End If
    Clear Index
    Msg(Index).Subject = Trim(Argument)
    SendHeader Infor
    SendReply vbCrLf & "Please enter the username of the recipitent. When you are finished", Index
    SendReply vbCrLf & "adding users, type #", Index
    Domain.CMD "Com:MSG.USER", Index
    Domain.ChLine Index, "SEND TO"
    Exit Function
    End If
    
    If UCase(CMD) = "USER" Then
        If Len(Trim(Argument)) = 0 Then
        GoTo listMofo
        Exit Function
        End If
        If UCase(Argument) = "EXIT" Then
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        SendReply vbCrLf & "Message abandoned.", Index
        Exit Function
        End If
        If UCase(Argument) = "#" Then
        SendHeader Infor
        SendReply vbCrLf & "Please enter the path of any files you wish to send. When you are finished", Index
        SendReply vbCrLf & "adding files, type #", Index
        Domain.CMD "Com:MSG.FILE", Index
        Domain.ChLine Index, "ATTACHMENT PATH"
        Exit Function
        End If
        If Msg(Index).UserCount < 100 Then
        Msg(Index).Users(Msg(Index).UserCount) = UCase(Trim(Argument))
        Msg(Index).UserCount = Msg(Index).UserCount + 1
listMofo:
        SendHeader Infor
        SendReply vbCrLf & "Please enter the username of the recipitent. When you are finished", Index
        SendReply vbCrLf & "adding users, type #", Index
        TempInt = 0
            Do While TempInt < Msg(Index).UserCount
            Splits = Split(Msg(Index).Users(TempInt), "@")
                If (UBound(Splits) + 1) <> 2 Then
                SendReply vbCrLf & Msg(Index).Users(TempInt) & " - Invalid address.", Index
                GoTo SkipIty
                End If
                
                If Main.RefExists(Splits(1), Indexy) = True Then
                    If Main.Domain(Indexy).UserExists(Splits(0)) = True Then
                    SendReply vbCrLf & Msg(Index).Users(TempInt) & " - Valid address.", Index
                    Else
                    SendReply vbCrLf & Msg(Index).Users(TempInt) & " - Does not exist.", Index
                    End If
                Else
                SendReply vbCrLf & Msg(Index).Users(TempInt) & " - Does not exist.", Index
                End If
SkipIty:
            TempInt = TempInt + 1
            DoEvents
            Loop
        Domain.CMD "Com:MSG.USER", Index
        Domain.ChLine Index, "SEND TO"
        End If
    Exit Function
    End If

    If UCase(CMD) = "FILE" Then
        If Len(Trim(Argument)) = 0 Then
        GoTo ListFiles
        Exit Function
        End If
        If UCase(Argument) = "EXIT" Then
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        SendReply vbCrLf & "Message abandoned.", Index
        Exit Function
        End If
        If UCase(Argument) = "#" Then
        SendHeader Infor
        SendReply vbCrLf & "Please enter your message. When you are finished, type #", Index
        Domain.CMD "Com:MSG.TEXT", Index
        Domain.ChLine Index, "MESSAGE"
        Exit Function
        End If
        If Msg(Index).FileCount < 100 Then
        Msg(Index).FilePath(Msg(Index).FileCount) = UCase(RealPath(Infor, Trim(Argument)))
        Msg(Index).FileCount = Msg(Index).FileCount + 1
ListFiles:
        SendHeader Infor
        SendReply vbCrLf & "Please enter the path of any files you wish to send. When you are finished", Index
        SendReply vbCrLf & "adding files, type #", Index
        TempInt = 0
            Do While TempInt < Msg(Index).FileCount
                If HasAccess(Infor, Msg(Index).FilePath(TempInt), aRead) = False Then
                SendReply vbCrLf & FormatPath(Infor, Msg(Index).FilePath(TempInt)) & " - Access Denied.", Index
                GoTo ItsOk
                End If
                If Exist(Msg(Index).FilePath(TempInt)) = False Then
                SendReply vbCrLf & FormatPath(Infor, Msg(Index).FilePath(TempInt)) & " - Does not exist.", Index
                GoTo ItsOk
                End If
            SendReply vbCrLf & FormatPath(Infor, Msg(Index).FilePath(TempInt)) & " - File ready.", Index
ItsOk:
            TempInt = TempInt + 1
            DoEvents
            Loop
        Domain.CMD "Com:MSG.FILE", Index
        Domain.ChLine Index, "ATTACHMENT PATH"
        End If
    Exit Function
    End If

    If UCase(CMD) = "TEXT" Then
    
        If UCase(Argument) = "EXIT" Then
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        SendReply vbCrLf & "Message abandoned.", Index
        Exit Function
        End If
        
        If UCase(Argument) = "#" Then
        Dim Tempint2 As Integer
        Dim TimeDec As String
        Dim Pathoy As String
        Dim RealAtNo As Integer
        SendHeader Infor
        SendReply vbCrLf & vbCrLf & "->Sending message...", Index
        TempInt = 0
            Do While TempInt < Msg(Index).UserCount
            Splits = Split(Msg(Index).Users(TempInt), "@")
                If (UBound(Splits) + 1) = 2 Then
                    If Main.RefExists(Splits(1), Indexy) = True Then
                        If Main.Domain(Indexy).UserExists(Splits(0)) = True Then
                        TimeDec = Day(Date) & "-" & Month(Date) & "-" & Right(Year(Date), 2)
                        TimeDec = TimeDec & " [" & Hour(Time) & Minute(Time) & "]"
                        Fringo = FreeFile
                        Pathoy = Main.Domain(Indexy).Home & slashval(Main.Domain(Indexy).Home) & "Data\Mail\" & Splits(0) & slashval(Splits(0)) & TimeDec & "_" & Round(Rnd * 1000, 0)
                        CreatePath Main.Domain(Indexy).Home & slashval(Main.Domain(Indexy).Home) & "Data\Mail\" & Splits(0)
                        SendReply vbCrLf & Msg(Index).Users(TempInt) & " received message,", Index
                            If Msg(Index).FileCount > 0 Then
                            CreatePath Pathoy
                            Tempint2 = 0
                            RealAtNo = 0
                                Do While Tempint2 < Msg(Index).FileCount
                                    If (Exist(Msg(Index).FilePath(Tempint2)) = True) And (HasAccess(Infor, Msg(Index).FilePath(Tempint2), aRead) = True) Then
                                    FileCopy Msg(Index).FilePath(Tempint2), Pathoy & slashval(Pathoy) & FileFromPath(Msg(Index).FilePath(Tempint2))
                                    RealAtNo = RealAtNo + 1
                                    End If
                                Tempint2 = Tempint2 + 1
                                DoEvents
                                Loop
                            End If
                        SendReply " " & RealAtNo & " file attachment(s) sent.", Index
                        Contents = Msg(Index).Subject & Chr(11) & Chr(22) & Chr(1) & Chr(2) & UCase(cUser.Username) & "@" & UCase(Domain.refer) & Chr(11) & Chr(22) & Chr(1) & Chr(2) & Date & " [" & Time & "]" & Chr(11) & Chr(22) & Chr(1) & Chr(2) & Msg(Index).Text & Chr(11) & Chr(22) & Chr(1) & Chr(2) & RealAtNo
                        Open Pathoy & ".NEW" For Binary As Fringo
                        Put Fringo, , Contents
                        Close Fringo

                        Else
                        SendReply vbCrLf & Msg(Index).Users(TempInt) & " does not exist.", Index
                        End If
                    Else
                    SendReply vbCrLf & Msg(Index).Users(TempInt) & " does not exist.", Index
                    End If
                Else
                SendReply vbCrLf & Msg(Index).Users(TempInt) & " is an invalid address.", Index
                End If
            TempInt = TempInt + 1
            DoEvents
            Loop
        Domain.CMD "Com:", Index
        Domain.ChLine Index, cUser.Username
        Exit Function
        End If
    SendHeader Infor
    SendReply vbCrLf & "Please enter your message. When you are finished type #", Index

        If Len(Msg(Index).Text & Argument & vbCrLf) < 1500 Then
        Msg(Index).Text = Msg(Index).Text & vbCrLf & Argument
        Else
        SendReply vbCrLf & "Message body cannot exceed 1500 characters, try sending the main part of your", Index
        SendReply vbCrLf & "message as an attachment.", Index
        End If
    SendReply vbCrLf & Right(Msg(Index).Text, 300), Index
    Domain.CMD "Com:MSG.TEXT", Index
    Domain.ChLine Index, "MESSAGE"
    Exit Function
    End If

'Returns a message if the command could not be found:
SendReply vbCrLf & "Unknown command.", Index
End Function

Private Function SendHeader(info As BaseData)
Dim cUser As UserInfoy                '  Creates a variable called cUser that
cUser = GetUser(info, info.Index, eAlive) '  holds all of the data on your current user
SendReply Chr(27) & "[2J", info.Index
SendReply Chr(27) & "[H", info.Index
SendReply vbCrLf & "-WRITE NEW MESSAGE---------------------------------------------------------", info.Index
SendReply vbCrLf & "NO. OF RECIPITENTS: " & Msg(info.Index).UserCount, info.Index
SendReply vbCrLf & "NO. OF ATTACHMENTS: " & Msg(info.Index).FileCount, info.Index
SendReply vbCrLf & "              FROM: " & UCase(cUser.Username) & "@" & UCase(info.Domain.refer), info.Index
SendReply vbCrLf, info.Index
SendReply vbCrLf & "           SUBJECT: " & Msg(info.Index).Subject, info.Index
SendReply vbCrLf & "===========================================================================", info.Index
End Function
