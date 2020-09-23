VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object

Private OverW(0 To 1000) As Boolean
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
Title = "SYS"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the system (sys)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to carry out system processes files." & vbCrLf
Description = Description & "Command list: sys.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
OverW(Index) = False
End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim cUser As UserInfoy
Dim TempInt As Integer
Dim FSO As New FileSystemObject
Dim FeedBack As String
Dim Infor As BaseData
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, eAlive)
Argument = Trim(Argument)

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for SYS module...", Index
    SendReply vbCrLf & "INIT               -Initializes module, must be executed to use module.", Index
    SendReply vbCrLf & "TIME               -Displays the time on the server.", Index
    SendReply vbCrLf & "DATE               -Displays the date on the server.", Index
    SendReply vbCrLf & "DIR                -Displays current directory contents.", Index
    SendReply vbCrLf & "[Drive]:           -Switches to a different drive (system admin only).", Index
    SendReply vbCrLf & "CD [Dir]           -Changes directory (by default uses the ':' wildcard).", Index
    SendReply vbCrLf & "MD [DirPath]       -Creates a new directory.", Index
    SendReply vbCrLf & "RD [DirPath]       -Deletes specified directory.", Index
    SendReply vbCrLf & "COPY [FilePath]    -Copies specifed file(s) to current directory.", Index
    SendReply vbCrLf & "DEL [FilePath]     -Deletes specifed file(s).", Index
    SendReply vbCrLf & "ATR [Path]         -Changes the attributes to the specifed item(s).", Index
    SendReply vbCrLf & "         :[R] :[H]  Replace the R and H with either 1 or 0 to set the attribs", Index
    SendReply vbCrLf & "                    to the specified item (these come after the path).", Index
    SendReply vbCrLf & "ATR Text.txt :0 :1  This example would set the txt file to hidden.", Index
    SendReply vbCrLf & "MOVE [FilePath]    -Copies specifed file(s) to current directory.", Index
    SendReply vbCrLf & "FORCE [Option]     -Enables\disables forceful mode.", Index
    SendReply vbCrLf & "                    0=Disabled [Default]", Index
    SendReply vbCrLf & "                    1=Enabled", Index
    SendReply vbCrLf, Index
    SendReply vbCrLf & "EXAMPLES:           WILD CARDS:", Index
    SendReply vbCrLf & "COPY ??llo.txt      ?=Variable character", Index
    SendReply vbCrLf & "COPY *.txt          *=Variable name", Index
    SendReply vbCrLf & "CD :Dir1            :=Abbreviation for the current dir.", Index
    SendReply vbCrLf & "CD \                \=Abbreviation for the root dir.", Index
    SendReply vbCrLf & "CD ..              ..=Abbreviation for the previus dir.", Index
    Exit Function
    End If

    If UCase(CMD) = "INIT" Then
    Domain.cd Index, Domain.Home & slashval(Domain.Home) & "Users\" & cUser.Username
    CreatePath Domain.cd(Index)
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    SendReply vbCrLf & "System module initialized.", Index
    Exit Function
    End If
    
    If (UCase(Right(CMD, 1)) = ":") And (Len(CMD) = 2) Then
        If cUser.Rights = aSystemAdmin Then
            If IfExist(CMD) = True Then
            Domain.cd Index, CMD
            Else
            SendReply vbCrLf & "Drive does not exist.", Index
            End If
        Else
        SendReply vbCrLf & "Access Denied.", Index
        End If
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If
    
    If UCase(CMD) = "CD" Then
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "Dir: " & FormatPath(Infor, Domain.cd(Index)), Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
        If HasAccess(Infor, RealPath(Infor, Argument), aRead) = True Then
            If IfExist(RealPath(Infor, Argument)) = False Then
            SendReply vbCrLf & "Dir does not exist.", Index
            Exit Function
            End If
        Else
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Domain.cd Index, RealPath(Infor, Argument)
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If
        
    If UCase(CMD) = "TIME" Then
    SendReply vbCrLf & "Time is " & Time & ".", Index
    Exit Function
    End If
    If UCase(CMD) = "DATE" Then
    SendReply vbCrLf & "Date is " & Date & ".", Index
    Exit Function
    End If
        
    If UCase(CMD) = "FORCE" Then
        If Number(Argument) = False Then
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
        If (Round(CInt(Argument)) = 0) Or (Round(CInt(Argument)) = 1) Then
        OverW(Index) = CBool(Round(Argument))
        If (Round(CInt(Argument)) = 0) Then SendReply vbCrLf & "Force disabled.", Index
        If (Round(CInt(Argument)) = 1) Then SendReply vbCrLf & "Force enabled.", Index
        Else
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
    Exit Function
    End If

    If UCase(CMD) = "RD" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
    On Error GoTo reply4
        If HasAccess(Infor, RealPath(Infor, Argument), aWrite) = True Then
            If IfExist(RealPath(Infor, Argument)) = False Then
            SendReply vbCrLf & "Dir does not exist.", Index
            Exit Function
            End If
        Else
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    FSO.DeleteFolder RealPath(Infor, Argument), OverW(Index)
reply4:
        If Len(Err.Description) > 0 Then
        SendReply vbCrLf & Err.Description & ".", Index
        Else
        SendReply vbCrLf & "Directory deleted.", Index
        End If
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If

    If UCase(CMD) = "MD" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "Must specify a path.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
        If HasAccess(Infor, RealPath(Infor, Argument), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    On Error GoTo end8
    MkDir (RealPath(Infor, Argument))
    SendReply vbCrLf & "Directory created.", Index
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
end8:
    SendReply vbCrLf & Err.Description & ".", Index
    Exit Function
    End If

    If UCase(CMD) = "ATR" Then
    Dim System As Boolean
    Dim ReadOnly As Boolean
    Dim Hidden As Boolean
    Dim Diry As Boolean
    Dim Splitty() As String
    Dim Pathyo As String
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
        If cUser.Rights < aServiceAdmin Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Splitty() = Split(Argument, ":")
        If (UBound(Splitty) + 1) >= 1 Then
            If FileAttrib(Trim(RealPath(Infor, Splitty(0))), aRead, ReadOnly, Hidden, System, Diry) = False Then
            SendReply vbCrLf & "Specified file/folder does not exist.", Index
            Exit Function
            End If
            If (UBound(Splitty) + 1) >= 3 Then
            Splitty(1) = Trim(Splitty(1))
            Splitty(2) = Trim(Splitty(2))
                If Number(Splitty(1)) And Number(Splitty(2)) Then
                FileAttrib RealPath(Infor, Splitty(0)), aWrite, CBool(Splitty(1)), CBool(Splitty(2)), System, Diry
                Else
                SendReply vbCrLf & "New attributes must come in the form of 1 or 0." & ReadOnly, Index
                End If
            End If
        DoEvents
        FileAttrib Trim(RealPath(Infor, Splitty(0))), aRead, ReadOnly, Hidden, System, Diry
        SendReply vbCrLf & "Item attributes... ", Index
        SendReply vbCrLf & "READ-ONLY: " & ReadOnly, Index
        SendReply vbCrLf & "HIDDEN:    " & Hidden, Index
        Else
        SendReply vbCrLf & "Must specify a path, refer to sys.help.", Index
        End If
    Exit Function
    End If

    If UCase(CMD) = "DEL" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
    On Error GoTo reply3
        If HasAccess(Infor, RealPath(Infor, Argument), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    FSO.DeleteFile RealPath(Infor, Argument), OverW(Index)
reply3:
        If Len(Err.Description) > 0 Then
        SendReply vbCrLf & Err.Description & ".", Index
        Else
        SendReply vbCrLf & "File(s) deleted.", Index
        End If
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If


    If UCase(CMD) = "MOVE" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
    On Error GoTo reply1
        If HasAccess(Infor, RealPath(Infor, Argument), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
        'Detects * and ? Wildcards
        If (Left(FileFromPath(Argument), 1) = "*") Or (Right(FileFromPath(Argument), 1) = "*") Or (Len(Replace(Argument, "?", "  ")) > Len(Argument)) Then
        FSO.MoveFile RealPath(Infor, Argument), Domain.cd(Index)
        Else
        FSO.MoveFile RealPath(Infor, Argument), Domain.cd(Index) & slashval(Domain.cd(Index)) & FileFromPath(Argument)
        End If
reply1:
        If Len(Err.Description) > 0 Then
        SendReply vbCrLf & Err.Description & ".", Index
        Else
        SendReply vbCrLf & "File(s) moved.", Index
        End If
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If

    If UCase(CMD) = "COPY" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
    On Error GoTo reply2
        If HasAccess(Infor, RealPath(Infor, Argument), aRead) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
        'Detects * and ? Wildcards
        If (Left(FileFromPath(Argument), 1) = "*") Or (Right(FileFromPath(Argument), 1) = "*") Or (Len(Replace(Argument, "?", "  ")) > Len(Argument)) Then
        FSO.CopyFile RealPath(Infor, Argument), Domain.cd(Index), OverW(Index)
        Else
        FSO.CopyFile RealPath(Infor, Argument), Domain.cd(Index) & slashval(Domain.cd(Index)) & FileFromPath(Argument), OverW(Index)
        End If
reply2:
        If Len(Err.Description) > 0 Then
        SendReply vbCrLf & Err.Description & ".", Index
        Else
        SendReply vbCrLf & "File(s) copied.", Index
        End If
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If

    If UCase(CMD) = "DIR" Then
        If Len(Domain.cd(Index)) = 0 Then
        SendReply vbCrLf & "Must initialise module first.", Index
        Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
        Exit Function
        End If
    Dim fol As Folder
    Dim sub1 As Folder
    Dim F As String
    Set fol = FSO.GetFolder(Domain.cd(Index))
    F = Dir(fol & "\*.*", vbReadOnly + vbHidden + vbNormal + vbSystem)
    SendReply vbCrLf & vbCrLf & " Directory of " & FormatPath(Infor, CStr(fol)) & vbCrLf, Index
    On Error Resume Next
        For Each sub1 In fol.SubFolders
            If HasAccess(Infor, sub1.Path, aRead) = True Then
            SendReply vbCrLf & FileFromPath(DELslash(sub1.Path)) & Space(30 - Len(FileFromPath(DELslash(sub1.Path)))) & "[DIR]", Index
            End If
        DoEvents
        Next
    TempInt = 0
        Do While Len(F) > 0
            If HasAccess(Infor, Domain.cd(Index) & slashval(Domain.cd(Index)) & F, aRead) = True Then
            SendReply vbCrLf & F, Index
            End If
        F = Dir
        TempInt = TempInt + 1
        DoEvents
        Loop
    SendReply vbCrLf & Space(29 - Len(CStr(TempInt))) & TempInt & " File(s)", Index
    SendReply vbCrLf & Space(29 - Len(CStr(fol.SubFolders.Count))) & fol.SubFolders.Count & " Dir(s)", Index
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
reply5:
SendReply vbCrLf & Err.Description, Index
    Domain.ChLine Index, FormatPath(Infor, Domain.cd(Index))
    Exit Function
    End If

SendReply vbCrLf & "Unknown command.", Index
End Function
