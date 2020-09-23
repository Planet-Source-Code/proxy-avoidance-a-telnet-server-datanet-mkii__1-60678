VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EXE"
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
Title = "EXE"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the execute (exe)" & vbCrLf
Description = Description & "module holds a few of commands that enables the user" & vbCrLf
Description = Description & "to execute console apps and scripts." & vbCrLf
Description = Description & "Command list: exe.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)

End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim cUser As UserInfoy
Dim TempInt As Integer
Dim Infor As BaseData
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, eAlive)
Argument = Trim(Argument)

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for EXE module...", Index
    SendReply vbCrLf & "SHELL               -Executes a console application.", Index
    SendReply vbCrLf & "SCRPT [Path] [Arg]  -Executes a DataNet script.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "SHELL" Then
    SendReply vbCrLf & "Command still under construction.", Index
    Exit Function
    End If

    If UCase(CMD) = "SCRPT" Then
    Dim Party() As String
    Dim Coms() As String
    Dim Pathy As String
    Dim Fringo As Integer
    Dim Contents As String
    Dim Commandy As String
    Argument = Trim(Argument)
        If Len(Argument) = 0 Then
        SendReply vbCrLf & "Must specify script's path.", Index
        Exit Function
        End If
    Party = Split(Argument, " ")
    Pathy = Party(0) 'Script Path
    Party = Split(Right(Argument, Len(Argument) - Len(Pathy)), "#") 'Script Arguments (untrimmed)
    Pathy = RealPath(Infor, Pathy) 'Script Path
        If HasAccess(Infor, Pathy, aRead) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Fringo = FreeFile
    Open Pathy For Binary As Fringo
    Contents = String$(LOF(Fringo), " ")
    Get Fringo, , Contents
    Close Fringo
    TempInt = 0
        Do While TempInt <= 9
            If TempInt + 1 <= UBound(Party) Then
            Contents = Replace(Contents, "%" & TempInt & "%", Trim(Party(TempInt + 1)))
            Else
            Contents = Replace(Contents, "%" & TempInt & "%", "")
            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
    Coms() = Split(Contents, vbCrLf)
    TempInt = 0
        Do While TempInt <= UBound(Coms())
            If Len(Coms(TempInt)) > 0 Then
            SendReply Domain.CMD(Coms(TempInt), Index), Index
            Else
                If Len(Domain.gAdopted(Index, True)) > 0 Then
                SendReply Domain.CMD(Coms(TempInt), Index), Index
                End If
            End If
        TempInt = TempInt + 1
        DoEvents
        Loop
    Exit Function
    End If
    
SendReply vbCrLf & "Unknown command.", Index
End Function




