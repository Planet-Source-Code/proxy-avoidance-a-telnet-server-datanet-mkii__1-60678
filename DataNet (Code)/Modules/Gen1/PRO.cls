VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRO"
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
Title = "PRO"
End Function
Public Function Description() As String
'Enter a description of the module here:
Description = "Member of the Generation 1 control library; the process (pro)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to carry out commands that can intereact with the system processes." & vbCrLf
Description = Description & "Command list: pro.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
'Enter the list of variables that should be reset when a new client connects:

End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim Infor As BaseData     '  Creates a BaseData variable, and fills
Infor.Argument = Argument '  it with the appropraite data, BaseData is
Infor.CMD = CMD           '  required when you use some commands from the
Set Infor.Domain = Domain '  BASE module
Infor.Index = Index       '
Dim TempInt As Integer
Dim cUser As UserInfoy                '  Creates a variable called cUser that
cUser = GetUser(Infor, Index, eAlive) '  holds all of the data on your current user

Argument = Trim(Argument)  ' Trims the arguemnt, you may want to delete this

'These are all of the commands that your module contains,
'this function works by supplying you with the user's command (CMD)
'and the command argument (Argument). So from there it is easy to
'construct your own commands:

    'Help command
    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for PRO module...", Index
    SendReply vbCrLf & "LIST               -Lists system processes.", Index
    SendReply vbCrLf & "KILL [PID]         -Kills a system process.", Index
    SendReply vbCrLf & "RUN [PATH]         -Runs a program.", Index
    SendReply vbCrLf, Index
    SendReply vbCrLf & "NOTE: System Administrators Only.", Index
    Exit Function
    End If

    If cUser.Rights < aSystemAdmin Then
    SendReply vbCrLf & "Access Denied.", Index
    Exit Function
    End If

    'Time command
    If UCase(CMD) = "LIST" Then
    Dim Porcess() As String
    Dim Nom As Integer
    Nom = Proc.GetActiveProcess
    Proc.GetActiveProcess
    SendReply vbCrLf & vbCrLf & "-System process list...", Index
    SendReply vbCrLf & "PID   : APPLICATION TITLE", Index
    TempInt = 1
        Do While TempInt <= Nom
        SendReply vbCrLf & Proc.th32ProcessID(TempInt) & Space(6 - Len(CStr(Proc.th32ProcessID(TempInt)))) & ": " & Proc.szExeFile(TempInt), Index
        TempInt = TempInt + 1
        DoEvents
        Loop
    Exit Function
    End If
    
    If UCase(CMD) = "KILL" Then
        If Number(Argument) = False Then
        SendReply vbCrLf & "Command requires a process ID (PID).", Index
        Exit Function
        End If
    Proc.Process_Kill CLng(Argument)
    SendReply vbCrLf & "Process killed.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "RUN" Then
    On Error GoTo end1
        If Exist(Argument) = False Then
        SendReply vbCrLf & "File not found.", Index
        Exit Function
        End If
    Shell Argument, vbNormalFocus
    SendReply vbCrLf & "Process started.", Index
    Exit Function
end1:
    SendReply vbCrLf & "Process would not start.", Index
    Exit Function
    End If


'Returns a message if the command could not be found:
SendReply vbCrLf & "Unknown command.", Index
End Function

