VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SCR"
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
Title = "SCR"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the screen (scr)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to manipulate their display/terminal." & vbCrLf
Description = Description & "Command list: scr.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)

End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim TmpUser As UserInfoy
Dim Tempint As Integer
Dim Infor As BaseData
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
Argument = Trim(Argument)

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for SCR module...", Index
    SendReply vbCrLf & "CLS                -Clears the screen.", Index
    SendReply vbCrLf & "ECHO               -Displays specified text on screen.", Index
    SendReply vbCrLf & "FORC [CODE]        -Sets the foreground colour.", Index
    SendReply vbCrLf & "BAKC [CODE]        -Sets the background colour.", Index
    SendReply vbCrLf & "                   -Colour codes:", Index
    SendReply vbCrLf & "                    0=Black", Index
    SendReply vbCrLf & "                    1=Red", Index
    SendReply vbCrLf & "                    2=Green", Index
    SendReply vbCrLf & "                    3=Yellow", Index
    SendReply vbCrLf & "                    4=Blue", Index
    SendReply vbCrLf & "                    5=Magenta", Index
    SendReply vbCrLf & "                    6=Cyan", Index
    SendReply vbCrLf & "                    7=White", Index
    SendReply vbCrLf, Index
    SendReply vbCrLf & "SCR commands do not return feedback if successfull,", Index
    SendReply vbCrLf & "this is to benefit scripts and mods that use SCR commands.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "CLS" Then
    SendReply Chr(27) & "[2J", Index
    SendReply Chr(27) & "[H", Index
    Exit Function
    End If
    
    If UCase(CMD) = "BEEP" Then
    SendReply Chr(7), Index
    Exit Function
    End If
    
    If UCase(CMD) = "FORC" Then
        If Number(Argument) = True Then
        Tempint = Round(CInt(Argument), 0)
            If (Tempint >= 0) And (Tempint <= 7) Then
            Tempint = Tempint + 30
            SendReply Chr(27) & "[" & Tempint & "m", Index
            Else
            SendReply vbCrLf & "Between 1-7.", Index
            End If
        Else
        SendReply vbCrLf & "Code required.", Index
        End If
    Exit Function
    End If
    
    If UCase(CMD) = "ECHO" Then
    SendReply vbCrLf & Argument, Index
    Exit Function
    End If
    
    If UCase(CMD) = "BAKC" Then
        If Number(Argument) = True Then
        Tempint = Round(CInt(Argument), 0)
            If (Tempint >= 0) And (Tempint <= 7) Then
            Tempint = Tempint + 40
            SendReply Chr(27) & "[" & Tempint & "m", Index
            Else
            SendReply vbCrLf & "Between 1-7.", Index
            End If
        Else
        SendReply vbCrLf & "Code required.", Index
        End If
    Exit Function
    End If
    
    If UCase(CMD) = "RES" Then
    SendReply Chr(27) & "[0m", Index
    Exit Function
    End If
    
SendReply vbCrLf & "Unknown command.", Index
End Function


