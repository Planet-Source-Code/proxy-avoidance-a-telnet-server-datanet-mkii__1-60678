VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Domain As Object
Private Page(0 To 1000) As Long
Private FiPath(0 To 1000) As String
Private PNT(0 To 1000) As Boolean
Private WRAP(0 To 1000) As Boolean
Private PNTer(0 To 1000) As Double
Private Armed(0 To 1000) As Boolean

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
Title = "TXT"
End Function
Public Function Description() As String
Description = "Member of the Standard control library; the text (txt)" & vbCrLf
Description = Description & "module holds a range of commands that enables the user" & vbCrLf
Description = Description & "to write to files, and read them." & vbCrLf
Description = Description & "Command list: txt.help" & vbCrLf
End Function
Public Sub Clear(Index As Integer)
FiPath(Index) = ""
PNT(Index) = False
WRAP(Index) = False
PNTer(Index) = 0
Armed(Index) = False
Page(Index) = 1
End Sub

Public Function RunCMD(CMD As String, Argument As String, Index As Integer) As String
Dim cUser As UserInfoy
Dim Tempint As Integer
Dim Infor As BaseData
Dim Fringo As Integer
Dim Data As String
Dim TmpArm As Boolean
Infor.Argument = Argument
Infor.CMD = CMD
Set Infor.Domain = Domain
Infor.Index = Index
cUser = GetUser(Infor, Index, ealive)

    If Armed(Index) = True Then
    TmpArm = True
    Armed(Index) = False
    End If

    If UCase(CMD) = "HELP" Then
    SendReply vbCrLf & vbCrLf & "Help page for SYS module...", Index
    SendReply vbCrLf & "OPEN [FilePath]    -Opens the specified file for read/write.", Index
    SendReply vbCrLf & "READ               -Prints the contents of the opened file to screen.", Index
    SendReply vbCrLf & "PAGE [PageNo]      -Sets the page number (each page is 1900 characters).", Index
    SendReply vbCrLf & "SIZE               -Returns the file size, and its length in pages.", Index
    SendReply vbCrLf & "PNT [Option]       -Displays pointers when a READ command is processed.", Index
    SendReply vbCrLf & "                    0=Disabled [Default]", Index
    SendReply vbCrLf & "                    1=Enabled", Index
    SendReply vbCrLf & "WRAP [Option]      -Automaticly fits text into console window.", Index
    SendReply vbCrLf & "                    0=Disabled [Default]", Index
    SendReply vbCrLf & "                    1=Enabled", Index
    SendReply vbCrLf & "= [Text]           -Writes text into the opened file.", Index
    SendReply vbCrLf & "/ [Text]           -Writes text into the opened file on a new line.", Index
    SendReply vbCrLf & "SET [Pnt]          -Sets the starting point for the next write (/).", Index
    SendReply vbCrLf & "                    (With no argument the pointer is set to end of the file)", Index
    SendReply vbCrLf & "DEL [Pnt1] [Pnt2]  -Deletes anything between the two specified points.", Index
    SendReply vbCrLf & "                    (No argument- the whole file's contents is deleted)", Index
    Exit Function
    End If
    
    If UCase(CMD) = "SET" Then
    On Error GoTo cont4
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
    Fringo = FreeFile
    Argument = Trim(Argument)
    Open FiPath(Index) For Binary As Fringo
        If Len(Argument) = 0 Then
        PNTer(Index) = LOF(Fringo)
        SendReply vbCrLf & "Pointer set to " & PNTer(Index) & ".", Index
        Close Fringo
        Exit Function
        End If
        If Number(Argument, True) = False Then
        SendReply vbCrLf & "Must specify a valid point (0-" & (LOF(Fringo)) & ").", Index
        Close Fringo
        Exit Function
        End If
        If (Round(CLng(Argument)) >= 0) And (Round(CLng(Argument)) <= LOF(Fringo)) Then
        PNTer(Index) = Round(CLng(Argument))
        SendReply vbCrLf & "Pointer set to " & PNTer(Index) & ".", Index
        Close Fringo
        Exit Function
        Else
        SendReply vbCrLf & "Must specify a valid point (0-" & (LOF(Fringo)) & ").", Index
        Close Fringo
        Exit Function
        End If
    Exit Function
cont4:
    SendReply vbCrLf & "Could not set pointer.", Index
    Exit Function
    End If
    
    If UCase(CMD) = "SIZE" Then
    Argument = Trim(Argument)
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If HasAccess(Infor, FiPath(Index), aRead) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    SendReply vbCrLf & vbCrLf & "File Size...", Index
    SendReply vbCrLf & "Size:  " & pFormatSize(FileLen(FiPath(Index))), Index
    SendReply vbCrLf & "Pages: " & RoundNum(FileLen(FiPath(Index)) / 1900, True), Index
    Exit Function
    End If
    
    If UCase(CMD) = "PAGE" Then
    Argument = Trim(Argument)
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If Len(Argument) = 0 Then
        GoTo end9
        End If
        If Number(Argument, True) = False Then
        SendReply vbCrLf & "Page number required.", Index
        Exit Function
        End If
        If Round(CLng(Argument), 0) < 1 Then
        Page(Index) = 1
        GoTo end9
        End If
        If Round(CLng(Argument), 0) > Round(FileLen(FiPath(Index)) / 1900, 0) Then
        Page(Index) = RoundNum(FileLen(FiPath(Index)) / 1900, True)
        GoTo end9
        End If
    Page(Index) = Round(CLng(Argument), 0)
end9:
    SendReply vbCrLf & "Now on page " & Page(Index) & ".", Index
    Exit Function
    End If
    
    If UCase(CMD) = "PNT" Then
    Argument = Trim(Argument)
        If Number(Argument) = False Then
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
        If (Round(CInt(Argument)) = 0) Or (Round(CInt(Argument)) = 1) Then
        PNT(Index) = CBool(Round(Argument))
        If (Round(CInt(Argument)) = 0) Then SendReply vbCrLf & "Pointers disabled.", Index
        If (Round(CInt(Argument)) = 1) Then SendReply vbCrLf & "Pointers enabled.", Index
        Else
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
    Exit Function
    End If

    If UCase(CMD) = "DEL" Then
    Argument = Trim(Argument)
    Dim SplitDel() As String
    Dim Lefty As Integer
    Dim Righty As Integer
    Dim DelIt As Boolean
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If HasAccess(Infor, FiPath(Index), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    SplitDel = Split(Argument, " ")
        If Len(Argument) = 0 Then
            If TmpArm = True Then
            Lefty = 0
            Righty = FileLen(FiPath(Index))
            DelIt = True
            Else
            SendReply vbCrLf & "Type DEL again to confirm file deletion.", Index
            Armed(Index) = True
            End If
        End If
        If UBound(SplitDel) + 1 = 2 Then
            If (Number(SplitDel(0)) = True) And (Number(SplitDel(1)) = True) Then
                If (CInt(SplitDel(0)) >= 0) And (CInt(SplitDel(1)) <= FileLen(FiPath(Index))) Then
                    If CInt(SplitDel(0)) < CInt(SplitDel(1)) Then
                    Lefty = CInt(SplitDel(0))
                    Righty = CInt(SplitDel(1))
                    DelIt = True
                    Else
                    SendReply vbCrLf & "Final point must be greater than starting point.", Index
                    End If
                Else
                SendReply vbCrLf & "Must specify valid points (0-" & (FileLen(FiPath(Index))) & ").", Index
                End If
            Else
            SendReply vbCrLf & "Must specify valid points (0-" & (FileLen(FiPath(Index))) & ").", Index
            End If
        End If
        If DelIt = True Then
        Fringo = FreeFile
        Open FiPath(Index) For Binary As Fringo
        Data = String$(LOF(Fringo), " ")
        Get Fringo, , Data
        Close Fringo
        Open FiPath(Index) For Output As Fringo
        Close Fringo
        Open FiPath(Index) For Binary As Fringo
        Data = Left(Data, Lefty) & Right(Data, Len(Data) - Righty)
        Put Fringo, 1, Data
        Close Fringo
        SendReply vbCrLf & "Text deleted.", Index
        PNTer(Index) = Lefty
        End If
    Exit Function
    End If


    If UCase(CMD) = "WRAP" Then
    Argument = Trim(Argument)
        If Number(Argument) = False Then
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
        If (Round(CInt(Argument)) = 0) Or (Round(CInt(Argument)) = 1) Then
        WRAP(Index) = CBool(Round(Argument))
        If (Round(CInt(Argument)) = 0) Then SendReply vbCrLf & "Text wrapping disabled.", Index
        If (Round(CInt(Argument)) = 1) Then SendReply vbCrLf & "Text wrapping enabled.", Index
        Else
        SendReply vbCrLf & "Requires either 0 or 1.", Index
        Exit Function
        End If
    Exit Function
    End If
    
    If UCase(CMD) = "OPEN" Then
    Argument = Trim(Argument)
        If HasAccess(Infor, RealPath(Infor, Argument), aRead) = True Then
        FiPath(Index) = RealPath(Infor, Argument)
        Fringo = FreeFile
        On Error GoTo cont5
        Open FiPath(Index) For Binary As Fringo
        PNTer(Index) = LOF(Fringo)
        Close Fringo
        Page(Index) = 1
        SendReply vbCrLf & "File opened.", Index
        Exit Function
cont5:
        SendReply vbCrLf & "Could not open file.", Index
        Else
        SendReply vbCrLf & "Access Denied.", Index
        End If
    Exit Function
    End If

    If UCase(CMD) = "/" Then
    On Error GoTo cont2
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If HasAccess(Infor, FiPath(Index), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Fringo = FreeFile
    Open FiPath(Index) For Binary As Fringo
    Data = String$(LOF(Fringo), " ")
    Get Fringo, , Data
    If Len(Data) < PNTer(Index) Then PNTer(Index) = Len(Data)
    Close Fringo
    Open FiPath(Index) For Output As Fringo
    Close Fringo
    Open FiPath(Index) For Binary As Fringo
    Data = Left(Data, PNTer(Index)) & vbCrLf & Argument & Right(Data, Len(Data) - PNTer(Index))
    Put Fringo, 1, Data
    Close Fringo
    PNTer(Index) = PNTer(Index) + Len(vbCrLf & Argument)
    SendReply vbCrLf & "Line added.", Index
    Exit Function
cont2:
    SendReply vbCrLf & "Could not write to file.", Index
    Exit Function
    End If

    If UCase(CMD) = "=" Then
    'On Error GoTo cont6
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If HasAccess(Infor, FiPath(Index), aWrite) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Fringo = FreeFile
    Open FiPath(Index) For Binary As Fringo
    Data = String$(LOF(Fringo), " ")
    Get Fringo, , Data
    If Len(Data) < PNTer(Index) Then PNTer(Index) = Len(Data)
    Close Fringo
    Open FiPath(Index) For Output As Fringo
    Close Fringo
    Open FiPath(Index) For Binary As Fringo
    Data = Left(Data, PNTer(Index)) & Argument & Right(Data, Len(Data) - PNTer(Index))
    Put Fringo, 1, Data
    Close Fringo
    PNTer(Index) = PNTer(Index) + Len(Argument)
    SendReply vbCrLf & "Line added.", Index
    Exit Function
cont6:
    SendReply vbCrLf & "Could not write to file.", Index
    Exit Function
    End If


    If UCase(CMD) = "READ" Then
    On Error GoTo cont3
        If Exist(FiPath(Index)) = False Then
        SendReply vbCrLf & "No file opened, or file does not exist.", Index
        Exit Function
        End If
        If HasAccess(Infor, FiPath(Index), aRead) = False Then
        SendReply vbCrLf & "Access Denied.", Index
        Exit Function
        End If
    Dim Tempint2 As Integer
    Dim Lengo As Long
    Dim WrapLeng As Integer
    Dim ColumSize As Integer
    Dim Liney As String
    Fringo = FreeFile
    Open FiPath(Index) For Binary As Fringo
    ColumSize = Len(CStr(LOF(Fringo))) + 1
        If (((Page(Index) - 1) * 1900) + 1900) <= LOF(Fringo) Then
        Data = String$(1900, " ")
        Else
        Data = String$(LOF(Fringo) - ((Page(Index) - 1) * 1900), " ")
        End If
    Get Fringo, ((Page(Index) - 1) * 1900) + 1, Data
    Close Fringo
    Data = RTrim(Data)
    SendReply vbCrLf & "File contents (page " & Page(Index) & " of " & RoundNum(FileLen(FiPath(Index)) / 1900, True) & ")..." & vbCrLf, Index
    Data = Data & vbCrLf
    DoEvents
    Tempint = 0
    WrapLeng = 75 - ColumSize
    Lengo = ((Page(Index) - 1) * 1900)
        Do While Tempint < Len(Data)
        Liney = ""
        Tempint = Tempint + 1
            'Finds the next new line
            Do While ((Right(Left(Data, Tempint), 1) = vbCr) = False) And ((Right(Left(Data, Tempint), 1) = vbLf) = False)
            Liney = Liney & Right(Left(Data, Tempint), 1)
            Tempint = Tempint + 1
            DoEvents
            Loop
        
            If Right(Left(Data, (Tempint + 1)), 1) = vbCr Then
            Tempint = Tempint + 1
            End If
            If Right(Left(Data, (Tempint + 1)), 1) = vbLf Then
            Tempint = Tempint + 1
            End If
        
            If WRAP(Index) = True Then
            Tempint2 = 1
                If Len(Liney) > 0 Then
                    Do While Tempint2 <= Len(Liney)
                        If PNT(Index) = True Then
                       SendReply vbCrLf & Lengo & Space(ColumSize - Len(CStr(Lengo))) & ": " & Mid(Liney, Tempint2, WrapLeng), Index
                        Else
                        SendReply vbCrLf & Mid(Liney, Tempint2, WrapLeng), Index
                        End If
                    Lengo = Lengo + Len(Mid(Liney, Tempint2, WrapLeng))
                    Tempint2 = Tempint2 + WrapLeng
                    DoEvents
                    Loop
                Else
                    If PNT(Index) = True Then
                    SendReply vbCrLf & Lengo & Space(ColumSize - Len(CStr(Lengo))) & ": ", Index
                    Else
                    SendReply vbCrLf, Index
                    End If
                End If
            Else
                If PNT(Index) = True Then
                SendReply vbCrLf & Lengo & Space(ColumSize - Len(CStr(Lengo))) & ": " & Liney, Index
                Else
                SendReply vbCrLf & Liney, Index
                End If
            End If
        Lengo = ((Page(Index) - 1) * 1900) + Tempint
        DoEvents
        Loop
    Exit Function
cont3:
    SendReply vbCrLf & "Could not read file.", Index
    Exit Function
    End If
    
SendReply vbCrLf & "Unknown command.", Index
End Function




