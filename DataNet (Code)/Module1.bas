Attribute VB_Name = "Module1"
Public WANport As Integer
Public APPdir As String
Public NoDomains As Integer
Public LogBox(0 To 1000) As frmTelnet
Public Alert(0 To 1000) As LozMes

    Public Enum Stage
    Login = 100
    Password = 200
    Clear = 300
    End Enum

    Public Enum Accy
    aSystemAdmin = 400
    aServiceAdmin = 300
    aStandard = 200
    aGuest = 100
    End Enum

    Public Type ConInfo
    PrevData As String
    Data As String
    Stage As Stage
    IP As String
    SignOn As String
    End Type
    
    Public Type UserInfo
    UserName As String
    Realname As String
    Password As String
    Rights As Accy
    Stage As Stage
    Status As eStatus
    End Type
Public Versy As String
Public CurUp As Long
Public CurDown As Long
Public TotUp As Double
Public TotDown As Double
Public TotBreached As Boolean
Public TotConnections As Double
Public Greet As String
Public AccessCode As String
Public Servy As String
Public fColour As Long
Public dColour As Long
Public fColourFor As Long
Public dColourFor As Long
Public BandMoniter As Boolean

Public Enum Shade
    Light = 100
    Dark = 200
End Enum

Public Const tGA = ""
Public Const tEX = "ÿ"
Public Const tECHO = ""
Public Const tWILL = "û"
Public Const tWONT = "ü"
Public Const tDO = "ý"
Public Const tDONT = "þ"
Public Declare Sub Sleep Lib "kernel32" (Optional ByVal dwMilliseconds As Long = 1)


Public Function Exist(strPath As String) As Boolean
Dim TempInt As Long
    TempInt = (Dir(strPath) = "")
    
If TempInt = 0 Then
Exist = True
Else
Exist = False
End If
End Function
Public Function slashval(path As String) As String
    If Right(path, 1) = "\" Then
    slashval = ""
    Else
    slashval = "\"
    End If
End Function
Public Function IfExist(path As String) As Boolean
Temp = CurDir         '
On Error GoTo endi    '
                      '
ChDir path            '
IfExist = True        '
ChDir Temp            '  Checks to see if the specified path exists,
Exit Function         '  and returns it as a boolean (true or false)
                      '
endi:                 '
IfExist = False       '
ChDir Temp            '
End Function

Public Sub Say(Noticey As String)
Master.NoticeTim.Enabled = False
Master.NOTICEexpirey = 0
Master.SayTxt.FontBold = True
Master.SayTxt.Text = Noticey
Master.NoticeTim.Enabled = True
End Sub

Public Function pFormatSize(ByVal dSize As Double) As String

' 1024  b = 1 kb: 1024 kb = 1 mb

    If dSize < 1024 Then
        pFormatSize = dSize & " bytes"
    Else
        dSize = dSize / 1024
        If dSize < 1000 Then
            pFormatSize = Format$(dSize, "#,##0.0") & " kb"
        Else
            pFormatSize = Format$(dSize / 1024, "#,##0.0") & " mb"
        End If
    End If
    
End Function

Public Function FolderFromPath(FilePath As String) As String
Dim Tempyo As String
Dim Lengy As Long
Lengy = 0
    Do While (Left(Tempyo, 1) = "\") = False
    Lengy = Lengy + 1
    Tempyo = Right(FilePath, Lengy)
        If Lengy > 100 Then
        FolderFromPath = "/NO FOLDER>"
        Exit Function
        End If
    DoEvents
    Loop

FolderFromPath = Left(FilePath, Len(FilePath) - Len(Tempyo))
End Function
Public Function FileFromPath(FilePath As String) As String
Dim Tempyo As String
Dim Lengy As Long
Lengy = 0
    Do While (Left(Tempyo, 1) = "\") = False
    If Lengy > 100 Then GoTo endio
    Lengy = Lengy + 1
    Tempyo = Right(FilePath, Lengy)
    DoEvents
    Loop
FileFromPath = Right(Tempyo, Len(Tempyo) - 1)

Exit Function
endio:
FileFromPath = ""
End Function


Public Function CreatePath(path As String)
On Error GoTo endi
Dim Section() As String
Dim Dump As String
Dim TempInt As Integer
Section = Split(path, "\")

    Do While Len(Section(TempInt)) > 0
    Dump = Dump & Section(TempInt) & slashval(Section(TempInt))
        If IfExist(Dump) = False Then
        MkDir Dump
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
endi:
End Function

Public Function Number(Value As Variant) As Boolean
On Error GoTo endi
Dim Tempy As Integer
Tempy = Value
Number = True
Exit Function
endi:
Number = False
End Function

Public Function HandleCommands(cmd As String) As String
    If cmd = tDO & tECHO Then
'    MsgBox "Server AGREED echo"
    HandleCommands = tEX & tDO & tECHO
    DoEvents
    End If
    If cmd = tDO & tGA Then
'    MsgBox "Server AGREED suppress go aheads"
    HandleCommands = tEX & tDO & tGA
    DoEvents
    End If
    
    If cmd = tWILL & tECHO Then
'    MsgBox "Server WILL echo"
    DoEvents
    End If
    If cmd = tWILL & tGA Then
'    MsgBox "Server WILL suppress go aheads"
    DoEvents
    End If
'MsgBox Asc(Left(cmd, 1)) & " " & Asc(Right(cmd, 1))
End Function

Public Sub SendComs(Index As Integer)
On Error GoTo endi
    If Master.WAN(Index).State = sckConnected Then
    Master.WAN(Index).SendData tEX & tWILL & tECHO
    DoEvents
    Master.WAN(Index).SendData tEX & tWILL & tGA
    DoEvents
    End If
endi:
End Sub
Public Function isLoaded(frm As Form) As Boolean
    Dim i As Integer
    isLoaded = False


    For i = 0 To Forms.Count - 1


        If Forms(i) Is frm Then
            isLoaded = True
            Exit Function
        End If
    Next
End Function
Public Function KeyExists(Tree As ListView, Key As String) As Boolean
On Error GoTo endi
If Len(Tree.ListItems.Item(Key)) > 0 Then KeyExists = True
Exit Function
endi:
KeyExists = False
End Function
Public Function ItemSelected(Tree As ListView) As Boolean
On Error GoTo endi
If Len(Tree.SelectedItem.Key) > 0 Then ItemSelected = True
Exit Function
endi:
ItemSelected = False
End Function

Public Sub ShadeIt(Obj As Object, Shadey As Shade)
    If Shadey = Light Then
    Obj.BackColor = fColour
    Obj.ForeColor = fColourFor
    Else
    Obj.BackColor = dColour
    Obj.ForeColor = dColourFor
    End If
End Sub


Public Function ReferExists(Reference As String, Optional Index As Integer) As Boolean
Dim TempInt As Integer
ReferExists = False
TempInt = 0
    Do While TempInt < UserNo
        If UCase(Trim(Reference)) = UCase(Trim(Master.Domain(TempInt).Refer)) Then
        Index = TempInt
        ReferExists = True
        Exit Function
        End If
    TempInt = TempInt + 1
    DoEvents
    Loop
End Function
