Attribute VB_Name = "Base"
Public Type UserInfoy
    Username As String
    Realname As String
    Password As String
    Rights As Righty
    Stage As Stagey
    Status As eStatus
End Type
Public Enum Stagey
    Login = 100
    Password = 200
    Clear = 300
End Enum
Public Enum Righty
    aSystemAdmin = 400
    aServiceAdmin = 300
    aStandard = 200
    aGuest = 100
End Enum
Public Enum eStatus
    aAlive = 100
    aDisabled = 200
    aDeleted = 300
End Enum
Public Enum eType
    eAlive = 100
    eArchive = 200
End Enum
Public Type BaseData
    Index As Integer
    CMD As String
    Argument As String
    Domain As Object
End Type
Public Type OverW
    Code As String
    Expirey As Integer
    Levely As Righty
    Remain As Integer
    Alive(0 To 1000) As Boolean
    ResetUser(0 To 1000) As Righty
    Owner As String
End Type
Public Enum Attry
    aRead = 100
    aWrite = 200
End Enum
Public Const MAX_PATH = 260


Type FILETIME ' 8 Bytes
    dwLowDateTime As Long
    dwHighDateTime As Long
    End Type


Type WIN32_FIND_DATA ' 318 Bytes
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved_ As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
    End Type


Public Declare Function FindFirstFile& Lib "kernel32" _
    Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData _
    As WIN32_FIND_DATA)


Public Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long



Public Function Number(Value As Variant, Optional isLong As Boolean = False) As Boolean
On Error GoTo endi
    If isLong = True Then
    Dim Tempy1 As Long
    Tempy1 = Value
    Else
    Dim Tempy2 As Integer
    Tempy2 = Value
    End If
Number = True
Exit Function
endi:
Number = False
End Function
Public Function slashval(Path As String) As String
    If Right(Path, 1) = "\" Then
    slashval = ""
    Else
    slashval = "\"
    End If
End Function
Public Function slashvalB(Path As String) As String
    If Left(Path, 1) = "\" Then
    slashvalB = ""
    Else
    slashvalB = "\"
    End If
End Function
Public Function DELslash(Path As String) As String
    If Right(Path, 1) = "\" Then
    DELslash = Left(Path, Len(Path) - 1)
    Else
    DELslash = Path
    End If
End Function
Public Function DELslashB(Path As String) As String
    If Left(Path, 1) = "\" Then
    DELslashB = Right(Path, Len(Path) - 1)
    Else
    DELslashB = Path
    End If
End Function
Public Function CreatePath(Path As String)
On Error GoTo endi
Dim Section() As String
Dim Dump As String
Dim TempInt As Integer
Section = Split(Path, "\")

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
Public Function IfExist(Path As String) As Boolean
Dim FSO As New FileSystemObject
IfExist = FSO.FolderExists(Path)
End Function
Public Function Exist(strPath As String) As Boolean
Dim FSO As New FileSystemObject
Exist = FSO.FileExists(strPath)
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
Public Function RevPath(FilePath As String) As String
Dim Tempyo As String
Dim Lengy As Long
Lengy = 0
    Do While (Right(Tempyo, 1) = "\") = False
    If Lengy > 100 Then GoTo endio
    Lengy = Lengy + 1
    Tempyo = Left(FilePath, Lengy)
    DoEvents
    Loop
RevPath = Left(Tempyo, Len(Tempyo) - 1)

Exit Function
endio:
RevPath = ""
End Function
Public Function GetUser(info As BaseData, Index As Integer, Typey As eType) As UserInfoy
Dim tRights As Integer
Dim tStage As Integer
Dim tStatus As Integer
    If Typey = eAlive Then
    info.Domain.gAliveUser Index, GetUser.Username, GetUser.Realname, GetUser.Password, tRights, tStage, tStatus
    GetUser.Rights = tRights
    GetUser.Stage = tStage
    GetUser.Status = tStatus
    End If
    If Typey = eArchive Then
    info.Domain.gArchiveUser Index, GetUser.Username, GetUser.Realname, GetUser.Password, tRights, tStage, tStatus
    GetUser.Rights = tRights
    GetUser.Stage = tStage
    GetUser.Status = tStatus
    End If
End Function
Public Function SetAliveUser(info As BaseData, Index As Integer, NewInfo As UserInfoy)
Dim tRights As Integer
Dim tStage As Integer
Dim tStatus As Integer
info.Domain.sAliveUser Index, NewInfo.Username, NewInfo.Realname, NewInfo.Password, CInt(NewInfo.Rights), CInt(NewInfo.Stage), CInt(NewInfo.Status)
End Function
Public Function HasAccess(info As BaseData, Path As String, Intentions As Attry) As Boolean
Dim cUser As UserInfoy
Dim aParent As String
Dim FSO As New FileSystemObject
Dim Hidden As Boolean
Dim System As Boolean
Dim Diry As Boolean
Dim ReadOnly As Boolean

cUser = GetUser(info, info.Index, eAlive)
    If cUser.Rights = aSystemAdmin Then
    HasAccess = True
        aParent = DELslashB(FolderFromPath(slashvalB(Path) & DELslash(Path)))
            If FileAttrib(Path, aRead) Then
            FileAttrib Path, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
            If (FileAttrib(aParent, aRead)) Then
            FileAttrib aParent, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
    Exit Function
    End If
    
    If cUser.Rights = aServiceAdmin Then
        If UCase(DELslash(info.Domain.Home)) = UCase(DELslash(Left(Path, Len(DELslash(info.Domain.Home))))) Then
        HasAccess = True
        aParent = DELslashB(FolderFromPath(slashvalB(Path) & DELslash(Path)))
            If FileAttrib(Path, aRead) Then
            FileAttrib Path, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
            If (FileAttrib(aParent, aRead)) Then
            FileAttrib aParent, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
            If UCase(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "data")) = UCase(DELslash(Left(Path, Len(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "data"))))) Then
            HasAccess = False
            End If
        Else
        HasAccess = False
        End If
        
    Exit Function
    End If
    
    If cUser.Rights <= aStandard Then
        If UCase(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "Users\" & cUser.Username)) = UCase(DELslash(Left(Path, Len(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "Users\" & cUser.Username))))) Then
        HasAccess = True
        aParent = DELslashB(FolderFromPath(slashvalB(Path) & DELslash(Path)))
            If FileAttrib(Path, aRead) Then
            FileAttrib Path, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
            If (FileAttrib(aParent, aRead)) Then
            FileAttrib aParent, aRead, ReadOnly, Hidden, System
                If ReadOnly And (Intentions = aWrite) Then
                HasAccess = False
                End If
                If Hidden Then
                HasAccess = False
                End If
                If System Then
                HasAccess = False
                End If
            End If
        Else
        HasAccess = False
        End If
    End If
End Function

Public Function FormatPath(info As BaseData, Path As String) As String
Dim Starter As String
Dim cUser As UserInfoy
cUser = GetUser(info, info.Index, eAlive)

    If cUser.Rights = aSystemAdmin Then
    FormatPath = UCase(DELslash(Path))
    Exit Function
    End If
    
    If cUser.Rights = aServiceAdmin Then
    Starter = FileFromPath(slashvalB(DELslash(info.Domain.Home)) & DELslash(info.Domain.Home))
    FormatPath = UCase(DELslash(Starter & slashval(Starter) & DELslashB(Right(Path, Len(Path) - Len(DELslash(info.Domain.Home))))))
    Exit Function
    End If
    
    If cUser.Rights <= aStandard Then
    Starter = cUser.Username
    FormatPath = UCase(DELslash(Starter & slashval(Starter) & DELslashB(Right(Path, Len(Path) - Len(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "Users\" & cUser.Username))))))
    End If
End Function

Public Function RealPath(info As BaseData, Path As String) As String
Dim cUser As UserInfoy
Dim TmpPath As String
Dim TmpStri As String
Dim Calc As Boolean
Dim NewPathy As String
cUser = GetUser(info, info.Index, eAlive)
TmpPath = Path
Calc = False
    If Left(TmpPath, 1) = ":" Then
    TmpStri = FormatPath(info, info.Domain.cd(info.Index))
    TmpPath = Right(TmpPath, Len(TmpPath) - 1)
    TmpPath = TmpStri & slashval(TmpStri) & TmpPath
    Calc = True
    GoTo Proces
    End If
    If Left(TmpPath, 2) = ".." Then
    TmpStri = FolderFromPath(DELslash(FormatPath(info, info.Domain.cd(info.Index))))
    TmpPath = Right(TmpPath, Len(TmpPath) - 2)
    TmpPath = TmpStri & slashval(TmpStri) & TmpPath
    Calc = True
    GoTo Proces
    End If
    If Left(TmpPath, 1) = "\" Then
    TmpStri = FormatPath(info, info.Domain.cd(info.Index))
    TmpPath = Right(TmpPath, Len(TmpPath) - 1)
    TmpPath = RevPath(TmpStri & slashval(TmpStri)) & slashval(RevPath(TmpStri & slashval(TmpStri))) & TmpPath
    Calc = True
    GoTo Proces
    End If
Proces:
    If Calc = False Then
    TmpPath = FormatPath(info, info.Domain.cd(info.Index)) & slashval(FormatPath(info, info.Domain.cd(info.Index))) & TmpPath
    End If
    
    If cUser.Rights = aSystemAdmin Then
    RealPath = TmpPath
    Exit Function
    End If
    
    If cUser.Rights = aServiceAdmin Then
    NewPathy = DELslashB(FolderFromPath(slashvalB(DELslash(info.Domain.Home)) & DELslash(info.Domain.Home)))
    RealPath = NewPathy & slashval(NewPathy) & TmpPath
    Exit Function
    End If
    
    If cUser.Rights <= aStandard Then
    NewPathy = DELslashB(FolderFromPath(slashvalB(DELslash(info.Domain.Home & slashval(info.Domain.Home) & "Users\" & cUser.Username)) & DELslash(info.Domain.Home & slashval(info.Domain.Home) & "Users\" & cUser.Username)))
    RealPath = NewPathy & slashval(NewPathy) & TmpPath
    End If

End Function

Public Function FileAttrib(Path As String, Action As Attry, Optional ReadOnly As Boolean = False, Optional Hidden As Boolean = False, Optional System As Boolean = False, Optional Diry As Boolean = False) As Boolean
Dim Attriby As Long
Dim FSO As New FileSystemObject
    If (FSO.FileExists(Path)) Or (FSO.FolderExists(Path)) Then
    FileAttrib = True
    Else
    FileAttrib = False
    Exit Function
    End If

    If Action = aRead Then
    Attriby = GetAttr(Path)
        If Attriby >= 64 Then
        Attriby = Attriby - 64
        End If
        If Attriby >= 32 Then
        Attriby = Attriby - 32
        End If
        If Attriby >= 16 Then
        Attriby = Attriby - 16
        Diry = True
        End If
        If Attriby >= 8 Then
        Attriby = Attriby - 8
        End If
        If Attriby >= 4 Then
        Attriby = Attriby - 4
        System = True
        End If
        If Attriby >= 2 Then
        Attriby = Attriby - 2
        Hidden = True
        End If
        If Attriby >= 1 Then
        Attriby = Attriby - 1
        ReadOnly = True
        End If
    Else
    Attriby = 0
        If ReadOnly = True Then
        Attriby = Attriby + 1
        End If
        If Hidden = True Then
        Attriby = Attriby + 2
        End If
        If System = True Then
        Attriby = Attriby + 4
        End If
    SetAttr Path, Attriby
    End If
End Function

Public Function StrCut(String1 As String, NChrs As Integer, StrArray() As String)
    ' make this a public function to be able
    '     to use the returned array (StrArray())
    Dim StrLen As Integer
    Dim CounterA As Integer
    Dim StartCount As Integer
    Dim LenDiff As Integer
    Dim LenTest As Single
    StrLen = Len(String1) 'get the length of our string
    LenTest = StrLen / NChrs ' initialize LenTest
    
    StartCount = 1 'initialize StartCount
    If LenTest > Int(LenTest) Then LenTest = Int(LenTest) + 1 'we only want To work With Integers
    ReDim StrArray(1 To LenTest) 'Size our array
    


    For CounterA = 1 To LenTest
        StrArray(CounterA) = Mid$(String1, StartCount, NChrs) 'fill our array
        StartCount = StartCount + NChrs 'increment our starting point in the Mid$ Function
    Next
    


    If Len(StrArray(LenTest)) < NChrs Then ' see if we need To put spaces on the End of StrArray(LenTest)
        LenDiff = NChrs - Len(StrArray(LenTest)) ' if so how many?

        For CounterA = 1 To LenDiff
            StrArray(LenTest) = StrArray(LenTest) + " " 'add spaces If needed To the String in StrArray(LenTest)
        Next
    End If
    
End Function

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

Public Function RoundNum(Number As Variant, Optional RoundUP As Boolean = True) As Long
    Dim TempNo As Long
    Dim Half As Variant
    Half = 0.5


    If RoundUP = True Then
        TempNo = Number + Half
    Else
        TempNo = Number - Half
    End If
    RoundNum = Round(TempNo, 0)
End Function

