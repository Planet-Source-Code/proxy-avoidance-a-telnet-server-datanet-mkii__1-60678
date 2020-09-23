Attribute VB_Name = "Proc"
Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260
Const sLocation As String = "mdlProcess"

Private Const READ_CONTROL As Long = &H20000

Private Const SYNCHRONIZE As Long = &H100000

Private Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Private Const STANDARD_RIGHTS_READ As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE As Long = READ_CONTROL
Private Const STANDARD_RIGHTS_ALL As Long = &H1F0000
Private Const SPECIFIC_RIGHTS_ALL As Long = &HFFFF
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_EXECUTE As Long = &H20000000
Private Const GENERIC_ALL As Long = &H10000000
Private Const PROCESS_TERMINATE As Long = &H1
Private Const PROCESS_CREATE_THREAD As Long = &H2
Private Const PROCESS_SET_SESSIONID As Long = &H4
Private Const PROCESS_VM_OPERATION As Long = &H8
Private Const PROCESS_VM_READ As Long = &H10
Private Const PROCESS_VM_WRITE As Long = &H20
Private Const PROCESS_DUP_HANDLE As Long = &H40
Private Const PROCESS_CREATE_PROCESS As Long = &H80
Private Const PROCESS_SET_QUOTA As Long = &H100
Private Const PROCESS_SET_INFORMATION As Long = &H200
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Private Const DELETE As Long = &H10000
Private Const WRITE_DAC As Long = &H40000
Private Const WRITE_OWNER As Long = &H80000


Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
    End Type


Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

    Dim ListOfActiveProcess() As PROCESSENTRY32

Public Function szExeFile(ByVal Index As Long) As String
    szExeFile = TrimIt(ListOfActiveProcess(Index).szExeFile)
End Function


Public Function dwFlags(ByVal Index As Long) As Long
    dwFlags = ListOfActiveProcess(Index).dwFlags
End Function


Public Function pcPriClassBase(ByVal Index As Long) As Long
    pcPriClassBase = ListOfActiveProcess(Index).pcPriClassBase
End Function


Public Function th32ParentProcessID(ByVal Index As Long) As Long
    th32ParentProcessID = ListOfActiveProcess(Index).th32ParentProcessID
End Function


Public Function cntThreads(ByVal Index As Long) As Long
    cntThreads = ListOfActiveProcess(Index).cntThreads
End Function


Public Function thModuleID(ByVal Index As Long) As Long
    thModuleID = ListOfActiveProcess(Index).th32ModuleID
End Function


Public Function th32DefaultHeapID(ByVal Index As Long) As Long
    th32DefaultHeapID = ListOfActiveProcess(Index).th32DefaultHeapID
End Function


Public Function th32ProcessID(ByVal Index As Long) As Long
    th32ProcessID = ListOfActiveProcess(Index).th32ProcessID
End Function


Public Function cntUsage(ByVal Index As Long) As Long
    cntUsage = ListOfActiveProcess(Index).cntUsage
End Function


Public Function dwSize(ByVal Index As Long) As Long
    dwSize = ListOfActiveProcess(Index).dwSize
End Function


Public Function GetActiveProcess() As Long
    Dim hToolhelpSnapshot As Long
    Dim tProcess As PROCESSENTRY32
    Dim r As Long, i As Integer
    hToolhelpSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)


    If hToolhelpSnapshot = 0 Then
        GetActiveProcess = 0
        Exit Function
    End If


    With tProcess
        .dwSize = Len(tProcess)
        r = ProcessFirst(hToolhelpSnapshot, tProcess)
        ReDim Preserve ListOfActiveProcess(20)


        Do While r
            i = i + 1
            If i Mod 20 = 0 Then ReDim Preserve ListOfActiveProcess(i + 20)
            ListOfActiveProcess(i) = tProcess
            r = ProcessNext(hToolhelpSnapshot, tProcess)
        Loop
    End With
    GetActiveProcess = i
    Call CloseHandle(hToolhelpSnapshot)
End Function



Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    
    Dim hProcess As Long
    Dim lExitCode As Long
    
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID): If hProcess = 0 Then Call Err_Dll(Err.LastDllError, "OpenProcess failed", sLocation, "Kill_Process")
    
    If GetExitCodeProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "GetExitCodeProcess failed", sLocation, "Kill_Process")
    If TerminateProcess(hProcess, lExitCode) = False Then Call Err_Dll(Err.LastDllError, "TerminateProcess failed", sLocation, "Kill_Process")
    
    If CloseHandle(hProcess) = False Then Call Err_Dll(Err.LastDllError, "CloseHandle failed", sLocation, "Kill_Process")
End Sub

Public Sub Err_Dll(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub Err_Vb(ErrorNum As Long, ErrorDesc As String, Source As String, SubOrFunction As String)
    File_WriteTo "VB ERROR: " & ErrorNum & " at " & Source & "\" & SubOrFunction & " >>> " & ErrorDesc
End Sub
Public Sub File_WriteTo(Text As String)
    '// Allways use this in programs:
    Open App.Path & "\PROGRAM.LOG" For Append As #1
        Print #1, Text
    Close #1
End Sub

Public Function Path_Parse(Path As String) As String
    '// Takes a full file specification and returns the path
    Dim A
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            '// Add the correct path separator for the input
            If Mid$(Path, A, 1) = "\" Then
                Path_Parse = LCase$(Left$(Path, A - 1) & "\")
            Else
                Path_Parse = LCase$(Left$(Path, A - 1) & "/")
            End If
            Exit Function
        End If
    Next A
End Function


Private Function TrimIt(ProcessPath As String) As String
On Error GoTo endi
Dim Spli() As String
Spli = Split(ProcessPath, Chr(0))
TrimIt = Spli(0)
endi:
End Function
