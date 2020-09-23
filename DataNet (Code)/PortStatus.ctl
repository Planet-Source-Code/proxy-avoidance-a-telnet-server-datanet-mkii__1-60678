VERSION 5.00
Begin VB.UserControl PortStatus 
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   ScaleHeight     =   3435
   ScaleWidth      =   6585
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PORT STATUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFC0FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "PortStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/////////////////////////////////////////////////////////////////////////
' This code were explicitly developed for PSC(Planet Source Code) Users,
' as Open Source Project. This code are property of their author.
'
' You may use any of this code in you're own application(s).
'
' (c) Luprix  2004
' luprixnet@hotmail.com
'/////////////////////////////////////////////////////////////////////////




'///////////////////////////// Constants and Types ////////////////////////
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Private Const MAX_PATH As Long = 260
Private Const SE_DEBUG_NAME As String = "SeDebugPrivilege"

Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Const PROCESS_VM_READ = &H10
Private Const PROCESS_DUP_HANDLE = &H40
Private Const PROCESS_QUERY_INFORMATION = &H400

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const GENERIC_ALL = &H10000000

Private Const INVALID_HANDLE_VALUE = -1
Private Const SystemHandleInformation = 16&
Private Const ObjectNameInformation = 1&

Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004

Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId  As Integer
    CreatorBackTraceIndex  As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    Object  As Long
    GrantedAccess As Long
End Type

Private Type SYSTEM_HANDLE_INFORMATION
    NumberOfHandles As Long
    Handles() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type

Private Type OBJECT_NAME_PRIVATE
    Length          As Integer
    MaximumLength   As Integer
    Buffer          As Long
    ObjName(23)     As Byte
End Type

Private Type TDI_CONNECTION_INFO
    State               As Long
    Event               As Long
    TransmittedTsdus    As Long
    ReceivedTsdus       As Long
    TransmissionErrors  As Long
    ReceiveErrors       As Long
    Throughput          As LARGE_INTEGER
    Delay               As LARGE_INTEGER
    SendBufferSize      As Long
    ReceiveBufferSize   As Long
    Unreliable          As Boolean
End Type

Private Type TDI_CONNECTION_INFORMATION
    UserDataLength      As Long
    UserData            As Long
    OptionsLength       As Long
    Options             As Long
    RemoteAddressLength As Long
    RemoteAddress       As Long
End Type

Private Type IO_STATUS_BLOCK
    Status As Long
    Information As Long
End Type


Public Enum ProtocallType
    UDP = 100
    TCP = 200
End Enum

'///////////////////////////// Declarations ///////////////////////////////

'Undocumented Native API
Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" ( _
    ByVal dwInfoType As Long, _
    ByVal lpStructure As Long, _
    ByVal dwSize As Long, _
    dwReserved As Long) As Long

Private Declare Function NtQueryObject Lib "ntdll.dll" ( _
    ByVal ObjectHandle As Long, _
    ByVal ObjectInformationClass As Long, _
    ObjectInformation As OBJECT_NAME_PRIVATE, _
    ByVal Length As Long, _
    ResultLength As Long) As Long

Private Declare Function NtDeviceIoControlFile Lib "ntdll.dll" ( _
    ByVal FileHandle As Long, _
    ByVal pEvent As Long, _
    ApcRoutine As Long, _
    ApcContext As Long, _
    IoStatusBlock As IO_STATUS_BLOCK, _
    ByVal IoControlCode As Long, _
    InputBuffer As TDI_CONNECTION_INFORMATION, _
    ByVal InputBufferLength As Long, _
    OutputBuffer As TDI_CONNECTION_INFO, _
    ByVal OutputBufferLength As Long) As Long

'Win32 API
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" _
    Alias "LookupPrivilegeValueA" ( _
    ByVal lpSystemName As String, _
    ByVal lpName As String, _
    lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" ( _
    ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, _
    ByRef NewState As TOKEN_PRIVILEGES, _
    ByVal BufferLength As Long, _
    ByRef PreviousState As TOKEN_PRIVILEGES, _
    ByRef ReturnLength As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
    ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    ByRef TokenHandle As Long) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long) As Long
    
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function DuplicateHandle Lib "kernel32" ( _
    ByVal hSourceProcessHandle As Long, _
    ByVal hSourceHandle As Long, _
    ByVal hTargetProcessHandle As Long, _
    lpTargetHandle As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwOptions As Long) As Long

Private Declare Function CreateEvent Lib "kernel32.dll" _
    Alias "CreateEventW" ( _
    ByRef lpEventAttributes As SECURITY_ATTRIBUTES, _
    ByVal bManualReset As Long, _
    ByVal bInitialState As Long, _
    ByVal lpName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" _
    Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As Long)

Private Declare Function EnumProcessModules Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByRef lphModule As Long, _
    ByVal cb As Long, _
    ByRef cbNeeded As Long) As Long

Private Declare Function ntohs Lib "ws2_32.dll" ( _
     ByVal netshort As Integer) As Integer

Private Declare Function GetModuleFileNameExA Lib "psapi.dll" ( _
    ByVal hProcess As Long, _
    ByVal hModule As Long, _
    ByVal ModuleName As String, _
    ByVal nSize As Long) As Long

'Global Vars
Dim Privilege As Boolean
Dim ResultPorts(1, 65535) As Long

Private Function OpenPort() As Boolean
Dim i As Long, Status As Long
Dim Ret As Long, NumHandles As Long
Dim HandleInfo As SYSTEM_HANDLE_INFORMATION
Dim RequiredLength As Long
Dim Buffer() As Byte

Do
    ReDim Buffer(0 To 19)
    RequiredLength = 20 'len SYSTEM_HANDLE_INFORMATION

    'first, find the RequiredLength for the SYSTEM_HANDLE_INFORMATION array
    Status = NtQuerySystemInformation(SystemHandleInformation, _
          ByVal VarPtr(Buffer(0)), ByVal RequiredLength, 0&)

    If Status = 0 Then
        Exit Do
    End If
    
    'obtain, RequiredLength
    CopyMemory ByVal VarPtr(NumHandles), ByVal VarPtr(Buffer(0)), 4
    RequiredLength = NumHandles * 16 + 4
    ReDim Buffer(0 To RequiredLength)

    'Native API NTDLL. Find system information
    Status = NtQuerySystemInformation(SystemHandleInformation, _
          ByVal VarPtr(Buffer(0)), ByVal RequiredLength, 0&)

    ReDim HandleInfo.Handles(NumHandles)
    CopyMemory ByVal VarPtr(HandleInfo.Handles(0)), _
        ByVal VarPtr(Buffer(4)), RequiredLength - 4

Loop While Status = STATUS_INFO_LENGTH_MISMATCH

For i = 0 To NumHandles - 1
    Call GetPortFromTcpHandle(HandleInfo.Handles(i).UniqueProcessId, _
         HandleInfo.Handles(i).HandleValue)
Next i

OpenPort = True

End Function

Private Function GetPortFromTcpHandle(ProcessId As Integer, hCurrent As Integer) As Boolean
Dim hPort As Long, Port As Long
Dim RequiredLength As Long
Dim ResultLength As Long
Dim Status As Long
Dim hProc As Long
Dim Ret As Long
Dim strFile As String
Dim pObjName As OBJECT_NAME_PRIVATE

If ProcessId = 0 Then
    Exit Function
End If

'Duplicate Handle for the Process
hProc = OpenProcess(PROCESS_DUP_HANDLE, 0&, ProcessId)
If hProc = INVALID_HANDLE_VALUE Then
    Exit Function
End If
Ret = DuplicateHandle(hProc, hCurrent, -1, hPort, _
    STANDARD_RIGHTS_ALL Or GENERIC_ALL, 0&, 0&)

If Ret Then
    RequiredLength = LenB(pObjName)
    
    'Native API. Find handle type "File"
    Status = NtQueryObject(hPort, ObjectNameInformation, _
         pObjName, RequiredLength, ResultLength)
    
    If Status = 0 Then
        'Filter handle names "\device\tcp" and "device\udp"
        If pObjName.Length = 11 * 2 Then   'len ( \device\tcp ) = 11
            Port = 0
            strFile = pObjName.ObjName
            strFile = UCase(Clip(strFile))
            
            Port = QueryDevice(hPort)
            If Port Then
                If InStr(strFile, "TCP") Then
                    ResultPorts(0, Port) = ProcessId
                Else
                    ResultPorts(1, Port) = ProcessId
                End If
            End If
        End If
    End If
End If

'Close all duplicated Handle's !!
Ret = CloseHandle(hPort)
Ret = CloseHandle(hProc)

GetPortFromTcpHandle = True

End Function

Private Function QueryDevice(hPort As Long) As Long
Dim TdiConnInfo As TDI_CONNECTION_INFO
Dim TdiConnInformation As TDI_CONNECTION_INFORMATION
Dim IoStatusBlock As IO_STATUS_BLOCK
Dim TdiIoControl As Long
Dim Status As Long
Dim hEven As Long
Dim secAttrib As SECURITY_ATTRIBUTES
Dim Ret As Long

'    //Tdi layer
' Create new Tdi Event
hEven = CreateEvent(secAttrib, 1, 0, 0)
TdiConnInformation.RemoteAddressLength = 3

TdiIoControl = &H210012 'FILE_DEVICE_TRANSPORT, Reserved Function 1, METHOD_OUT_DIRECT

'Native API. Fill TDI_CONNECTION_INFORMATION
Status = NtDeviceIoControlFile(hPort, hEven, 0&, 0&, IoStatusBlock, TdiIoControl, _
    TdiConnInformation, LenB(TdiConnInformation), TdiConnInfo, LenB(TdiConnInfo))

If hEven Then
    Ret = CloseHandle(hEven)
End If

If Status Then
    Exit Function
End If

'Obtains the Port
QueryDevice = ntohs(UnsignedToInteger(TdiConnInfo.ReceivedTsdus And 65535))

If QueryDevice < 0 Then
    QueryDevice = QueryDevice + 65536
End If

End Function

Private Function UnsignedToInteger(Value As Long) As Integer
'Convert "Unsigned Integer" to "Vb Integer"
    If Value <= MAXINT_2 Then
        UnsignedToInteger = Value
    Else
        UnsignedToInteger = Value - OFFSET_2
    End If
End Function

Private Function Clip(strClip As String) As String
'Discard final null
Dim intNullPos As Integer
   
intNullPos = InStr(strClip, vbNullChar)
If intNullPos > 0 Then
    Clip = Left(strClip, intNullPos - 1)
End If

End Function

Private Function LoadPrivilege(ByVal Privilege As String) As Boolean
'The access
Dim hToken As Long
Dim SEDebugNameValue As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim hProcessHandle As Long
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lbuffer As Long

hProcessHandle = GetCurrentProcess()
If GetLastError <> 0 Then
    MsgBox "GetCurrentProcess, Error: " & GetLastError()
    Exit Function
End If

OpenProcessToken hProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hToken
If GetLastError <> 0 Then
    MsgBox "OpenProcessToken, Error: " & GetLastError()
    Exit Function
End If

LookupPrivilegeValue "", Privilege, SEDebugNameValue
If GetLastError <> 0 Then
    MsgBox "LookupPrivilegeValue, Error: " & GetLastError()
    Exit Function
End If

With tkp
    .PrivilegeCount = 1
    .TheLuid = SEDebugNameValue
    .Attributes = SE_PRIVILEGE_ENABLED
End With

AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkpNewButIgnored, lbuffer
If GetLastError <> 0 Then
    MsgBox "AdjustTokenPrivileges, Error: " & GetLastError()
    Exit Function
End If
    
LoadPrivilege = True

End Function


Private Function ProcessPathByPID(PID As Long) As String
'Return path to the executable from PID
'http://support.microsoft.com/default.aspx?scid=kb;en-us;187913
Dim cbNeeded As Long
Dim Modules(1 To 200) As Long
Dim Ret As Long
Dim ModuleName As String
Dim nSize As Long
Dim hProcess As Long

hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
    Or PROCESS_VM_READ, 0, PID)
            
If hProcess <> 0 Then
                
    Ret = EnumProcessModules(hProcess, Modules(1), _
        200, cbNeeded)
                
    If Ret <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = 500
        Ret = GetModuleFileNameExA(hProcess, _
            Modules(1), ModuleName, nSize)
        ProcessPathByPID = Left(ModuleName, Ret)
    End If
End If
          
Ret = CloseHandle(hProcess)

If ProcessPathByPID = "" Then
    ProcessPathByPID = "SYSTEM"
End If

End Function


Public Function PortInUse(Port As Long, ProtoType As ProtocallType) As Boolean
Dim PathBuf As String
Dim txtBuffer As String
Dim i As Long

If Not Privilege Then
    'Require Admin privileges
    If Not (LoadPrivilege(SE_DEBUG_NAME)) Then
'        End
    End If
End If
Privilege = True

PortInUse = False
If OpenPort() Then
        'Lists only Processes assigned to Ports
    If ProtoType = TCP Then
        If ResultPorts(0, Port) Then
        PortInUse = True
        End If
    End If
    If ProtoType = UDP Then
        If ResultPorts(1, Port) Then
        PortInUse = True
        End If
    End If
End If
End Function

Public Function ApplicationUsingPort(Port As Integer, ProtoType As ProtocallType) As String
Dim PathBuf As String
Dim txtBuffer As String
Dim i As Long

If Not Privilege Then
    'Require Admin privileges
    If Not (LoadPrivilege(SE_DEBUG_NAME)) Then
'        End
    End If
End If
Privilege = True

If OpenPort() Then
        'Lists only Processes assigned to Ports
    If ProtoType = TCP Then
        If ResultPorts(0, Port) Then
        ApplicationUsingPort = ProcessPathByPID(ResultPorts(0, Port))
        End If
    End If
    If ProtoType = UDP Then
        If ResultPorts(1, Port) Then
        ApplicationUsingPort = ProcessPathByPID(ResultPorts(1, Port))
        End If
    End If
End If
End Function

Private Sub UserControl_Initialize()
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
UserControl.Width = 1935
UserControl.Height = 495
End Sub
