VERSION 5.00
Begin VB.Form frmTelnet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Telnet Session"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFC0C0&
   Icon            =   "Telnet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   9930
   Begin VB.Timer cursor_timer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3840
      Top             =   1800
   End
End
Attribute VB_Name = "frmTelnet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Const GO_NORM = 0

Const GO_ESC1 = 1
Const GO_ESC2 = 2
Const GO_ESC3 = 3
Const GO_ESC4 = 4
Const GO_ESC5 = 5

Const GO_IAC1 = 6
Const GO_IAC2 = 7
Const GO_IAC3 = 8
Const GO_IAC4 = 9
Const GO_IAC5 = 10
Const GO_IAC6 = 11


Const SUSP = 237
Const ABORT = 238      'Abort
Const SE = 240         'End of Subnegotiation
Const NOP = 241
Const DM = 242         'Data Mark
Const BREAK = 243      'BREAK
Const IP = 244         'Interrupt Process
Const AO = 245         'Abort Output
Const AYT = 246        'Are you there
Const EC = 247         'Erase character
Const EL = 248         'Erase Line
Const GOAHEAD = 249    'Go Ahead
Const SB = 250         'What follows is subnegotiation
Const WILLTEL = 251
Const WONTTEL = 252
Const DOTEL = 253
Const DONTTEL = 254
Const IAC = 255

Const BINARY = 0
Const ECHO = 1
Const RECONNECT = 2
Const SGA = 3
Const AMSN = 4
Const Status = 5
Const TIMING = 6
Const RCTAN = 7
Const OLW = 8
Const OPS = 9
Const OCRD = 10
Const OHTS = 11
Const OHTD = 12
Const OFFD = 13
Const OVTS = 14
Const OVTD = 15
Const OLFD = 16
Const XASCII = 17
Const LOGOUT = 18
Const BYTEM = 19
Const DET = 20
Const SUPDUP = 21
Const SUPDUPOUT = 22
Const SENDLOC = 23
Const TERMTYPE = 24
Const EOR = 25
Const TACACSUID = 26
Const OUTPUTMARK = 27
Const TERMLOCNUM = 28
Const REGIME3270 = 29
Const X3PAD = 30
Const NAWS = 31
Const TERMSPEED = 32
Const TFLOWCNTRL = 33
Const LINEMODE = 34
Const DISPLOC = 35
Const ENVIRON = 36
Const AUTHENTICATION = 37
Const UNKNOWN39 = 39
Const EXTENDED_OPTIONS_LIST = 255
Const RANDOM_LOSE = 256




'------------------------------------------------------------
Private Operating       As Boolean
Private Connected       As Boolean
Public Receiving        As Boolean

Private parsedata(10)   As Integer
Private ppno            As Integer


Private control_on      As Boolean


Public RemoteIPAd  As String
Public RemotePort  As Integer

Public TraceTelnet As Boolean
Public Tracevt100   As Boolean

Private sw_ugoahead As Boolean
Private sw_igoahead As Boolean
Private sw_echo     As Boolean
Private sw_linemode As Boolean
Private sw_termsent As Boolean
Private VTcom As VT100

Private Sub cursor_timer_Timer()

   ' Debug.Print "Timer"
    VTcom.term_DriveCursor

End Sub


Public Sub Initialise()
Set VTcom = New VT100
VTcom.Indexy = Me.Tag
VTcom.term_init
End Sub

Private Sub Form_Load()
ShadeIt Me, Light
End Sub

Private Sub Form_Paint()
 VTcom.term_redrawscreen
End Sub


Public Sub AddText(Text As String)

    Dim CH()     As Byte
    Dim Test()   As Integer
    Dim i        As Integer
    Static cmd   As Byte
'------------------------------------------------------------
    If Not Receiving Then
        Receiving = True
        VTcom.term_CaretControl True
    Else
        Exit Sub
    End If

   
    If (Len(Text) > 0) Then  ' If there is any data...
        
       ' CH = Buf
        For i = 1 To Len(Text)
            Select Case cmd
                Case GO_NORM
                  cmd = VTcom.term_process_char(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC1
                  cmd = iac1(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC2
                  cmd = iac2(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC3
                  cmd = iac3(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC4
                  cmd = iac4(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC5
                  cmd = iac5(Asc(Right(Left(Text, i), 1)))
                Case GO_IAC6
                  cmd = iac6(Asc(Right(Left(Text, i), 1)))
                Case Else
                 If TraceTelnet Then Debug.Print "Invalid 'next (" + Str$(cmd) + ")' processing routine in cmd loop"
            End Select
        Next i
    End If
    
    VTcom.term_CaretControl False
    Receiving = False
End Sub



Private Function iac1(CH As Byte) As Integer

  ' Debug.Print "IAC : ";
  iac1 = GO_NORM

  Select Case CH
    Case DOTEL
      iac1 = GO_IAC2
    Case DONTTEL
      iac1 = GO_IAC6
    Case WILLTEL
      iac1 = GO_IAC3
    Case WONTTEL
      iac1 = GO_IAC4
    Case SB
      iac1 = GO_IAC5
      ppno = 0
    Case SE
      ' End of negotiation string, string is in parsedata()
      Select Case parsedata(0)
        Case TERMTYPE
          If parsedata(1) = 1 Then
               If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
          End If
        Case TERMSPEED
          If parsedata(1) = 1 Then
                ' Debug.Print "TERMSPEED"
                If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 38400"
          End If
      End Select
  End Select

End Function

Private Function iac2(CH As Byte) As Integer

  'DO Processing Respond with WILL or WONT

  If TraceTelnet Then Debug.Print "                                                                   RECEIVED DO : ";
  iac2 = GO_NORM

  Select Case CH
    Case BINARY
        If TraceTelnet Then Debug.Print "BINARY"
        If TraceTelnet Then Debug.Print "SENT: WONT BINARY"
    Case ECHO
        If TraceTelnet Then Debug.Print "ECHO"
        If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
    Case NAWS
        If TraceTelnet Then Debug.Print "WINDOW SIZE"
        If TraceTelnet Then Debug.Print "SENT: SB WINDOW SIZE 80x24"
    Case SGA
        If TraceTelnet Then Debug.Print "SGA"
        If Not sw_igoahead Then
            If TraceTelnet Then Debug.Print "SENT: WILL SGA"
            sw_igoahead = True
        Else
           If TraceTelnet Then Debug.Print "DID NOT RESPOND"
        End If
    Case TERMTYPE
        If TraceTelnet Then Debug.Print "TERMTYPE"
        If Not sw_termsent Then
            If TraceTelnet Then Debug.Print "SENT: WILL TERMTYPE"
              sw_termsent = True
            If TraceTelnet Then Debug.Print "SENT: SB TERMTYPE VT100"
         Else
            If TraceTelnet Then Debug.Print "DID NOT RESPOND"
         End If
 
    Case TERMSPEED
        If TraceTelnet Then Debug.Print "TERMSPEED"
        If TraceTelnet Then Debug.Print "SENT: WILL TERMSPEED"
      
    If TraceTelnet Then Debug.Print "SENT: SB TERMSPEED 57600"
      
    Case TFLOWCNTRL
        If TraceTelnet Then Debug.Print "TFLOWCNTRL"
        If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
      
    Case LINEMODE
        If TraceTelnet Then Debug.Print "LINEMODE"
        If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
      
    Case Status
        If TraceTelnet Then Debug.Print "STATUS"
        If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
      
    Case TIMING
        If TraceTelnet Then Debug.Print "TIMING"
        If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
      
    Case DISPLOC
        If TraceTelnet Then Debug.Print "DISPLOC"
        If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
    
    Case ENVIRON
        If TraceTelnet Then Debug.Print "ENVIRON"
        If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
    
    Case UNKNOWN39
        If TraceTelnet Then Debug.Print "UNKNOWN39"
        If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
    
    Case AUTHENTICATION
        If TraceTelnet Then Debug.Print "AUTHENTICATION"
        If TraceTelnet Then Debug.Print "SENT: WILL "; AUTHENTICATION; ""
      
        If TraceTelnet Then Debug.Print "SENT: SB AUTHENTICATION"
    Case Else
        If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
        If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & CH
  End Select

End Function

Private Function iac3(CH As Byte) As Integer

  ' WILL Processing - Respond with DO or DONT
  
If TraceTelnet Then Debug.Print "                                                                   RECEIVED WILL : ";

  iac3 = GO_NORM

  Select Case CH
    Case ECHO
    If TraceTelnet Then Debug.Print "ECHO"
      If Not sw_echo Then
        sw_echo = True
      If TraceTelnet Then Debug.Print "SENT: DO ECHO"
      End If
    Case SGA
    If TraceTelnet Then Debug.Print "SGA"
      If Not sw_ugoahead Then
        sw_ugoahead = True
      If TraceTelnet Then Debug.Print "SENT: DOTEL SGA"
      End If
    
    Case TERMSPEED
    If TraceTelnet Then Debug.Print "TERMSPEED"
    If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
      
    Case TFLOWCNTRL
    If TraceTelnet Then Debug.Print "TFLOWCNTRL"
    If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
      
    Case LINEMODE
    If TraceTelnet Then Debug.Print "LINEMODE"
    If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
      
    Case Status
    If TraceTelnet Then Debug.Print "STATUS"
    If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
      
    Case TIMING
    If TraceTelnet Then Debug.Print "TIMING"
    If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
      
    Case DISPLOC
    If TraceTelnet Then Debug.Print "DISPLOC"
    If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
    
    Case ENVIRON
    If TraceTelnet Then Debug.Print "ENVIRON"
    If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
    
    Case UNKNOWN39
    If TraceTelnet Then Debug.Print "UNKNOWN39"
    If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
    
    
    Case Else
    If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
    If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
  End Select

End Function

Private Function iac4(CH As Byte) As Integer

  ' WONT Processing
  
    If TraceTelnet Then Debug.Print "                                                                   RECEIVED WONT : ";

  iac4 = GO_NORM

  Select Case CH
    
    Case ECHO
    If TraceTelnet Then Debug.Print "ECHO"
      If sw_echo = True Then
      If TraceTelnet Then Debug.Print "SENT: DONTEL ECHO"
        sw_echo = False
      End If
      
    Case SGA
    If TraceTelnet Then Debug.Print "SGA"
    If TraceTelnet Then Debug.Print "SENT: DONT SGA"
      sw_igoahead = False
    
    Case TERMSPEED
    If TraceTelnet Then Debug.Print "TERMSPEED"
    If TraceTelnet Then Debug.Print "SENT: DONT TERMSPEED"
    
    Case TFLOWCNTRL
    If TraceTelnet Then Debug.Print "FLOWCONTROL"
    If TraceTelnet Then Debug.Print "SENT: DONT FLOWCONTROL"
      
    Case LINEMODE
    If TraceTelnet Then Debug.Print "LINEMODE"
    If TraceTelnet Then Debug.Print "SENT: DONT LINEMODE"
      
    Case Status
    If TraceTelnet Then Debug.Print "STATUS"
    If TraceTelnet Then Debug.Print "SENT: DONT STATUS"
      
    Case TIMING
    If TraceTelnet Then Debug.Print "TIMING"
    If TraceTelnet Then Debug.Print "SENT: DONT TIMING"
      
    Case DISPLOC
    If TraceTelnet Then Debug.Print "DISPLOC"
    If TraceTelnet Then Debug.Print "SENT: DONT DISPLOC"
    
    Case ENVIRON
    If TraceTelnet Then Debug.Print "ENVIRON"
    If TraceTelnet Then Debug.Print "SENT: DONT ENVIRON"
    
    Case UNKNOWN39
    If TraceTelnet Then Debug.Print "UNKNOWN39"
    If TraceTelnet Then Debug.Print "SENT: DONT " & Asc(CH)
    
    Case Else
    If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
    If TraceTelnet Then Debug.Print "SENT: DONT UNKNOWN CMD " & Asc(CH)
  End Select

End Function

Private Function iac5(CH As Byte) As Integer

Dim ich As Integer
  ' Collect parms after SB and until another IAC

  
    ich = CH
    If ich = IAC Then
      iac5 = GO_IAC1
      Exit Function
    End If
    
    If TraceTelnet Then Debug.Print "                                                                   RECEIVED : ";
    If TraceTelnet Then Debug.Print "SB("; ppno; ") = " & ich
    
    parsedata(ppno) = ich
    ppno = ppno + 1
    
    iac5 = GO_IAC5

End Function


Private Function iac6(CH As Byte) As Integer

  'DONT Processing

 
  iac6 = GO_NORM
        

  Select Case CH
    Case SE
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED SE"
      If TraceTelnet Then Debug.Print "SENT: SE_ACK " & CH

    Case ECHO
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "ECHO"
      If Not sw_echo Then
        sw_echo = True
        If TraceTelnet Then Debug.Print "SENT: WONT ECHO"
      End If
    Case SGA
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "SGA"
      If Not sw_ugoahead Then
        sw_ugoahead = True
        If TraceTelnet Then Debug.Print "SENT: WONT SGA"
      End If
    
    Case TERMSPEED
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TERMSPEED"
      If TraceTelnet Then Debug.Print "SENT: WONT TERMSPEED"
      
    Case TFLOWCNTRL
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TFLOWCNTRL"
      If TraceTelnet Then Debug.Print "SENT: WONT FLOWCONTROL"
      
    Case LINEMODE
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "LINEMODE"
      If TraceTelnet Then Debug.Print "SENT: WONT LINEMODE"
      
    Case Status
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "STATUS"
      If TraceTelnet Then Debug.Print "SENT: WONT STATUS"
      
    Case TIMING
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "TIMING"
      If TraceTelnet Then Debug.Print "SENT: WONT TIMING"
      
    Case DISPLOC
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "DISPLOC"
      If TraceTelnet Then Debug.Print "SENT: WONT DISPLOC"
    
    Case ENVIRON
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "ENVIRON"
      If TraceTelnet Then Debug.Print "SENT: WONT ENVIRON"
    
    Case UNKNOWN39
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "UNKNOWN39"
      If TraceTelnet Then Debug.Print "SENT: WONT " & Asc(CH)
        
    Case Else
      If TraceTelnet Then Debug.Print "                                                                   RECEIVED DONT : ";
      If TraceTelnet Then Debug.Print "UNKNOWN CMD " & Asc(CH)
      If TraceTelnet Then Debug.Print "SENT: WONT UNKNOWN CMD " & Asc(CH)
  End Select

End Function

