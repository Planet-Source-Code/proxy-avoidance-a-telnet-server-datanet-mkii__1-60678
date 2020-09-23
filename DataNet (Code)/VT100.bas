Attribute VB_Name = "VT100"
Option Explicit

'Windows RECT structure
Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    bottom  As Long
End Type


Private Declare Function ScrollWindow Lib "user32" (ByVal hWnd As Long, ByVal XAmount As Long, ByVal YAmount As Long, lpRect As RECT, lpClipRect As RECT) As Long
Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hRgnUpdate As Long, lprcUpdate As RECT) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nwidth As Long, ByVal nheight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetBkMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal newcolor As Long) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long


'=================== Ternary raster operations ============
Private Const PATCOPY = &HF00021         ' (DWORD) dest = pattern
Private Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Private Const BLACKNESS = &H42&          ' (DWORD) dest = BLACK
Private Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Private Const TRANSPARENT = 1
Private Const OPAQUE = 2
Private Const GO_IAC1 = 6



Private Const LinesPerPage = 25
Private Const CharsPerLine = 80
Private Const TabsPerPage = 20

Private Const LastLine = LinesPerPage - 1
Private Const LastChar = CharsPerLine - 1
Private Const LastTab = 19

Private ScrImage(LinesPerPage)    As String * CharsPerLine
Private ScrAttr(LinesPerPage)     As String * CharsPerLine
Private Norm_Attr                 As String * CharsPerLine
Private Blank_Line                As String * CharsPerLine

Private TermTextColor             As Long
Private TermBkColor               As Long

Private tabno                     As Integer
Private tab_table(TabsPerPage)    As Integer
Private curattr                   As String

Private lprcScroll                As RECT
Private lprcClip                  As RECT
Private hRgnUpdate                As Integer
Private lprcUpdate                As RECT


'
'   Current Buffered Text waiting for output on screen
'

Private OutStr          As String
Private outlen          As Integer

'
'   Flag to indicate that we're ready to run
'
Private FlagInit        As Integer

Private CurX            As Integer
Private CurY            As Integer
Private SavecurX        As Integer
Private SavecurY        As Integer

Private InEscape        As Boolean    ' Processing an escape seq?
Private EscString       As String     ' String so far

Private charheight      As Single
Private charWidth       As Single

Private CurState        As Boolean
Private Ret             As Long



Public Function term_process_char(CH As Byte)

       
    If (InEscape) Then
        
        Call term_escapeProcess(CH)
    
    Else
        
        Select Case CH

        Case 0

        Case 7

            Beep

        Case 8


            If CurX > 0 Then                    '   if not at line begin
                CurX = CurX - 1                 '   adjust back 1 spc
            End If

        Case 9
            Dim tY As Integer
            For tY = 0 To 19
              If CurY < tab_table(tY) Then
                Exit For
              End If
            Next tY
            CurY = tab_table(tY)

        Case 10, 11, 12

            If (CurY = LastLine) Then           '   if line left on scrn
                Call term_scroll_up             '   ..  scroll upwards
                CurY = LastLine                 '   ..  use blank line
            Else
                CurY = CurY + 1                 '   goto next line
            End If

        Case 13
        
            CurX = 0
            
        Case 27

            InEscape = True
            EscString = ""

        Case 255
            
            term_process_char = GO_IAC1
        
        Case Else
            
          ' if (CH > 31) Then ' And (CH < 128)
                term_write CH
                Mid$(ScrImage(CurY + 1), CurX, 1) = Chr$(CH)
                Mid$(ScrAttr(CurY + 1), CurX, 1) = curattr
           ' End If

        End Select
        
    End If
End Function
Public Sub term_CaretControl(TurnOff As Boolean)
Static SaveState As Boolean

    If TurnOff = True Then
        SaveState = CurState
        term_Carethide
    Else
        If SaveState = True Then
            term_Caretshow
        End If
    End If
    
End Sub
Private Sub term_Carethide()

    If CurState = True Then
        If frmTelnet.WindowState <> 1 Then
            Ret = PatBlt(frmTelnet.hdc, CurX * charWidth, CurY * charheight, charWidth, charheight, DSTINVERT)
        End If
        CurState = False
    End If
End Sub

Private Sub term_Caretshow()

    '------------------------------------------------------------------------
    '   term_CaretShow
    '
    '   display the inverted block cursor on the screen.
    '   currently uses PatBlt.
    '------------------------------------------------------------------------
    Dim Ret As Integer

    If frmTelnet.WindowState <> 1 Then
       Ret = PatBlt(frmTelnet.hdc, CurX * charWidth, CurY * charheight, charWidth, charheight, DSTINVERT)
    End If

    CurState = True

End Sub
Public Sub term_DriveCursor()
        Call term_Carethide
End Sub
Private Sub term_eraseBOL()
'------------------------------------------------------------------------
'   term_eraseBOL
'   erase from beginning of current line
'------------------------------------------------------------------------
    Dim Ret As Integer

    If frmTelnet.WindowState <> 1 Then
       ' Ret = PatBlt(frmTelnet.hdc, 0, CurY * charheight, curX * charWidth, charheight, BLACKNESS)
        Ret = TextOut(frmTelnet.hdc, 0, CurY * charheight, Blank_Line, CharsPerLine)
        
    End If

    Mid$(ScrImage(CurY + 1), 1, CurX + 1) = Space$(CurX + 1)
    Mid$(ScrAttr(CurY + 1), 1, CurX + 1) = String$(CurX + 1, "0")
End Sub

Private Sub term_eraseBOS()
'------------------------------------------------------------------------
'   term_eraseBOS
'   erase all lines from beginning of screen to and including current
'------------------------------------------------------------------------
    Dim Y As Integer

    'Erase the current line first
    Call term_eraseBOL

    'Erase everything up to current line
    If (CurY > 0) Then
        If frmTelnet.WindowState <> 1 Then
            Ret = TextOut(frmTelnet.hdc, 0, 0, Space$(CharsPerLine * CurY + CurX), CharsPerLine * CurY + CurX)
            
        End If

        ' reset screen buffer contents
        For Y = 1 To CurY
           ScrImage(Y) = Blank_Line
           ScrAttr(Y) = Norm_Attr
        Next Y
    End If
End Sub

Private Sub term_eraseBUFFER()
    Dim i As Integer
    For i = 1 To LinesPerPage
        ScrImage(i) = Blank_Line
        ScrAttr(i) = Norm_Attr
    Next i
End Sub

Private Sub term_eraseEOL()
'
'   Erase to End of Line
'
    If frmTelnet.WindowState <> 1 Then
        Ret = TextOut(frmTelnet.hdc, CurX * charWidth, CurY * charheight, Space$(CharsPerLine - CurX), CharsPerLine - CurX)
    End If

    'Update screen buffer
    Mid$(ScrImage(CurY + 1), CurX + 1, CharsPerLine - CurX) = Space$(CharsPerLine - CurX)
    Mid$(ScrAttr(CurY + 1), CurX + 1, CharsPerLine - CurX) = String$(CharsPerLine - CurX, "0")

End Sub

Private Sub term_eraseEOS()
'
'   Erase to end of screen
'
    Dim Y As Integer

    Call term_eraseEOL
    If (CurY <> LastLine) Then

        If frmTelnet.WindowState <> 1 Then
            Ret = TextOut(frmTelnet.hdc, 0, (CurY + 1) * charheight, Space$((LastLine - CurY) * CharsPerLine), (LastLine - CurY) * CharsPerLine)
        End If

        For Y = CurY + 2 To LinesPerPage
            ScrImage(Y) = Blank_Line
            ScrAttr(Y) = Norm_Attr
        Next Y

     End If
End Sub

Private Sub term_eraseLINE()

'   Erase Line

    If frmTelnet.WindowState <> 1 Then
        Ret = TextOut(frmTelnet.hdc, 0, CurY * charheight, Blank_Line, CharsPerLine)
    End If

    ScrImage(CurY + 1) = Blank_Line
    ScrAttr(CurY + 1) = Norm_Attr

End Sub

Private Sub term_eraseSCREEN()

    'Assume that they want to repaint using the latest background color
    
    TermBkColor = GetBkColor(frmTelnet.hdc)
    TermTextColor = GetTextColor(frmTelnet.hdc)
    
    frmTelnet.BackColor = TermBkColor
    frmTelnet.ForeColor = TermTextColor

    
    term_eraseBUFFER
    frmTelnet.Cls
    CurX = 0
    CurY = 0

End Sub

Private Function term_escapeParseArg(s As String) As String
'
'   PopArg takes the next argument (digits up to a ;) and
'   returns it.  It also removes the arg and the ; from
'   the "s"

    Dim i As Integer

    i = InStr(s, ";")
    If i = 0 Then
        term_escapeParseArg = s
        s = ""
    Else
        term_escapeParseArg = Left$(s, i - 1)
        s = Mid$(s, i + 1)
    End If

End Function

Private Sub term_escapeProcess(CH As Byte)

Dim c           As String
Dim yDiff       As Integer
Dim xDiff       As Integer


    c = Chr$(CH)
    If EscString = "" Then
      'No start character yet
      Select Case c
        Case "["
        
        Case "("
        
        Case ")"
        
        Case "#"
        
        Case Chr$(8)             ' embedded backspace
          CurX = CurX - 1
          term_validatecurX
          InEscape = False
        
        Case "7"                 ' save cursor
          'Save cursor position
          SavecurX = CurX
          SavecurY = CurY
          InEscape = False
        
        Case "8"                 ' restore cursor
          'restore cursor position
          CurX = SavecurX
          CurY = SavecurY
          InEscape = False
        
        Case "c"                 ' look at VSIreset()
        
        Case "D"                 ' cursor down
          CurY = CurY + 1
          term_validatecurY
          InEscape = False
        
        Case "E"                 ' next line
          CurY = CurY + 1
          CurX = 0
          term_validatecurY
          term_validatecurX
          InEscape = False
        
        Case "H"                 ' set tab
          tab_table(tabno) = CurY
          tabno = tabno + 1
          InEscape = False
        
        Case "I"                 ' look at bp_ESC_I()
          InEscape = False
        
        Case "M"                 ' cursor up
          CurY = CurY - 1
          term_validatecurY
                
        Case "Z"                 ' send ident
          InEscape = False
        
        Case Else
              'Invalid start of escape sequence
            If frmTelnet.Tracevt100 Then Debug.Print CH
            
            InEscape = False
            Exit Sub
      End Select
    End If

    EscString = EscString & c
    If IsCharAlpha(CH) = 0 Then
        ' Not a character ...
        If Len(EscString) > 15 Then
          If frmTelnet.Tracevt100 Then Debug.Print CH
            InEscape = False
        End If
        Exit Sub
    End If


    Select Case c

        Case "A"

            ' A ==> move cursor up
            
            EscString = Mid$(EscString, 2)

            yDiff = Val(term_escapeParseArg(EscString))
            If yDiff = 0 Then
                yDiff = 1
            End If

            CurY = CurY - yDiff
            term_validatecurY
        
        Case "B"

            ' B ==> move cursor down
            
            EscString = Mid$(EscString, 2)

            yDiff = Val(term_escapeParseArg(EscString))
            If yDiff = 0 Then
                yDiff = 1
            End If

            CurY = CurY + yDiff
            term_validatecurY

        Case "C"
            ' C ==> move cursor right

            EscString = Mid$(EscString, 2)

            xDiff = Val(term_escapeParseArg(EscString))
            If xDiff = 0 Then
                xDiff = 1
            End If

            CurX = CurX + xDiff
            term_validatecurX
        
        Case "D"
            ' D ==> move cursor left

            EscString = Mid$(EscString, 2)

            xDiff = Val(term_escapeParseArg(EscString))
            If xDiff = 0 Then
                xDiff = 1
            End If
            CurX = CurX - xDiff
            term_validatecurX
        
        Case "H"

            'Goto cursor position indicated by escape sequence

            EscString = Mid$(EscString, 2)

            CurY = Val(term_escapeParseArg(EscString)) - 1
            term_validatecurY

            CurX = Val(EscString) - 1
            term_validatecurX

        Case "J"

            'Erase display

            Select Case Val(Mid$(EscString, 2))

                Case 0
                    If CurX = 0 And CurY = 0 Then
                        Call term_eraseSCREEN
                    Else
                        Call term_eraseEOS
                    End If

                Case 1
                    Call term_eraseBOS

                Case 2
                    Call term_eraseSCREEN

            End Select

        Case "K"

            'Erase line
            Select Case Val(Mid$(EscString, 2))
                Case 0
                    'erase to end of line
                    Call term_eraseEOL
                Case 1
                    'erase to end of line
                    Call term_eraseBOL
                Case 2
                    Call term_eraseLINE
            End Select

        Case "f"

            'Goto cursor position indicated by escape sequence

            EscString = Mid$(EscString, 2)

            CurY = Val(term_escapeParseArg(EscString)) - 1
            term_validatecurY

            CurX = Val(EscString) - 1
            term_validatecurX
        
        Case "g"
            ' clear tabs
            
            Dim tY As Integer
            For tY = 0 To 19
              tab_table(tY) = 0
            Next tY
        
        Case "h"

            'restore cursor position
            CurX = SavecurX
            CurY = SavecurY

        Case "i"
            ' print though mode
        
        Case "l"
            'Save cursor position
            SavecurX = CurX
            SavecurY = CurY

        Case "m"

            'Change text attributes, screen colors
            
            EscString = Mid$(EscString, 2)
            Do
                Call term_setattr(Chr$(Val(term_escapeParseArg(EscString))))
            Loop While EscString <> ""

        Case "r"
            
            'Set scrollable region
            EscString = Mid$(EscString, 2)

            lprcScroll.Top = (Val(term_escapeParseArg(EscString)) - 1) * charheight
            lprcClip = lprcScroll
        
        Case "s"
            'Save cursor position
            SavecurX = CurX
            SavecurY = CurY

        Case "u"

            'restore cursor position
            CurX = SavecurX
            CurY = SavecurY


        Case Else

          If frmTelnet.Tracevt100 Then Debug.Print EscString

    End Select

    InEscape = False
    EscString = ""

End Sub

Public Sub term_init()

    'Get the pixel metrics of the current font
    frmTelnet.FontUnderline = False
    frmTelnet.FontItalic = False
    frmTelnet.FontBold = False
    
    frmTelnet.ScaleMode = 3
    charheight = frmTelnet.TextHeight("M")
    charWidth = frmTelnet.TextWidth("M")



    'Set the user scale of the display
    frmTelnet.ScaleMode = 0
    frmTelnet.ScaleWidth = LinesPerPage
    frmTelnet.ScaleWidth = CharsPerLine
    frmTelnet.Scale (0, 0)-(LastChar, LastLine)

    'Set up the scoll region and clip region structures
    lprcScroll.Top = 0
    lprcScroll.Left = 0
    lprcScroll.Right = CharsPerLine * charWidth
    lprcScroll.bottom = LinesPerPage * charheight
    lprcClip = lprcScroll
    hRgnUpdate = 0

    'Initialize module level flags and variables
    InEscape = False
    CurState = False
    curattr = "0"
    CurX = 0
    CurY = 0

    'Set the default foreground and background colors
    Ret = SetBkMode(frmTelnet.hdc, OPAQUE)
    frmTelnet.ForeColor = QBColor(15)
    frmTelnet.BackColor = QBColor(0)
    Ret = SetBkColor(frmTelnet.hdc, frmTelnet.BackColor)
    Ret = SetTextColor(frmTelnet.hdc, frmTelnet.ForeColor)

    TermTextColor = GetTextColor(frmTelnet.hdc)
    TermBkColor = GetBkColor(frmTelnet.hdc)


    'Initialize repaint buffer
    Norm_Attr = String$(CharsPerLine, "0")
    Blank_Line = Space$(CharsPerLine)
    term_eraseBUFFER

    FlagInit = True

    'Do the cursor thing
    term_Caretshow
    frmTelnet.cursor_timer.Enabled = True

End Sub
Private Function Term_FindChange(InArray As String, ByVal CurrentValue As String, ByteLen As Integer) As Integer
On Error GoTo endi
Dim RetValue As Integer
Dim CurrentByte As Byte
Dim InByte() As Byte

CurrentByte = CurrentValue
InByte = InArray

For RetValue = 1 To ByteLen
    If InByte(RetValue) <> CurrentByte Then
        Exit For
    End If
Next

Term_FindChange = RetValue - 1
endi
End Function
Public Sub term_redrawscreen()

    If Not FlagInit Or frmTelnet.WindowState = 1 Then
        Exit Sub
    End If

    Dim oldcur      As Boolean
    Dim oldattr     As String
    Dim newattr     As String
    Dim Y           As Integer
    Dim X1          As Integer
    Dim X2          As Integer
    Dim AttrChange  As Integer
    Dim tAttr       As String * CharsPerLine
    Dim tLine       As String * CharsPerLine
    
    
    oldcur = CurState
    oldattr = curattr

    If Not frmTelnet.Receiving Then
        Call term_Carethide
    End If

    Call term_setattr("0")

    For Y = 1 To LinesPerPage
        tAttr = ScrAttr(Y)
        tLine = ScrImage(Y)
        If (tAttr = Norm_Attr) Then
            'Normal Lines can be repainted directly
            Ret = TextOut(frmTelnet.hdc, 0, (Y - 1) * charheight, tLine, CharsPerLine)
        Else
            'Complex lines must have attribute changes found using the
            'Term_function FindChange.
            X1 = 1                          'Start the scan on the complete line
            X2 = CharsPerLine
            Do While (X2 > 0)
                AttrChange = Term_FindChange(Mid(tAttr, X1, X2), curattr, X2)
                Ret = TextOut(frmTelnet.hdc, (X1 - 1) * charWidth, (Y - 1) * charheight, Mid$(tLine, X1, AttrChange), AttrChange)
                X2 = X2 - AttrChange
                If X2 > 0 Then
                    X1 = X1 + AttrChange
                    newattr = Mid$(tAttr, X1, 1)
                    If newattr <> "0" Then
                        term_setattr newattr
                    Else
                        curattr = newattr
                    End If
                End If
            Loop
        End If
    Next Y


    Call term_setattr(oldattr)
    If Not frmTelnet.Receiving Then
        If oldcur = True Then
            Call term_Caretshow
        End If
    End If
    

End Sub

Private Sub term_scroll_up()

    Dim i As Integer
    Dim s As Integer

    If frmTelnet.WindowState <> 1 Then
         Ret = ScrollDC(frmTelnet.hdc, 0, -charheight, lprcScroll, lprcClip, hRgnUpdate, lprcUpdate)
         Ret = TextOut(frmTelnet.hdc, 0, CurY * charheight, Blank_Line, CharsPerLine)
    End If

    'Update the redisplay buffer (only update the scrollable region)
    'Might consider making this a circular array so only one line
    'needs to be written per scroll, rather than relinking the array
    s = ((lprcScroll.Top + 1) \ (charheight + 1))
    For i = s To LastLine
        ScrImage(i) = ScrImage(i + 1)
        ScrAttr(i) = ScrAttr(i + 1)
    Next i
    ScrImage(LinesPerPage) = Blank_Line
    ScrAttr(LinesPerPage) = Norm_Attr


End Sub

Private Sub term_setattr(CH As String)
Dim Attr_BitMap As Integer

    Select Case Asc(CH)

            Case 0  '   Normal
               ' Attr_BitMap = Attr_Norm
                
                frmTelnet.FontUnderline = False
                frmTelnet.FontItalic = False
                frmTelnet.FontBold = False
                Ret = SetTextColor(frmTelnet.hdc, TermTextColor)
                Ret = SetBkColor(frmTelnet.hdc, TermBkColor)

            Case 1  '   Bold
               ' Attr_BitMap = Attr_BitMap And Attr_Norm
                frmTelnet.FontBold = True
'                Ret = SetTextColor(frmTelnet.hdc, QBColor(9))

            Case 5  '   Blinking
               ' Attr_BitMap = Attr_BitMap And Attr_Blink
                frmTelnet.FontItalic = True
'                Ret = SetTextColor(frmTelnet.hdc, QBColor(3))

            Case 4  '   Underscore
               ' Attr_BitMap = Attr_BitMap And Attr_Under
                frmTelnet.FontUnderline = True

            Case 7  '   Reverse Video
               ' Attr_BitMap = Attr_BitMap And ATTR_REVERSE
                Ret = SetTextColor(frmTelnet.hdc, TermBkColor)
                Ret = SetBkColor(frmTelnet.hdc, TermTextColor)

            Case 8  '   Cancel (Invisible)
                'Attr_BitMap = Attr_BitMap And ATTR_INVISIBLE
                Ret = SetTextColor(frmTelnet.hdc, TermBkColor)
                Ret = SetBkColor(frmTelnet.hdc, TermBkColor)

            '===============================================================

            Case 30 '   Black Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(0))

            Case 31 '   Red Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(4))

            Case 32 '   Green Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(2))

            Case 33 '   Yellow Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(14))

            Case 34 '   Blue Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(1))

            Case 35 '   Magenta Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(5))

            Case 36 '   Cyan Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(3))

            Case 37 '   White Foreground
                Ret = SetTextColor(frmTelnet.hdc, QBColor(15))

            '===============================================================

            Case 40 '   Black Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(0))

            Case 41 '   Red Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(4))

            Case 42 '   Green Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(2))

            Case 43 '   Yellow Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(14))

            Case 44 '   Blue Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(1))

            Case 45 '   Magenta Background
               Ret = SetBkColor(frmTelnet.hdc, QBColor(5))

            Case 46 '   Cyan Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(3))

            Case 47 '   White Background
                Ret = SetBkColor(frmTelnet.hdc, QBColor(15))

            Case Else
                Exit Sub
    End Select

    curattr = CH
End Sub

Private Sub term_validatecurX()
   If (CurX < 0) Then
        CurX = 0
   ElseIf CurX > LastChar Then
        CurX = LastChar
   End If
End Sub

Private Sub term_validatecurY()
   If (CurY < 0) Then
        CurY = 0
   ElseIf CurY > LastLine Then
        CurY = LastLine
   End If
End Sub

Private Sub term_write(CH As Byte)

    If frmTelnet.WindowState <> 1 Then
        Ret = TextOut(frmTelnet.hdc, CurX * charWidth, CurY * charheight, Chr$(CH), 1)
    End If

    If Not (CurX = LastChar) Then
        CurX = CurX + 1
    End If

End Sub


