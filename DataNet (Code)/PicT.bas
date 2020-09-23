Attribute VB_Name = "PicT"
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
    '
    ' API structure definition for Point
    '


Public Type POINTAPI
    X As Long
    Y As Long
    End Type
    '
    ' API structure definition for Brush
    '


Public Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
    End Type
    '
    ' API structure definition for Pen
    '


Public Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
    End Type
    '
    ' API function declarations
    '


Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long


Public Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)


Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long


Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long


Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long


Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long


Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Declare Sub InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)


Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long


Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long


Public Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long


Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
    '
    ' Windows messages watched for by MsgHoo
    '     k
    '
    Public Const WM_ERASEBKGND = &H14
    Public Const WM_PAINT = &HF
    Public Const WM_QUERYDRAGICON = &H37
    Public Const WM_WINDOWPOSCHANGED = &H47
    '
    ' Constant used with GetWindow() to obta
    '     in handle
    ' to MDIForm's client space
    '
    Public Const GW_CHILD = 5
    '
    ' Raster-op for Blt's
    '
    Public Const SRCCOPY = &HCC0020
    '
    ' Pen Style constant
    '
    Public Const PS_SOLID = 0


Public Sub mdiBitBltCentered(sWnd As Long, sDC As Long, dWnd As Long)
    Dim nRet As Long
    Dim cDC As Long
    Dim cWnd As Long
    Dim dX As Long
    Dim dY As Long
    Dim sR As RECT
    Dim dR As RECT
    '
    ' Get DC to client space (assumes we're
    '     Blt'ing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(dWnd, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Get source and destination rectangles
    '
    Call GetClientRect(sWnd, sR)
    Call GetClientRect(cWnd, dR)
    '
    ' Calc parameters
    '
    dX = (dR.Right - sR.Right) \ 2
    dY = (dR.Bottom - sR.Bottom) \ 2
    '
    ' Do the BitBlt and clean up
    '
    nRet = BitBlt(Master.Picture3.hDC, dX, dY, sR.Right, sR.Bottom, _
    sDC, 0, 0, SRCCOPY)
    nRet = ReleaseDC(cWnd, cDC)
End Sub


Public Sub mdiBitBltTiled(sWnd As Long, sDC As Long, dWnd As Long)
    Dim nRet As Long
    Dim cDC As Long
    Dim cWnd As Long
    Dim dX As Long
    Dim dY As Long
    Dim Rows As Integer
    Dim Cols As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sR As RECT
    Dim dR As RECT
    '
    ' Get DC to client space (assumes we're
    '     Blt'ing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(dWnd, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Get source and destination rectangles
    '
    Call GetClientRect(sWnd, sR)
    Call GetClientRect(cWnd, dR)
    '
    ' Calc parameters
    '
    Rows = dR.Right \ sR.Right
    Cols = dR.Bottom \ sR.Bottom
    '
    ' Spray out across destination
    '


    For i = 0 To Rows
        dX = i * sR.Right


        For j = 0 To Cols
            dY = j * sR.Bottom
            nRet = BitBlt(cDC, dX, dY, sR.Right, sR.Bottom, _
            sDC, 0, 0, SRCCOPY)
        Next j
    Next i
    '
    ' and clean up
    '
    nRet = ReleaseDC(cWnd, cDC)
End Sub


Public Sub mdiPaintGradient(hWndParent As Long)
    Const Shades% = 64
    Dim cWnd As Long
    Dim cDC As Long
    Dim nRet As Long
    Dim FillBoxHeight As Integer
    Dim NewBrush As Long
    Dim i As Integer
    Dim cRect As RECT
    Static fRect(1 To Shades) As RECT
    '
    ' Get DC to client space (assumes we're
    '     drawing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(hWndParent, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Set up a structure of rectangles for f
    '     ills
    '
    Call GetClientRect(cWnd, cRect)
    FillBoxHeight = cRect.Bottom \ Shades


    For i = 1 To Shades
        fRect(i).Left = cRect.Left
        fRect(i).Right = cRect.Right
        fRect(i).Top = (i - 1) * FillBoxHeight
        fRect(i).Bottom = fRect(i).Top + FillBoxHeight
    Next i
    '
    ' Make up for slop on last one
    '
    fRect(Shades).Bottom = cRect.Bottom
    '
    ' Fill-it-up!
    '


    For i = Shades - 1 To 0 Step -1
        NewBrush = CreateSolidBrush(RGB(0, 0, (i + 1) * 4 - 1))
        nRet = FillRect(cDC, fRect(Shades - i), NewBrush)
        nRet = DeleteObject(NewBrush)
    Next i
    '
    ' and clean up
    '
    nRet = ReleaseDC(cWnd, cDC)
End Sub


Public Sub mdiPaintTunnel1(hWndParent As Long)
    Const Shades% = 64
    Dim cWnd As Long
    Dim cDC As Long
    Dim nRet As Long
    Dim i As Integer
    Dim dX As Long
    Dim dY As Long
    Dim NewBrush As Long
    Dim cRect As RECT
    '
    ' Get DC to client space (assumes we're
    '     drawing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(hWndParent, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Get target dimensions and calculate sh
    '     rinkage factors
    '
    Call GetClientRect(cWnd, cRect)
    dX = cRect.Right / Shades \ 2
    dY = cRect.Bottom / Shades \ 2
    '
    ' Fill-it-up!
    '


    For i = Shades - 1 To 0 Step -1
        NewBrush = CreateSolidBrush(RGB((i + 1) * 4 - 1, 0, 0))
        nRet = FillRect(cDC, cRect, NewBrush)
        nRet = DeleteObject(NewBrush)
        InflateRect cRect, -dX, -dY
    Next i
    '
    ' and clean up
    '
    nRet = ReleaseDC(cWnd, cDC)
End Sub


Public Sub mdiPaintTunnel2(hWndParent As Long)
    Const Shades% = 32
    Dim cWnd As Long
    Dim cDC As Long
    Dim nRet As Long
    Dim i As Integer
    Dim dX As Long
    Dim dY As Long
    Dim NewBrush As Long
    Dim eRgn As Long
    Dim cRect As RECT
    '
    ' Get DC to client space (assumes we're
    '     drawing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(hWndParent, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Get target dimensions and calculate sh
    '     rinkage factors
    '
    Call GetClientRect(cWnd, cRect)
    dX = cRect.Right / Shades / 2
    dY = cRect.Bottom / Shades / 2
    '
    ' Fill background with solid green
    '
    NewBrush = CreateSolidBrush(RGB(0, 255, 0))
    nRet = FillRect(cDC, cRect, NewBrush)
    nRet = DeleteObject(NewBrush)
    '
    ' Fill-it-up!Shades from Green to Black
    '


    For i = Shades - 1 To 0 Step -1
        NewBrush = CreateSolidBrush(RGB(0, (i + 1) * 8 - 8, 0))
        eRgn = CreateEllipticRgn(cRect.Left, cRect.Top, cRect.Right, cRect.Bottom)
        nRet = FillRgn(cDC, eRgn, NewBrush)
        nRet = DeleteObject(NewBrush)
        nRet = DeleteObject(eRgn)
        Call InflateRect(cRect, -dX, -dY)
    Next i
    '
    ' and clean up
    '
    nRet = ReleaseDC(cWnd, cDC)
End Sub


Public Sub mdiStretchBlt(sWnd As Long, sDC As Long, dWnd As Long, Proportional As Boolean)
    Dim nRet As Long
    Dim cDC As Long
    Dim cWnd As Long
    Dim sR As RECT
    Dim dR As RECT
    Dim factor As Single
    Dim dX As Long
    Dim dY As Long
    '
    ' Get DC to client space (assumes we're
    '     Blt'ing
    ' onto an MDI client space)
    '
    cWnd = GetWindow(dWnd, GW_CHILD)
    cDC = GetDC(cWnd)
    '
    ' Get source and destination rectangles
    '
    Call GetClientRect(sWnd, sR)
    Call GetClientRect(cWnd, dR)
    '
    ' Alter destination if proportional to r
    '     espect constraining
    ' dimension
    '


    If Proportional Then


        If dR.Bottom / sR.Bottom < dR.Right / sR.Right Then
            'Height is constraining dimension
            factor! = dR.Bottom / sR.Bottom
            dX = (dR.Right - (factor! * sR.Right)) \ -2
        Else
            'Width is constraining dimension
            factor! = dR.Right / sR.Right
            dY = (dR.Bottom - (factor! * sR.Bottom)) \ -2
        End If
        InflateRect dR, dX, dY
    End If
    '
    ' Stretch out across destination and cle
    '     an up
    '
    nRet = StretchBlt(cDC, dR.Left, dR.Top, CLng(dR.Right - dR.Left), _
    CLng(dR.Bottom - dR.Top), sDC, 0&, 0&, _
    sR.Right, sR.Bottom, SRCCOPY)
    nRet = ReleaseDC(cWnd, cDC)
End Sub



