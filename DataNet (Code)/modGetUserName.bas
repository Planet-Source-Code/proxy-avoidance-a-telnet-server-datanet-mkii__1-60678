Attribute VB_Name = "modGetUserName"
'Written by Chris Pietschmann
'http://PietschSoft.itgo.com

Public Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long)

Public Function UserName()
    Dim strUserName As String 'Declare the buffer to hole the username
    Dim lngSize As Long 'This holds to size of the buffer
    Dim blnStatus As Boolean 'Declare variable to get success status
    
    lngSize = 255 'initialize the size of the buffer
    strUserName = Space(lngSize) 'initialize the buffer that will hold the username
    
    blnStatus = GetUserName(strUserName, lngSize) 'The actually API call to get the username
    'After the API call the variable lngSize contains the length of the Username returned by the call into the variable strUserName.
    
    If blnStatus = False Then
        'GetUserName will return False if it fails
        MsgBox "The Call Failed!"
        Exit Function
    End If
    
    'After the call has been made lngSize will contain the length of the username returned
    UserName = Left(strUserName, lngSize - 1) 'return the username from the function
    
End Function
