VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7440
      Top             =   2160
   End
   Begin VB.Image Image1 
      Height          =   3495
      Left            =   0
      Picture         =   "Splash.frx":0000
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
Call Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Load Master
DoEvents
    If UCase(Trim(Command$)) = "HIDE" Then
    Master.Visible = False
    Else
    Master.Visible = True
    Master.Show
    End If
Unload Me
End Sub
