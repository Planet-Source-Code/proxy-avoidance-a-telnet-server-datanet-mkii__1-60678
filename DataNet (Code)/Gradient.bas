Attribute VB_Name = "Gradient"
Private mGradient       As New clsGradient
Public Sub DrawGradient(Picture As Object, Angle As Single, Colour1 As Long, Colour2 As Long)
    With mGradient
        .Angle = Angle
        .Color1 = Colour1
        .Color2 = Colour2
        .Draw Picture
    End With
'Picture.Refresh
End Sub
Public Sub DrawGradient2(Picture As Object, Angle As Single, Colour1 As Long, Colour2 As Long)
    With mGradient
        .Angle = Angle
        .Color1 = Colour1
        .Color2 = Colour2
        .Draw Picture
    End With
'Picture.Refresh
End Sub
