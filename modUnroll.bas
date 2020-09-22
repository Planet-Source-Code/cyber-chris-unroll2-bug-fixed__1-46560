Attribute VB_Name = "modEffects"
'// (c) Copyright by CC
'//     Email: cyber_chris235@gmx.net

Public Sub SimpleUnroll(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Cls
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, 0, 1, Source.ScaleHeight, Offset - TubeWidth, 0, 1, Source.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube

    Next Offset

End Sub

Public Sub Lens2(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Picture = Nothing
    Target.Picture = Source.Picture
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, 0, 1, Source.ScaleHeight, Offset - TubeWidth, 0, 1, Source.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = XTube / (TubeWidth / 2) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube

    Next Offset
End Sub

Public Sub SlowUnroll(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double, x
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, 0, 1, Source.ScaleHeight, Offset - TubeWidth, 0, 1, Source.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                            For x = 1 To 20
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
                            Next x
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube

    Next Offset

End Sub
Public Sub NeedlePrinterDownUp(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double, x
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.5083185307 / (TubeWidth * 50)
    For x = 1 To Source.ScaleHeight / 50
    For Offset = 0 To Source.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, Source.ScaleHeight - x * 50, 1, 50, Offset - TubeWidth, Source.ScaleHeight - x * 50, 1, 50
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 50)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, Source.ScaleHeight - x * 50, 1, 50, Offset + XPicture, Source.ScaleHeight - x * 50, 1, 50
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, Source.ScaleHeight - x * 50, 1, 50, Offset + XTube - TubeWidth, Source.ScaleHeight - x * 50, 1, 50
            End If
        Next XTube

    Next Offset
    Next x
End Sub

Public Sub NeedlePrinter(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double, x
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.5083185307 / (TubeWidth * 50)
    For x = 0 To Source.ScaleHeight / 50
    For Offset = 0 To Source.ScaleWidth - 1
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, x * 50, 1, 50, Offset - TubeWidth, x * 50, 1, 50
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 50)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, x * 50, 1, 50, Offset + XPicture, x * 50, 1, 50
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, x * 50, 1, 50, Offset + XTube - TubeWidth, x * 50, 1, 50
            End If
        Next XTube

    Next Offset
    Next x
End Sub

Public Sub UnrollWithAfterEffect(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1 Step 3
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, 0, 1, Source.ScaleHeight, Offset - TubeWidth, 0, 1, Source.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube
    Next Offset
End Sub
Public Sub CutThePic(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1 Step TubeWidth
        If Offset - TubeWidth >= 0 Then Target.PaintPicture Source.Picture, Offset - TubeWidth, 0, 1, Source.ScaleHeight, Offset - TubeWidth, 0, 1, Source.ScaleHeight
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube
    Next Offset
End Sub
Public Sub UnrollWithMirror(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Picture = Nothing
    Target.Cls
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube
    Next Offset
End Sub

Public Sub UnrollWithLens(Source As PictureBox, Target As PictureBox, Optional TubeWidth = 50)
    Dim XTube As Long, Offset As Long, XPicture As Long, RDS As Double
    Target.Picture = Nothing
    Target.Picture = Source.Picture
    RDS = 6.283185307 / (TubeWidth * 2)
    For Offset = 0 To Source.ScaleWidth - 1
        For XTube = 1 To TubeWidth
            XPicture = ACos(XTube / (TubeWidth / 2)) / RDS
            If Offset + XPicture < Source.ScaleWidth Then
                Target.PaintPicture Source.Picture, Target.ScaleWidth - Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XPicture, 0, 1, Source.ScaleHeight
            Else
                Target.PaintPicture Source.Picture, Target.ScaleWidth - Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight, Offset + XTube - TubeWidth, 0, 1, Source.ScaleHeight
            End If
        Next XTube
    Next Offset

End Sub
Private Function ACos(x As Double)
    x = x - 1
    If x < 1 And x > -1 Then
        ACos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    Else
        ACos = 0
    End If
End Function
