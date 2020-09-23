Attribute VB_Name = "modForme"
Option Explicit
' Source originale : Inconnu

Public Declare Function SetWindowRgn Lib "user32" ( _
                                ByVal hwnd As Long, _
                                ByVal hRgn As Long, _
                                ByVal bRedraw As Boolean) As Long
Private Declare Function GetPixel Lib "gdi32" ( _
                                ByVal hDC As Long, _
                                ByVal x As Long, _
                                ByVal y As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" ( _
                                ByVal X1 As Long, _
                                ByVal Y1 As Long, _
                                ByVal X2 As Long, _
                                ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" ( _
                                ByVal hDestRgn As Long, _
                                ByVal hSrcRgn1 As Long, _
                                ByVal hSrcRgn2 As Long, _
                                ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" ( _
                                ByVal hObject As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                                ByVal hwnd As Long, _
                                ByVal wMsg As Long, _
                                ByVal wParam As Long, _
                                lParam As Any) As Long
Private Const RGN_OR = 2
'

Public Function MakeRegion(picSkin As PictureBox) As Long
    
    ' faites une fenêtre "région" basée sur une picture de picture box
    ' Ceci ce fait en passant l'image pixel par pixel et en créant une
    ' région pour chaque pixel non transparent
    ' Le code est optimisé, il est donc assez rapide
    
    Dim x As Long, y As Long, StartLineX As Long
    Dim FullRegion As Long, LineRegion As Long
    Dim TransparentColor As Long
    Dim InFirstRegion As Boolean
    Dim InLine As Boolean
    Dim hDC As Long
    Dim PicWidth As Long
    Dim PicHeight As Long
    
    hDC = picSkin.hDC
    PicWidth = picSkin.ScaleWidth
    PicHeight = picSkin.ScaleHeight
    
    InFirstRegion = True: InLine = False
    x = y = StartLineX = 0
    
    ' Ici, la couleur de transparence est basé sur le pixel en haut a gauche
    ' Mais vous pouvez mettre la couleur ke vous voulez
    TransparentColor = GetPixel(hDC, 0, 0)
    
    For y = 0 To PicHeight - 1
        For x = 0 To PicWidth - 1
            
            If GetPixel(hDC, x, y) = TransparentColor Or x = PicWidth Then

                If InLine Then
                    InLine = False
                    LineRegion = CreateRectRgn(StartLineX, y, x, y + 1)
                    
                    If InFirstRegion Then
                        FullRegion = LineRegion
                        InFirstRegion = False
                    Else
                        CombineRgn FullRegion, FullRegion, LineRegion, RGN_OR

                        DeleteObject LineRegion
                    End If
                End If
            Else
                If Not InLine Then
                    InLine = True
                    StartLineX = x
                End If
            End If
        Next
    Next
    
    MakeRegion = FullRegion
End Function


