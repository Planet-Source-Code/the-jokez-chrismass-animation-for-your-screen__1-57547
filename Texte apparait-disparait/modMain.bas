Attribute VB_Name = "modMain"
Option Explicit
' Source originale : http://www.vbfrance.com/code.aspx?ID=27944
' Auteur : Jack

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private MesFormes() As New Forme
Private Degré() As Byte, Sens() As Boolean, Incrément() As Byte
Private OnOff As Boolean
'

Sub Main()

    Dim r As Integer, t As Integer
    Dim Largeur As Long, Longueur As Long, Compteur As Long, MémoHwnd As Long
    
    Dim Interval As Long, NbFormes As Integer
    
    MsgBox "Il suffit de taper sur la touche 'Echap' (doucement) " & vbCrLf & _
           "pour stopper l'application" & vbCrLf & vbCrLf & _
           "Press Esc key to end the program"
    
    ' Paramétrage :
    NbFormes = 10           '/ Number of objects
    Interval = 75 ' msec    '/ Refresh period
    
    ' Prépare les instances
    ReDim MesFormes(NbFormes)
    ReDim Degré(NbFormes)
    ReDim Sens(NbFormes)
    ReDim Incrément(NbFormes)
    
    ' Mémo taille écran - taille d'une image pour ne pas sortir de l'écran
    Largeur = Screen.Width - Forme.Picture1.Width
    Longueur = Screen.Height - Forme.Picture1.Height
    
    ' Charge et place les formes aléatoirement
    Randomize
    For r = 1 To NbFormes
        ' Charge la forme
        Load MesFormes(r)
        ' On la positionne au hasard
        MesFormes(r).Left = CLng(Rnd() * Largeur)
        MesFormes(r).Top = CLng(Rnd() * Longueur)
        ' On l'affiche
        MesFormes(r).Show
        ' On choisi un degré de transparence au hasard
        Degré(r) = CByte(Rnd() * 255)
        Call Transparence("On", MesFormes(r), Degré(r))
        ' et un sens de variation
        Sens(r) = True
        ' et de l'incrément
        Incrément(r) = CByte(Rnd() * 20)
    Next r
    
    ' Démarre le timer de gestion de transparence
    OnOff = True
    Do While OnOff
        If GetTickCount > (Compteur + Interval) Then
            ' Ceci s'exécute toutes les 'Interval' millisecondes
            For r = 1 To NbFormes
                ' Ajoute ou retranche l'incrément selon le sens
                t = Degré(r) + IIf(Sens(r) = False, Incrément(r), Incrément(r) * (-1))
                ' Vérifie les dépassements
                If t < 0 Then
                    t = 0
                    ' On change de sens (apparition)
                    Sens(r) = Not Sens(r)
                    ' On change la position
                    MesFormes(r).Left = CLng(Rnd() * Largeur)
                    MesFormes(r).Top = CLng(Rnd() * Longueur)
                    ' Récupère le handle de l'appli qui a le focus
                    MémoHwnd = GetForegroundWindow
                    If MémoHwnd <> MesFormes(r).hwnd Then
                        ' Place notre image devant
                        Call SetTop(MesFormes(r), True)
                        ' Redonne le focus à l'appli de travail
                        Call SetActiveWindow(MémoHwnd)
                    End If
                ElseIf t > 255 Then
                    t = 255
                    ' On change de sens (disparition)
                    Sens(r) = Not Sens(r)
                End If
                ' Applique la nouvelle transparence
                Degré(r) = CByte(t)
                Call Transparence("On", MesFormes(r), Degré(r))
            Next r
            ' Mémorise "l'heure"
            Compteur = GetTickCount
        End If
        DoEvents    ' Repasse la main au système
    Loop
    
End Sub

Public Sub DéchargeTout()
    
    ' Demande au Timer de s'arrêter
    OnOff = False
    DoEvents
    ' Décharge toutes les formes
    Dim xx As Form
    For Each xx In Forms
        Call Transparence("Off", xx)
        Unload xx
    Next
    ' et bye bye
    End
    
End Sub
