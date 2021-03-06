Attribute VB_Name = "modMain"
Option Explicit
' Source originale : http://www.vbfrance.com/code.aspx?ID=27944
' Auteur : Jack

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private MesFormes() As New Forme
Private Degr�() As Byte, Sens() As Boolean, Incr�ment() As Byte
Private OnOff As Boolean
'

Sub Main()

    Dim r As Integer, t As Integer
    Dim Largeur As Long, Longueur As Long, Compteur As Long, M�moHwnd As Long
    
    Dim Interval As Long, NbFormes As Integer
    
    MsgBox "Il suffit de taper sur la touche 'Echap' (doucement) " & vbCrLf & _
           "pour stopper l'application" & vbCrLf & vbCrLf & _
           "Press Esc key to end the program"
    
    ' Param�trage :
    NbFormes = 10           '/ Number of objects
    Interval = 75 ' msec    '/ Refresh period
    
    ' Pr�pare les instances
    ReDim MesFormes(NbFormes)
    ReDim Degr�(NbFormes)
    ReDim Sens(NbFormes)
    ReDim Incr�ment(NbFormes)
    
    ' M�mo taille �cran - taille d'une image pour ne pas sortir de l'�cran
    Largeur = Screen.Width - Forme.Picture1.Width
    Longueur = Screen.Height - Forme.Picture1.Height
    
    ' Charge et place les formes al�atoirement
    Randomize
    For r = 1 To NbFormes
        ' Charge la forme
        Load MesFormes(r)
        ' On la positionne au hasard
        MesFormes(r).Left = CLng(Rnd() * Largeur)
        MesFormes(r).Top = CLng(Rnd() * Longueur)
        ' On l'affiche
        MesFormes(r).Show
        ' On choisi un degr� de transparence au hasard
        Degr�(r) = CByte(Rnd() * 255)
        Call Transparence("On", MesFormes(r), Degr�(r))
        ' et un sens de variation
        Sens(r) = True
        ' et de l'incr�ment
        Incr�ment(r) = CByte(Rnd() * 20)
    Next r
    
    ' D�marre le timer de gestion de transparence
    OnOff = True
    Do While OnOff
        If GetTickCount > (Compteur + Interval) Then
            ' Ceci s'ex�cute toutes les 'Interval' millisecondes
            For r = 1 To NbFormes
                ' Ajoute ou retranche l'incr�ment selon le sens
                t = Degr�(r) + IIf(Sens(r) = False, Incr�ment(r), Incr�ment(r) * (-1))
                ' V�rifie les d�passements
                If t < 0 Then
                    t = 0
                    ' On change de sens (apparition)
                    Sens(r) = Not Sens(r)
                    ' On change la position
                    MesFormes(r).Left = CLng(Rnd() * Largeur)
                    MesFormes(r).Top = CLng(Rnd() * Longueur)
                    ' R�cup�re le handle de l'appli qui a le focus
                    M�moHwnd = GetForegroundWindow
                    If M�moHwnd <> MesFormes(r).hwnd Then
                        ' Place notre image devant
                        Call SetTop(MesFormes(r), True)
                        ' Redonne le focus � l'appli de travail
                        Call SetActiveWindow(M�moHwnd)
                    End If
                ElseIf t > 255 Then
                    t = 255
                    ' On change de sens (disparition)
                    Sens(r) = Not Sens(r)
                End If
                ' Applique la nouvelle transparence
                Degr�(r) = CByte(t)
                Call Transparence("On", MesFormes(r), Degr�(r))
            Next r
            ' M�morise "l'heure"
            Compteur = GetTickCount
        End If
        DoEvents    ' Repasse la main au syst�me
    Loop
    
End Sub

Public Sub D�chargeTout()
    
    ' Demande au Timer de s'arr�ter
    OnOff = False
    DoEvents
    ' D�charge toutes les formes
    Dim xx As Form
    For Each xx In Forms
        Call Transparence("Off", xx)
        Unload xx
    Next
    ' et bye bye
    End
    
End Sub
