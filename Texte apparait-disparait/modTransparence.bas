Attribute VB_Name = "modTransparence"
Option Explicit
' Source originale : http://www.vbfrance.com/code.aspx?ID=24602

'''''Déclaration des constantes
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = (-20)

'''''Apis nécessaires pour la transparence
Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                                    ByVal hwnd As Long, _
                                    ByVal crKey As Long, _
                                    ByVal bAlpha As Byte, _
                                    ByVal dwFlags As Long) As Boolean
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                                    ByVal hwnd As Long, _
                                    ByVal nIndex As Long, _
                                    ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                                    ByVal hwnd As Long, _
                                    ByVal nIndex As Long) As Long
'

Public Sub Transparence(State As String, _
                        Fenêtre As Form, _
                        Optional ByVal Alpha As Byte = 255)
    
    Select Case UCase(State)
        Case "ON"
                SetWindowLong Fenêtre.hwnd, GWL_EXSTYLE, GetWindowLong(Fenêtre.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED
                SetLayeredWindowAttributes Fenêtre.hwnd, 0, Alpha, LWA_ALPHA
        Case "OFF"
                SetWindowLong Fenêtre.hwnd, GWL_EXSTYLE, GetWindowLong(Fenêtre.hwnd, GWL_EXSTYLE) - WS_EX_LAYERED
    End Select
End Sub
