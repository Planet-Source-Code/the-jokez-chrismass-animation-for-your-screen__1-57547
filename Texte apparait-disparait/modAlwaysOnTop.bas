Attribute VB_Name = "modAlwaysOnTop"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" ( _
                                ByVal hwnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal x As Long, _
                                ByVal y As Long, _
                                ByVal cx As Long, _
                                ByVal cy As Long, _
                                ByVal uFlags As Long) As Long

' Constantes de SetWindowPos :
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
'

Public Sub SetTop(Form As Form, _
                  ByVal Topmost As Boolean)
    
    Dim hWndInsertAfter As Long
    
    If Topmost Then
        hWndInsertAfter = HWND_TOPMOST
    Else
        hWndInsertAfter = HWND_NOTOPMOST
    End If
    
    SetWindowPos Form.hwnd, hWndInsertAfter, 0, 0, 0, 0, _
        SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
        
End Sub

