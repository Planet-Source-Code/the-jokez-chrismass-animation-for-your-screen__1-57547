VERSION 5.00
Begin VB.Form Forme 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   360
   ClientTop       =   4080
   ClientWidth     =   8160
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Forme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   120
      Picture         =   "Forme.frx":000C
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "Forme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
'

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' On arrive ici grace au KeyPreview de la forme
    If KeyCode = KeyCodeConstants.vbKeyEscape Then Call DéchargeTout
End Sub

Private Sub Form_Load()

    Dim WindowRegion As Long
    
    'Propriétés de la picture box
    Picture1.AutoRedraw = True
    Picture1.AutoSize = True
    Picture1.BorderStyle = 0
    Picture1.ScaleMode = 3
    
    'Position de la picture box
    Picture1.Top = 0: Picture1.Left = 0
    
    ' "Découpe" la form suivant Picture1
    WindowRegion = MakeRegion(Picture1)
    SetWindowRgn Me.hwnd, WindowRegion, True

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


