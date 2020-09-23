VERSION 5.00
Begin VB.Form frmMouseOver 
   Caption         =   "Microsoft Mouse Buttons"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   3465
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMouse 
      Height          =   3000
      Left            =   240
      ScaleHeight     =   2940
      ScaleWidth      =   2940
      TabIndex        =   1
      Tag             =   "Skinned"
      Top             =   210
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Tag             =   "Skinned"
      Top             =   3375
      Width           =   3000
   End
   Begin VB.Image imgWheelDown 
      Height          =   3000
      Left            =   5655
      Picture         =   "frmMicrosoftMice.frx":0000
      Top             =   4575
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgWheelUp 
      Height          =   3000
      Left            =   2280
      Picture         =   "frmMicrosoftMice.frx":1D502
      Top             =   4545
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgX2 
      Height          =   3000
      Left            =   630
      Picture         =   "frmMicrosoftMice.frx":3AA04
      Top             =   4755
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgLeft 
      Height          =   3000
      Left            =   4815
      Picture         =   "frmMicrosoftMice.frx":57F06
      Top             =   6480
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgMiddle 
      Height          =   3000
      Left            =   5175
      Picture         =   "frmMicrosoftMice.frx":75408
      Top             =   5700
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgRight 
      Height          =   3000
      Left            =   6630
      Picture         =   "frmMicrosoftMice.frx":9290A
      Top             =   5220
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgX1 
      Height          =   3000
      Left            =   1260
      Picture         =   "frmMicrosoftMice.frx":AFE0C
      Top             =   5850
      Visible         =   0   'False
      Width           =   3000
   End
   Begin VB.Image imgNone 
      Height          =   3000
      Left            =   2985
      Picture         =   "frmMicrosoftMice.frx":CD30E
      Top             =   5595
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmMouseOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Test program to demonstrate trapping of MouseEnter, MouseExit and MouseWheel events.
'
' This method uses subclassing to detect the events, and will only report MouseEnter
' and MouseExit events for controls with the text "Skinnned" in the tag field.
'
' This program is a progression on the MouseEnter/MouseExit work submitted by Evan Toder.
'
' This new method allows you to caputer Events for controls that do not have a hwnd
' property (such as images and labels).

Private WithEvents cMouseEvents As clsMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

Private Sub Form_Load()
    ' Start listening for MouseEnter and MouseExit events.
    Set cMouseEvents = New clsMouseEvents
    Call cMouseEvents.HookMouseEvents(Me)
    
    picMouse = imgNone.Picture

End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' The form is being unloaded so we clean up.
    Call cMouseEvents.UnhookMouseEvents
    Set cMouseEvents = Nothing
    'Set cHook = Nothing
End Sub

Private Sub cMouseEvents_MouseDown(Button As Integer, x As Single, y As Single)
    Text2.Text = "MouseDown, Button = " & Button & ""
    Select Case Button
        Case Is = 0:    picMouse = imgNone.Picture
        Case Is = 1:    picMouse = imgLeft.Picture
        Case Is = 2:    picMouse = imgRight.Picture
        Case Is = 4:    picMouse = imgMiddle.Picture
        Case Is = 8:    picMouse = imgX1.Picture
        Case Is = 16:   picMouse = imgX2.Picture
    End Select
End Sub

Private Sub cMouseEvents_MouseUp(Button As Integer, x As Single, y As Single)
    Text2.Text = "MouseUp, Button = " & Button & ""
    picMouse = imgNone.Picture
End Sub

Private Sub cMouseEvents_MouseWheelDown(x As Single, y As Single)
    Text2.Text = "MouseWheel Down"
    ' MouseWheelDown Event detected.
    picMouse = imgWheelDown.Picture
End Sub

Private Sub cMouseEvents_MouseWheelUp(x As Single, y As Single)
    Text2.Text = "MouseWheel Up"
    ' MouseWheelUp Event detected.
    picMouse = imgWheelUp.Picture
End Sub
