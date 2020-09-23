VERSION 5.00
Begin VB.Form SForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome To Hover Cars"
   ClientHeight    =   4176
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5640
   Icon            =   "SForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cmdLevelDesign 
      BorderStyle     =   1  'Fixed Single
      Height          =   1092
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   5172
   End
   Begin VB.Image cmdPlay 
      BorderStyle     =   1  'Fixed Single
      Height          =   1092
      Left            =   240
      Stretch         =   -1  'True
      Top             =   720
      Width           =   5172
   End
End
Attribute VB_Name = "SForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLevelDesign_Click()
DForm.Visible = True
Unload Me
End Sub

Private Sub cmdPlay_Click()
CarForm.Visible = True
Unload Me
End Sub

Private Sub Form_Load()
Randomize Timer
cmdPlay = LoadPicture(App.Path & "\Resources\Pictures\Misc\PlayTheGame.bmp")
cmdLevelDesign = LoadPicture(App.Path & "\Resources\Pictures\Misc\DesignACourse.bmp")
End Sub
