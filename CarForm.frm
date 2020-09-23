VERSION 5.00
Begin VB.Form CarForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Car Selection"
   ClientHeight    =   3996
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5172
   Icon            =   "CarForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   333
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   431
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame OppF 
      Caption         =   "No. of Opponents"
      Height          =   2172
      Left            =   3600
      TabIndex        =   16
      Top             =   1200
      Width           =   1452
      Begin VB.CheckBox ColC 
         Caption         =   "Collision Detection"
         Height          =   372
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "Twenty !"
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "Ten"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "Five"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "Three"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "Just One"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1212
      End
      Begin VB.OptionButton OppO 
         Caption         =   "None"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.Frame DiffF 
      Caption         =   "Difficulty :"
      Height          =   1092
      Left            =   3600
      TabIndex        =   12
      Top             =   120
      Width           =   1452
      Begin VB.OptionButton DiffO 
         Caption         =   "Pro"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1212
      End
      Begin VB.OptionButton DiffO 
         Caption         =   "Intermediate"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   1212
      End
      Begin VB.OptionButton DiffO 
         Caption         =   "Novice"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1212
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK !"
      Height          =   492
      Left            =   3720
      TabIndex        =   11
      Top             =   3480
      Width           =   1212
   End
   Begin VB.Frame CarF 
      Caption         =   "Choose your hover-car :"
      Height          =   3732
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3372
      Begin VB.PictureBox PicCar 
         Height          =   1692
         Left            =   840
         ScaleHeight     =   137
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   6
         Top             =   600
         Width           =   1692
         Begin VB.PictureBox CarPB 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1212
            Left            =   120
            ScaleHeight     =   1212
            ScaleWidth      =   1212
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Timer SpinT 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   1320
            Top             =   0
         End
      End
      Begin VB.PictureBox AccBar 
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   1080
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   2
         TabIndex        =   5
         Top             =   2520
         Width           =   2172
      End
      Begin VB.PictureBox SpeedBar 
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   1080
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   20
         TabIndex        =   4
         Top             =   2880
         Width           =   2172
      End
      Begin VB.PictureBox HandlingBar 
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   1080
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   20
         TabIndex        =   3
         Top             =   3240
         Width           =   2172
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next>"
         Height          =   732
         Left            =   2640
         TabIndex        =   2
         Top             =   1080
         Width           =   612
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "<Back"
         Height          =   732
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   612
      End
      Begin VB.Label CarL 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SMOOTHY"
         Height          =   252
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "Acceleration"
         Height          =   252
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Speed"
         Height          =   252
         Left            =   480
         TabIndex        =   8
         Top             =   2880
         Width           =   612
      End
      Begin VB.Label Label5 
         Caption         =   "Handling"
         Height          =   252
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   732
      End
   End
End
Attribute VB_Name = "CarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AccBar_Paint()
AccBar.Line (0, 0)-(Car(1).Acceleration, 1), vbBlue, BF
End Sub

Private Sub cmdBack_Click()
If CarL = "SMOOTHY" Then
  CarL = "GRIPPER"
  cmdBack.Enabled = False
  UserCar = GRIPPY
Else
  CarL = "SMOOTHY"
  cmdNext.Enabled = True
  UserCar = SMOOTHY
End If
ChangeCar
End Sub

Private Sub cmdNext_Click()
If CarL = "SMOOTHY" Then
  CarL = "SPEEDER"
  cmdNext.Enabled = False
  UserCar = SPEEDER
Else
  CarL = "SMOOTHY"
  cmdBack.Enabled = True
  UserCar = SMOOTHY
End If
ChangeCar
End Sub

Private Sub ChangeCar()
MousePointer = vbHourglass
Select Case UserCar
  Case GRIPPY
    Car(1).Acceleration = 1
    Car(1).MaxSpeed = 12
    Car(1).Handling = 10
  Case SMOOTHY
    Car(1).Acceleration = 1
    Car(1).MaxSpeed = 16
    Car(1).Handling = 15
  Case SPEEDER
    Car(1).Acceleration = 2
    Car(1).MaxSpeed = 20
    Car(1).Handling = 20
End Select
LoadCarPics (UserCar)
AccBar.Refresh
SpeedBar.Refresh
HandlingBar.Refresh
MousePointer = vbDefault
End Sub

Private Sub cmdOK_Click()
'set the collision detection
Collisions = ColC.Value
'set user hovercar
Select Case CarL
  Case "GRIPPER": UserCar = GRIPPY
  Case "SMOOTHY": UserCar = SMOOTHY
  Case "SPEEDER": UserCar = SPEEDER
End Select
'set how many hovercars to make
ReDim Car(1 To Opponents + 1)
    Car(1).Check = 0
    Car(1).Speed = 0
    Car(1).X = 50 + (Rnd * 50)
    Car(1).Y = 50 + (Rnd * 50)
    Car(1).Angle = Int(Rnd * 36)
    Car(1).MaxSpeed = 10
    Car(1).Handling = 15
    Car(1).Acceleration = 1
'change difficulty
Select Case Difficulty
Case EASY
    For i = 2 To UBound(Car)
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 8
    Car(i).Handling = 20
    Car(i).Acceleration = 1
    Next
Case MEDIUM
    For i = 2 To UBound(Car)
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 10
    Car(i).Handling = 15
    Car(i).Acceleration = 1
    Next
Case HARD
    For i = 2 To UBound(Car)
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 15
    Car(i).Handling = 10
    Car(i).Acceleration = 2
    Next
End Select

    i = 1
Select Case UserCar
Case GRIPPY
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 8
    Car(i).Handling = 20
    Car(i).Acceleration = 1
Case SMOOTHY
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 10
    Car(i).Handling = 15
    Car(i).Acceleration = 1
Case SPEEDER
    Car(i).Check = 0
    Car(i).Speed = 0
    Car(i).X = 50 + (Rnd * 50)
    Car(i).Y = 50 + (Rnd * 50)
    Car(i).Angle = Int(Rnd * 36)
    Car(i).MaxSpeed = 15
    Car(i).Handling = 10
    Car(i).Acceleration = 2
End Select
SelectForm.Visible = True
Unload Me
End Sub

Private Sub DiffO_Click(Index As Integer)
Difficulty = Index
End Sub

Private Sub Form_Load()
ReDim Car(1 To 1)
PicCar.BackColor = vbWhite
Opponents = 5
SpinT.Enabled = True
UserCar = SMOOTHY
ChangeCar
End Sub

Private Sub HandlingBar_Paint()
HandlingBar.Line (0, 0)-(25 - Car(1).Handling, 1), vbBlue, BF
End Sub

Private Sub OppO_Click(Index As Integer)
Opponents = Index
End Sub

Private Sub SpeedBar_Paint()
SpeedBar.Line (0, 0)-(Car(1).MaxSpeed, 1), vbBlue, BF
End Sub

Private Sub SpinT_Timer()
If Car(1).Angle = 35 Then
  Car(1).Angle = 0
Else
  Car(1).Angle = Car(1).Angle + 1
End If

CarPB = CarPic(Car(1).Angle)
BitBlt PicCar.hdc, 19, 19, 100, 100, CarPB.hdc, 0, 0, vbSrcCopy
End Sub
