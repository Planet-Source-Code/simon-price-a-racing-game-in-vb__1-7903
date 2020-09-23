VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Cars"
   ClientHeight    =   4920
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5604
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox LevelPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   13860
      Left            =   2520
      ScaleHeight     =   11.55
      ScaleMode       =   0  'User
      ScaleWidth      =   11.25
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   13500
   End
   Begin VB.PictureBox Map 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   12000
      Left            =   1200
      ScaleHeight     =   10
      ScaleMode       =   0  'User
      ScaleWidth      =   10
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   12000
   End
   Begin VB.PictureBox TempPB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   3720
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Mask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   1920
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicCar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1200
      Left            =   120
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   240
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   9600
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileName As String
Public Paused As Boolean

'throttle constants : the cars can hover, accelerate, or reverse(brake)
Const HOVER = 0
Const ACC = 1
Const REVERSE = 2
'these contsants affect the physics of the game
Const FRICTION = 1.05
Const WALLHIT = 2
'backbuffer size constants
Const PB_WIDTH = 1000
Const PB_HEIGHT = 1000
Const PB_WIDTHdiv2 = PB_WIDTH \ 2
Const PB_HEIGHTdiv2 = PB_HEIGHT \ 2
Const PB_WIDTHdiv4 = PB_WIDTH \ 4
Const PB_HEIGHTdiv4 = PB_HEIGHT \ 4
Const PB_WIDTHdiv8 = PB_WIDTH \ 8
Const PB_HEIGHTdiv8 = PB_HEIGHT \ 8
'level size constants
Const LEVEL_WIDTH = 1200
Const LEVEL_HEIGHT = 800

Private Sub LayoutForm()
'sets up the form's layout and works out other sizes needed

Move 0, 0, Screen.Width, Screen.Height
PB.Move 0, 100, PB_WIDTH, PB_HEIGHT
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyLeft 'turn left
     Car(1).Steer = dLEFT
  Case vbKeyRight 'turn right
     Car(1).Steer = dRIGHT
  Case vbKeyP
     Paused = Not Paused
     If Paused = False Then MainLoop
  Case vbKeyEscape
     Paused = True
     SForm.Visible = True
     Unload Me
End Select
End Sub

Public Sub GoAhead()
On Error Resume Next

LayoutForm
Show 'sort out the form's layout before showing it
MainLoop 'go to the main game loop
End Sub

Public Sub LoadCourse()
PaintCourse TempPB, LevelPic, True
End Sub

Public Sub LoadCourseOutline()
PaintCourse TempPB, Map, False
End Sub

Public Sub SetUpLevel()
'put in default pics
PicCar = CarPic(0)
Mask = MaskCarPic(0)
End Sub

Public Sub MainLoop()
'this is the main game loop

On Error Resume Next

Do

DoEvents

'********************************************
'this is the AI bit, it decides what each car will do
For i = 2 To UBound(Car)
  X = Car(i).X \ 100 + 1
  Y = (Car(i).Y + 20) \ 100 + 1
Select Case Course.Tile(X, Y).Target
   Case N
     Select Case Car(i).Angle
        Case 18 To 35: Car(i).Steer = dRIGHT: Car(i).Throttle = HOVER
        Case 1 To 17: Car(i).Steer = dLEFT: Car(i).Throttle = HOVER
        Case 0: Car(i).Steer = dSTRAIGHT: Car(i).Throttle = ACC
     End Select
   Case E
     Select Case Car(i).Angle
        Case 27 To 35: Car(i).Steer = dRIGHT: Car(i).Throttle = HOVER
        Case 0 To 8: Car(i).Steer = dRIGHT: Car(i).Throttle = HOVER
        Case 10 To 26: Car(i).Steer = dLEFT: Car(i).Throttle = HOVER
        Case 9: Car(i).Steer = dSTRAIGHT: Car(i).Throttle = ACC
     End Select
   Case S
     Select Case Car(i).Angle
        Case 0 To 17: Car(i).Steer = dRIGHT: Car(i).Throttle = HOVER
        Case 19 To 35: Car(i).Steer = dLEFT: Car(i).Throttle = HOVER
        Case 18: Car(i).Steer = dSTRAIGHT: Car(i).Throttle = ACC
     End Select
   Case W
     Select Case Car(i).Angle
        Case 28 To 35: Car(i).Steer = dLEFT: Car(i).Throttle = HOVER
        Case 0 To 9: Car(i).Steer = dLEFT: Car(i).Throttle = HOVER
        Case 10 To 26: Car(i).Steer = dRIGHT: Car(i).Throttle = HOVER
        Case 27: Car(i).Steer = dSTRAIGHT: Car(i).Throttle = ACC
     End Select
End Select
Next

'********************************************
'this next bit moves everything

For i = 1 To UBound(Car)
'see what the car is doing
Select Case Car(i).Throttle
  'accelerating
  Case ACC
  Car(i).Speed = Car(i).Speed + Car(i).Acceleration
  'reversing or braking
  Case REVERSE
  Car(i).Speed = Car(i).Speed - Car(i).Acceleration
End Select

Select Case Car(i).Steer
  Case dLEFT 'turn left
  If Car(i).Angle = 0 Then
    Car(i).Angle = 35
  Else
   Car(i).Angle = Car(i).Angle - 1
  End If
  
  Case dRIGHT 'turn right
  If Car(i).Angle = 35 Then
    Car(i).Angle = 0
  Else
    Car(i).Angle = Car(i).Angle + 1
  End If
End Select

'work out the direction of travel
Car(i).xm = ((Car(i).xm * Car(i).Handling) + (Car(i).Speed * Sine(Car(i).Angle))) / (Car(i).Handling + 1)
Car(i).ym = ((Car(i).ym * Car(i).Handling) + (-Car(i).Speed * Cosine(Car(i).Angle))) / (Car(i).Handling + 1)

'check for collisions with scenery
If GetPixel(Map.hdc, Car(i).X + Car(i).xm, Car(i).Y + Car(i).ym + 35) = vbRed Then
  Car(i).ym = -Car(i).ym / WALLHIT
End If

If GetPixel(Map.hdc, Car(i).X + Car(i).xm, Car(i).Y + Car(i).ym + 15) = vbRed Then
  Car(i).ym = -Car(i).ym / WALLHIT
End If

If GetPixel(Map.hdc, Car(i).X + Car(i).xm + 10, Car(i).Y + Car(i).ym + 25) = vbRed Then
  Car(i).xm = -Car(i).xm / WALLHIT
End If

If GetPixel(Map.hdc, Car(i).X + Car(i).xm - 10, Car(i).Y + Car(i).ym + 25) = vbRed Then
  Car(i).xm = -Car(i).xm / WALLHIT
End If

On Error Resume Next
Dim HitRatio As Single
If Collisions Then

'check for collisions with other cars
For i2 = i + 1 To UBound(Car)
  Select Case Car(i).X
     Case Car(i2).X - 15 To Car(i2).X + 15
        Select Case Car(i).Y
           Case Car(i2).Y - 15 To Car(i2).Y + 15
              Car(i).xm = -Car(i).xm
              Car(i).ym = -Car(i).ym
              Car(i2).xm = -Car(i2).xm
              Car(i2).ym = -Car(i2).ym
'               HitRatio = Abs((Car(i).xm) + Abs(Car(i).ym)) / (Abs(Car(i2).xm) + Abs(Car(i2).ym))
'               Car(i2).xm = Car(i2).xm * -HitRatio
'               Car(i).xm = Car(i).xm * (-1 / HitRatio)
'               Car(i2).ym = Car(i2).ym * -HitRatio
'               Car(i).ym = Car(i).ym * (-1 / HitRatio)
        End Select
  End Select
Next

End If

'now move in that direction
Car(i).X = Car(i).X + Car(i).xm
Car(i).Y = Car(i).Y + Car(i).ym

'now simulate friction by slowing the hovercar a little
Car(i).Speed = Car(i).Speed / FRICTION

Next

'********************************************
'this next bit draws everything

'first clear the backbuffer
PB.Cls
'paint the level onto the backbuffer in the current camera position
StretchBlt PB.hdc, 0, 0, PB_WIDTH, PB_HEIGHT, LevelPic.hdc, Car(1).X - 125, Car(1).Y - 94, PB_WIDTHdiv4, PB_HEIGHTdiv4, vbSrcCopy

For i = 1 To UBound(Car)
'swap the pics over to the correct angle
PicCar = CarPic(Car(i).Angle)
Mask = MaskCarPic(Car(i).Angle)
'create a space on the backbuffer for the car to go in
BitBlt PB.hdc, 450 - (Car(1).X - Car(i).X) * 4, 450 - (Car(1).Y - Car(i).Y) * 4, 100, 100, Mask.hdc, 0, 0, vbMergePaint
'now drop the hovercar in the white space
BitBlt PB.hdc, 450 - (Car(1).X - Car(i).X) * 4, 450 - (Car(1).Y - Car(i).Y) * 4, 100, 100, PicCar.hdc, 0, 0, vbSrcAnd
Next

'now copy the backbuffer pic into sight on the form
StretchBlt hdc, 0, 0, Disp_Width, Disp_Height, PB.hdc, 0, 0, PB_WIDTH, PB_HEIGHT, vbSrcCopy

Caption = Int(Car(1).X) & "," & Int(Car(1).Y)

Loop Until Paused

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'steer straight
Car(1).Steer = dSTRAIGHT
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button

Case 1 'accelerate
Car(1).Throttle = ACC

Case 2 'brake(or reverse)
Car(1).Throttle = REVERSE

End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'just hover along
Car(1).Throttle = HOVER
End Sub
