Attribute VB_Name = "HoverCarsMod"
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Any) As Long

Public Type tCoOrd
   X As Byte
   Y As Byte
End Type

Public Type tCar
   X As Single
   Y As Single 'positon
   xm As Single
   ym As Single 'velocity
   Speed As Single 'speed
   Check As Byte 'checkpoints passed
   Angle As Byte 'direction the car is facing
   Throttle As Byte 'how the car's speed is changing
   Acceleration As Byte
   MaxSpeed As Byte
   Handling As Byte 'car attributes
   Steer As Byte 'state of the steering wheel
End Type
'steer constants
Public Const dSTRAIGHT = 0
Public Const dLEFT = 1
Public Const dRIGHT = 2

Public Type tTile
   Theme As Byte 'what folder is the pic in?
   ID As Byte 'what file is it in that folder?
   Target As Byte 'the target is what the computer cars aim for
End Type
'theme type constants
Public Const URBAN = 0
Public Const SEA = 1
Public Const MUDDY = 2
Public Const TEST = 3
Public Const BEACH = 4
'target constants
Public Const N = 1
Public Const E = 4
Public Const S = 2
Public Const W = 3
'tile ID's
Public Const NE = 5
Public Const NW = 6
Public Const SE = 7
Public Const SW = 8
Public Const NS = 9
Public Const EW = 10
Public Const BLANK = 0
Public Const F1 = 11
Public Const F2 = 12
Public Const F3 = 13
Public Const F4 = 14
Public Const F5 = 15
Public Const F6 = 16
Public Const STARTGRID = 17

Public Type tCourse
   Tile(1 To 10, 1 To 10) As tTile
End Type

Public Car() As tCar
Public Course As tCourse

Public Sine(0 To 35) As Single 'my angle system has only
Public Cosine(0 To 35) As Single '36 points in it!
Public Const PI = 3.14159265358979 'obvious
Public Const PIdiv18 = PI / 18 'used to convert 10degrees to radians

Public i As Integer 'used for loops
Public i2 As Integer
Public i3 As Integer
Public X As Integer 'used for loops
Public Y As Integer 'used for loops

Public Disp_Width As Integer 'size of drawing area
Public Disp_Height As Integer

Public CarPic(0 To 35) As IPictureDisp
Public MaskCarPic(0 To 35) As IPictureDisp

Public Opponents As Byte
Public Difficulty As Byte
Public UserCar As Byte
'difficulty levels
Public Const EASY = 0
Public Const MEDIUM = 1
Public Const HARD = 2
'cars to choose
Public Const GRIPPY = 0
Public Const SMOOTHY = 1
Public Const SPEEDER = 2
'collison detection can be turned on or off
Public Collisions As Boolean
'what type of painting of the course it is
Public Const OUTLINE = False
Public Const NORMAL = True

Public Sub BuildTrigTable()
'remembers all the sin and cos values needed
'(my system has 36 points to a circle, not 360)

For i = 0 To 35
  Sine(i) = Sin(i * PIdiv18)
  Cosine(i) = Cos(i * PIdiv18)
Next
End Sub

Public Sub LoadCarPics(Car_ID As Byte)
On Error Resume Next
'loads all the hover cars and masks needed

For i = 0 To 35
  Set CarPic(i) = LoadPicture(App.Path & "\Resources\Pictures\HoverCars\HoverCar" & Car_ID & "\HoverCar" & i * 10 & ".bmp")
  Set MaskCarPic(i) = LoadPicture(App.Path & "\Resources\Pictures\HoverCars\Masks\mHoverCar" & i * 10 & ".bmp")
Next
End Sub


Public Sub CalcDispSize()
'calculates the size of the drawing area

Disp_Width = Screen.Height * 1.3 / Screen.TwipsPerPixelX
Disp_Height = Screen.Height * 0.975 / Screen.TwipsPerPixelY
End Sub

Public Sub CreateDefaultCourse()
'creates the defualt course

'first blank out all tiles + give default directions
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = URBAN
   Course.Tile(X, Y).ID = BLANK
   Course.Tile(X, Y).Target = N
Next
Next

'and make some perimeter walls
For i = 2 To 9
   Course.Tile(i, 1).ID = N
   Course.Tile(i, 10).ID = S
   Course.Tile(1, i).ID = W
   Course.Tile(10, i).ID = E
Next
Course.Tile(1, 1).ID = NW
Course.Tile(1, 10).ID = SW
Course.Tile(10, 1).ID = NE
Course.Tile(10, 10).ID = SE

End Sub

Public Sub PaintCourse(TempPB As PictureBox, PB As PictureBox, Mode As Boolean)
'On Error Resume Next
Select Case Mode
   Case OUTLINE
        For X = 1 To 10
        For Y = 1 To 10
          TempPB.Picture = LoadPicture(App.Path & "\Resources\Pictures\Courses\" & Course.Tile(X, Y).Theme & "\Masks\" & Course.Tile(X, Y).ID & ".bmp")
          TempPB.ForeColor = vbBlue
          Select Case Course.Tile(X, Y).Target
             Case N
                TempPB.Line (0.5, 0.2)-(0.5, 0.8)
                TempPB.Line (0.5, 0.2)-(0.8, 0.5)
                TempPB.Line (0.5, 0.2)-(0.2, 0.5)
             Case S
                TempPB.Line (0.5, 0.8)-(0.5, 0.2)
                TempPB.Line (0.5, 0.8)-(0.8, 0.5)
                TempPB.Line (0.5, 0.8)-(0.2, 0.5)
             Case W
                TempPB.Line (0.2, 0.5)-(0.8, 0.5)
                TempPB.Line (0.2, 0.5)-(0.5, 0.8)
                TempPB.Line (0.2, 0.5)-(0.5, 0.2)
             Case E
                TempPB.Line (0.8, 0.5)-(0.2, 0.5)
                TempPB.Line (0.8, 0.5)-(0.5, 0.8)
                TempPB.Line (0.8, 0.5)-(0.5, 0.2)
        End Select
        TempPB.Picture = TempPB.Image
        PB.PaintPicture TempPB.Picture, X - 1, Y - 1
        Next
        Next
   Case NORMAL
        For X = 1 To 10
        For Y = 1 To 10
          TempPB.Picture = LoadPicture(App.Path & "\Resources\Pictures\Courses\" & Course.Tile(X, Y).Theme & "\" & Course.Tile(X, Y).ID & ".bmp")
          PB.PaintPicture TempPB.Picture, X - 1, Y - 1
        Next
        Next
End Select

End Sub

Public Sub PaintTile(TempPB As PictureBox, PB As PictureBox, X As Byte, Y As Byte, Mode As Boolean)
'On Error Resume Next
Select Case Mode
   Case OUTLINE
          TempPB.Picture = LoadPicture(App.Path & "\Resources\Pictures\Courses\" & Course.Tile(X, Y).Theme & "\Masks\" & Course.Tile(X, Y).ID & ".bmp")
          TempPB.ForeColor = vbBlue
          Select Case Course.Tile(X, Y).Target
             Case N
                TempPB.Line (0.5, 0.2)-(0.5, 0.8)
                TempPB.Line (0.5, 0.2)-(0.8, 0.5)
                TempPB.Line (0.5, 0.2)-(0.2, 0.5)
             Case S
                TempPB.Line (0.5, 0.8)-(0.5, 0.2)
                TempPB.Line (0.5, 0.8)-(0.8, 0.5)
                TempPB.Line (0.5, 0.8)-(0.2, 0.5)
             Case W
                TempPB.Line (0.2, 0.5)-(0.8, 0.5)
                TempPB.Line (0.2, 0.5)-(0.5, 0.8)
                TempPB.Line (0.2, 0.5)-(0.5, 0.2)
             Case E
                TempPB.Line (0.8, 0.5)-(0.2, 0.5)
                TempPB.Line (0.8, 0.5)-(0.5, 0.8)
                TempPB.Line (0.8, 0.5)-(0.5, 0.2)
        End Select
        TempPB.Picture = TempPB.Image
        PB.PaintPicture TempPB.Picture, X - 1, Y - 1
   Case NORMAL
          TempPB.Picture = LoadPicture(App.Path & "\Resources\Pictures\Courses\" & Course.Tile(X, Y).Theme & "\" & Course.Tile(X, Y).ID & ".bmp")
          PB.PaintPicture TempPB.Picture, X - 1, Y - 1
End Select

End Sub

Public Function SaveCourse(FileName As String) As Boolean
On Error GoTo muffup
Open FileName For Random As #1 Len = 1
For X = 1 To 10
For Y = 1 To 10
   i = ((Y * 10) + X) * 3
   Put #1, i, Course.Tile(X, Y).Theme
   Put #1, i + 1, Course.Tile(X, Y).ID
   Put #1, i + 2, Course.Tile(X, Y).Target
Next
Next
Close #1
SaveCourse = True
Exit Function
muffup:
SaveCourse = False
Close #1
End Function

Public Function LoadCourse(FileName As String) As Boolean
On Error GoTo muffup
Open FileName For Random As #1 Len = 1
For X = 1 To 10
For Y = 1 To 10
   i = ((Y * 10) + X) * 3
   Get #1, i, Course.Tile(X, Y).Theme
   Get #1, i + 1, Course.Tile(X, Y).ID
   Get #1, i + 2, Course.Tile(X, Y).Target
Next
Next
LoadCourse = True
Exit Function
muffup:
LoadCourse = False
Close #1
End Function
