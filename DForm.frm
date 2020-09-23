VERSION 5.00
Begin VB.Form DForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hover Cars Course Designer"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   7680
   Icon            =   "DForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame IDF 
      Caption         =   "Type"
      Height          =   2772
      Left            =   5040
      TabIndex        =   10
      Top             =   2520
      Width           =   2532
      Begin VB.OptionButton IDO 
         Caption         =   "Start &Grid"
         Height          =   252
         Index           =   17
         Left            =   1200
         TabIndex        =   29
         Top             =   2280
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &6"
         Height          =   252
         Index           =   16
         Left            =   1200
         TabIndex        =   27
         Top             =   2040
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &5"
         Height          =   252
         Index           =   15
         Left            =   1200
         TabIndex        =   26
         Top             =   1800
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &4"
         Height          =   252
         Index           =   14
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &3"
         Height          =   252
         Index           =   13
         Left            =   1200
         TabIndex        =   24
         Top             =   1320
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &2"
         Height          =   252
         Index           =   12
         Left            =   1200
         TabIndex        =   23
         Top             =   1080
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "Feature &1"
         Height          =   252
         Index           =   11
         Left            =   1200
         TabIndex        =   22
         Top             =   840
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "EW Wall"
         Height          =   252
         Index           =   10
         Left            =   1200
         TabIndex        =   21
         Top             =   600
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "NS Wall"
         Height          =   252
         Index           =   9
         Left            =   1200
         TabIndex        =   20
         Top             =   360
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "SW Wall"
         Height          =   252
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "SE Wall"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "NW Wall"
         Height          =   252
         Index           =   6
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "NE Wall"
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   1560
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "&East Wall"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   972
      End
      Begin VB.OptionButton IDO 
         Caption         =   "&West Wall"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1092
      End
      Begin VB.OptionButton IDO 
         Caption         =   "&South Wall"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1212
      End
      Begin VB.OptionButton IDO 
         Caption         =   "&North Wall"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1092
      End
      Begin VB.OptionButton IDO 
         Caption         =   "&Blank"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   972
      End
   End
   Begin VB.Frame ThemeF 
      Caption         =   "Theme"
      Height          =   1692
      Left            =   5040
      TabIndex        =   5
      Top             =   720
      Width           =   2532
      Begin VB.OptionButton Theme 
         Caption         =   "Beach"
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   28
         Top             =   1200
         Width           =   972
      End
      Begin VB.OptionButton Theme 
         Caption         =   "Countryside"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1212
      End
      Begin VB.OptionButton Theme 
         Caption         =   "Muddy"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   852
      End
      Begin VB.OptionButton Theme 
         Caption         =   "Sea"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   852
      End
      Begin VB.OptionButton Theme 
         Caption         =   "Urban"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
   End
   Begin VB.PictureBox TempPB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   5
      Height          =   852
      Left            =   0
      ScaleHeight     =   852
      ScaleWidth      =   972
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.PictureBox Preview 
      Height          =   4812
      Left            =   120
      ScaleHeight     =   4764
      ScaleWidth      =   4764
      TabIndex        =   0
      Top             =   120
      Width           =   4812
      Begin VB.PictureBox PB 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFF00&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         DrawStyle       =   1  'Dash
         DrawWidth       =   3
         Height          =   4812
         Left            =   240
         ScaleHeight     =   4812
         ScaleWidth      =   4812
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   4812
      End
   End
   Begin VB.Label TileL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tile (1,1) Properties"
      Height          =   492
      Left            =   5040
      TabIndex        =   4
      Top             =   120
      Width           =   2532
   End
   Begin VB.Label PreviewL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Preview (Normal Mode - double click preview to change mode)"
      Height          =   372
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   4812
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit to Main Menu"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuChangeTheme 
         Caption         =   "Change &Theme"
         Begin VB.Menu mnuUrban 
            Caption         =   "&Urban"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuSea 
            Caption         =   "&Sea"
            Shortcut        =   ^C
         End
         Begin VB.Menu mnuMuddy 
            Caption         =   "&Muddy"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuDrivingTest 
            Caption         =   "&Countryside"
         End
         Begin VB.Menu mnuBeach 
            Caption         =   "&Beach"
            Shortcut        =   ^B
         End
      End
   End
End
Attribute VB_Name = "DForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private PreviewMode As Boolean
Private Cur As tCoOrd
Private Mouse As tCoOrd

Private Sub Form_Load()
PB.Move 0, 0, 1000 * Screen.TwipsPerPixelX, 1000 * Screen.TwipsPerPixelY
TempPB.Move 0, 0, 100, 100
Preview.ScaleWidth = 10
Preview.ScaleHeight = 10
PB.ScaleWidth = 10
PB.ScaleHeight = 10
TempPB.ScaleWidth = 1
TempPB.ScaleHeight = 1
Cur.X = 1
Cur.Y = 1
Show 'layout from before showing it
CreateDefaultCourse 'make a plain course so that there's something in the preview
PreviewMode = NORMAL
RefreshCourse
End Sub


Private Sub Form_Unload(Cancel As Integer)
SForm.Visible = True
End Sub

Private Sub IDO_Click(Index As Integer)
Course.Tile(Cur.X, Cur.Y).ID = Index
PaintTile TempPB, PB, Cur.X, Cur.Y, PreviewMode
PB.Line (Cur.X - 1, Cur.Y - 1)-(Cur.X, Cur.Y), vbBlack, B
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
End Sub

Private Sub mnuBeach_Click()
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = BEACH
Next
Next
End Sub

Private Sub mnuDrivingTest_Click()
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = TEST
Next
Next
RefreshCourse
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuMuddy_Click()
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = MUDDY
Next
Next
RefreshCourse
End Sub

Private Sub mnuLoad_Click()
SCForm.ShowSave = False
SCForm.Visible = True
End Sub

Private Sub mnuNew_Click()
CreateDefaultCourse
RefreshCourse
End Sub

Private Sub mnuSave_Click()
SCForm.ShowSave = True
SCForm.Visible = True
End Sub

Public Sub RefreshCourse()
MousePointer = vbHourglass
PaintCourse TempPB, PB, PreviewMode
'paint squares on the course
For i = 1 To 9
  PB.Line (i, 0)-(i, 10), vbBlack
  PB.Line (0, i)-(10, i), vbBlack
Next

'now copy the finished level into sight
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy

'paint in the cursor position
Preview.Line (Cur.X - 1, Cur.Y - 1)-(Cur.X, Cur.Y), vbRed, B

MousePointer = vbDefault
End Sub

Private Sub mnuSea_Click()
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = SEA
Next
Next
End Sub

Private Sub mnuUrban_Click()
For X = 1 To 10
For Y = 1 To 10
   Course.Tile(X, Y).Theme = URBAN
Next
Next
End Sub

Private Sub Preview_Click()
'update cursor position
Cur.X = Mouse.X
Cur.Y = Mouse.Y
're-paint with the new cursor position
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
Preview.Line (Cur.X - 1, Cur.Y - 1)-(Cur.X, Cur.Y), vbRed, B
'update option buttons
Theme(Course.Tile(Cur.X, Cur.Y).Theme).Value = True
IDO(Course.Tile(Cur.X, Cur.Y).ID).Value = True
TileL = "Tile (" & Cur.X & "," & Cur.Y & ") Properties"
End Sub

Private Sub Preview_DblClick()
PreviewMode = Not PreviewMode 'change mode
Preview.Refresh
If PreviewMode = OUTLINE Then
  PreviewL = "Preview (Outline mode - double click the preview to change mode)"
Else
  PreviewL = "Preview (Normal mode - double click the preview to change mode)"
End If
RefreshCourse
End Sub

Private Sub Preview_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyUp
    Course.Tile(Cur.X, Cur.Y).Target = N
  Case vbKeyLeft
    Course.Tile(Cur.X, Cur.Y).Target = W
  Case vbKeyDown
    Course.Tile(Cur.X, Cur.Y).Target = S
  Case vbKeyRight
    Course.Tile(Cur.X, Cur.Y).Target = E
End Select
PaintTile TempPB, PB, Cur.X, Cur.Y, PreviewMode
PB.Line (Cur.X - 1, Cur.Y - 1)-(Cur.X, Cur.Y), vbBlack, B
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
End Sub

Private Sub Preview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Mouse.X = Int(X + 1)
Mouse.Y = Int(Y + 1)
End Sub

Private Sub Preview_Paint()
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
End Sub


Private Sub Theme_Click(Index As Integer)
Course.Tile(Cur.X, Cur.Y).Theme = Index
PaintTile TempPB, PB, Cur.X, Cur.Y, PreviewMode
PB.Line (Cur.X, Cur.Y)-(Cur.X - 1, Cur.Y - 1), vbBlack, B
StretchBlt Preview.hdc, 0, 0, 400, 400, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
End Sub
