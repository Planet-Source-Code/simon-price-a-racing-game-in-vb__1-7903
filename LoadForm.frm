VERSION 5.00
Begin VB.Form LoadForm 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1128
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3492
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   94
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox LoadBar 
      BackColor       =   &H000000FF&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   240
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      ToolTipText     =   "Loading.. 0% Complete"
      Top             =   600
      Width           =   3012
   End
End
Attribute VB_Name = "LoadForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, _
  ByVal cx As Long, ByVal cy As Long, _
  ByVal wFlags As Long) As Long

Public Progress As Byte
Public FileName As String

Public Sub LoadIt()
Visible = True
Show
DoEvents
Progress = 5
LoadBar.Refresh
BuildTrigTable 'pre-calculate sin + cos values
Progress = 10
LoadBar.Refresh
DoEvents
If LoadCourse(SelectForm.File1.Path & "\" & SelectForm.File1.FileName) Then
    Progress = 25
    LoadBar.Refresh
    PaintCourse GForm.TempPB, GForm.LevelPic, True
    Progress = 40
    LoadBar.Refresh
    DoEvents
    GForm.LoadCourseOutline  'load the collision map
    Progress = 65
    LoadForm.LoadBar.Refresh
    GForm.LevelPic.Picture = GForm.LevelPic.Image
    GForm.Map.Picture = GForm.Map.Image
    GForm.Visible = True
    Progress = 80
    LoadForm.LoadBar.Refresh
    GForm.SetUpLevel
    CalcDispSize
    Progress = 90
    LoadForm.LoadBar.Refresh
    PlaceCarsAtStart
    Progress = 100
    LoadForm.LoadBar.Refresh
    GForm.GoAhead
Else
  MsgBox "Error occured during loading!", , "Error!"
End If
End Sub

Private Sub Form_Load()
Picture = LoadPicture(App.Path & "\Resources\Pictures\Misc\Loading.bmp")
SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 300, 100, SWP_NOMOVE + SWP_NOSIZE
Show
End Sub

Private Sub LoadBar_Paint()
LoadBar.Line (0, 0)-(Progress, 1), vbBlue, BF
LoadBar.ToolTipText = "Loading... " & Progress & "% complete"
End Sub

Public Sub PlaceCarsAtStart()
On Error Resume Next
'workout where the cars start
For X = 0 To 9
For Y = 0 To 9
   If Course.Tile(X + 1, Y + 1).ID = STARTGRID Then
     Select Case Course.Tile(X, Y).Target
       Case N
         For i = 1 To UBound(Car) Step 2
             Car(i).Angle = 0
             Car(i + 1).Angle = 0
             Car(i).X = X * 100 + 40
             Car(i + 1).X = X * 100 + 60
             Car(i).Y = Y * 100 + (i - 1.5) * 25
             Car(i + 1).Y = Y * 100 + (i - 1.5) * 25
         Next
       Case S
         For i = 1 To UBound(Car) Step 2
             Car(i).Angle = 18
             Car(i + 1).Angle = 18
             Car(i).X = X * 100 + 40
             Car(i + 1).X = X * 100 + 60
             Car(i).Y = Y * 100 + (i - 1.5) * 25
             Car(i + 1).Y = Y * 100 + (i - 1.5) * 25
         Next
       Case E
         For i = 1 To UBound(Car) Step 2
             Car(i).Angle = 9
             Car(i + 1).Angle = 9
             Car(i).Y = Y * 100 + 40
             Car(i + 1).Y = Y * 100 + 60
             Car(i).X = X * 100 + (i - 1.5) * 25
             Car(i + 1).X = X * 100 + (i - 1.5) * 25
         Next
       Case W
         For i = 1 To UBound(Car) Step 2
             Car(i).Angle = 27
             Car(i + 1).Angle = 27
             Car(i).Y = Y * 100 + 40
             Car(i + 1).Y = Y * 100 + 60
             Car(i).X = X * 100 + (i - 1.5) * 25
             Car(i + 1).X = X * 100 + (i - 1.5)
         Next
       End Select
    End If
Next
Next
End Sub
