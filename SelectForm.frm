VERSION 5.00
Begin VB.Form SelectForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Your HoverCar and Course!"
   ClientHeight    =   5400
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   5640
   Icon            =   "SelectForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGO 
      Caption         =   "GO ! ! !"
      Height          =   372
      Left            =   1800
      TabIndex        =   7
      Top             =   4920
      Width           =   2052
   End
   Begin VB.Frame CourseF 
      Caption         =   "Choose a course :"
      Height          =   4692
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5412
      Begin VB.CheckBox ShowIt 
         Caption         =   "Show Preview"
         Height          =   252
         Left            =   2400
         TabIndex        =   5
         Top             =   3600
         Value           =   1  'Checked
         Width           =   2652
      End
      Begin VB.PictureBox Preview 
         Height          =   2892
         Left            =   2400
         ScaleHeight     =   237
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   237
         TabIndex        =   2
         Top             =   600
         Width           =   2892
         Begin VB.PictureBox TempPB 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1200
            Left            =   360
            ScaleHeight     =   1
            ScaleMode       =   0  'User
            ScaleWidth      =   1
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.PictureBox PB 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFF00&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            DrawStyle       =   1  'Dash
            DrawWidth       =   3
            Height          =   12000
            Left            =   480
            ScaleHeight     =   10
            ScaleMode       =   0  'User
            ScaleWidth      =   10
            TabIndex        =   3
            Top             =   1440
            Visible         =   0   'False
            Width           =   12000
         End
      End
      Begin VB.FileListBox File1 
         Height          =   4296
         Left            =   120
         Pattern         =   "*.hcc"
         TabIndex        =   1
         Top             =   240
         Width           =   2172
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Here is a list of available courses - it includes both the normal courses and your custom courses!"
         Height          =   612
         Left            =   2400
         TabIndex        =   6
         Top             =   3960
         Width           =   2892
      End
      Begin VB.Label Label1 
         Caption         =   "Preview of course :"
         Height          =   252
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   2052
      End
   End
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGO_Click()
If File1.FileName = "" Then
   MsgBox "Please select a course first!", vbExclamation, "No course selected!"
   Exit Sub
End If
LoadForm.FileName = File1.Path & "\" & File1.FileName
LoadForm.Visible = True
Hide
LoadForm.LoadIt
Unload Me
End Sub

Private Sub File1_Click()
If ShowIt Then ShowPreview
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\resources\Courses\"
End Sub

Private Sub ShowPreview()
MousePointer = vbHourglass
LoadCourse File1.Path & "\" & File1.FileName
LoadCourse File1.Path & "\" & File1.FileName
  PaintCourse TempPB, PB, True
  StretchBlt Preview.hdc, 0, 0, 200, 200, PB.hdc, 0, 0, 1000, 1000, vbSrcCopy
MousePointer = vbDefault
End Sub
