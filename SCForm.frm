VERSION 5.00
Begin VB.Form SCForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load / Save Course"
   ClientHeight    =   4836
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   4284
   Icon            =   "SCForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4836
   ScaleWidth      =   4284
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   2880
      TabIndex        =   4
      Top             =   4320
      Width           =   1332
   End
   Begin VB.CommandButton cmdSaveLoad 
      Caption         =   "Load/Save"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Width           =   1452
   End
   Begin VB.TextBox FileNameT 
      Height          =   288
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   4092
   End
   Begin VB.FileListBox File1 
      Height          =   3528
      Left            =   120
      Pattern         =   "*.hcc"
      TabIndex        =   0
      Top             =   120
      Width           =   4092
   End
   Begin VB.Label FilenameL 
      Caption         =   "Filename :"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2172
   End
End
Attribute VB_Name = "SCForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ShowSave As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSaveLoad_Click()
TryAgain:
If ShowSave Then
  If SaveCourse(File1.Path & "\" & FileNameT & ".hcc") = True Then
     Unload Me
  Else
        If MsgBox("There was an error when attempting to save " & FileNameT & ". Try again?", vbRetryCancel + vbExclamation, "Errors occured during save") = vbRetry Then
           GoTo TryAgain
        Else
           Unload Me
        End If
  End If
Else
  If LoadCourse(File1.Path & "\" & File1.FileName) = True Then
     Unload Me
  Else
        If MsgBox("There was an error when attempting to load " & FileNameT & ". Try again?", vbRetryCancel + vbExclamation, "Errors occured during save") = vbRetry Then
           GoTo TryAgain
        Else
           Unload Me
        End If
  End If
End If

DForm.RefreshCourse
End Sub

Private Sub File1_Click()
FileNameT = Left(File1.FileName, Len(File1.FileName) - 4)
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\Resources\Courses\"
If ShowSave Then
  Caption = "Save Course"
  cmdSaveLoad.Caption = "Save"
Else
  Caption = "Load Course"
  cmdSaveLoad.Caption = "Load"
End If
End Sub
