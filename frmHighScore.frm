VERSION 5.00
Begin VB.Form frmHighScore 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   3270
   ClientLeft      =   10455
   ClientTop       =   5205
   ClientWidth     =   6225
   Icon            =   "frmHighScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6225
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2760
      Width           =   4815
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   ">>"
      Height          =   1095
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<<"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   375
   End
   Begin VB.PictureBox picScores 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   600
      ScaleHeight     =   2355
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmHighScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempMode As Integer

Private Sub cmdBack_Click()
    If TempMode = 1 Then
        TempMode = 3
    Else
        TempMode = TempMode - 1
    End If
    LoadScores TempMode
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdForward_Click()
    If TempMode = 3 Then
        TempMode = 1
    Else
        TempMode = TempMode + 1
    End If
    LoadScores TempMode
End Sub

Private Sub Form_Load()
    TempMode = Mode
    LoadScores TempMode
End Sub

Private Sub LoadScores(GameMode As Integer)
    Dim K As Integer
    
    picScores.Cls
    
    If GameMode = 1 Then
        picScores.Print "EasyBot"
    ElseIf GameMode = 2 Then
        picScores.Print "MediumBot"
    Else
        picScores.Print "HardBot"
    End If
    
    picScores.Print Tab(5); "Name";
    picScores.Print Tab(30); "Time";
    picScores.Print Tab(45); "Score"
    For K = 1 To 500
        picScores.Print "-";
    Next K
    picScores.Print
    For K = 1 To 5
        With Scores(GameMode, K)
            picScores.Print Str$(K) & ".";
            picScores.Print Tab(5); Trim$(.Name);
            picScores.Print Tab(30); Format$(Str$(.Time \ 60), "00") & ":" & Format$(Str$(.Time Mod 60), "00");
            picScores.Print Tab(45); Format$(.FirstScore, "000"); "-"; Format$(.SecondScore, "000")
            picScores.Print
        End With
    Next K
    
End Sub
