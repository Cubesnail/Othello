VERSION 5.00
Begin VB.Form frmOthello 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Othello"
   ClientHeight    =   6375
   ClientLeft      =   8985
   ClientTop       =   4335
   ClientWidth     =   9435
   Icon            =   "frmOthello.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9435
   Begin VB.Frame fraTimer 
      Caption         =   "Elapsed Time"
      Height          =   1215
      Left            =   6840
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.Label lblTime 
         Caption         =   "00:00"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Time:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Timer tmrTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5160
      Top             =   720
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   10320
      TabIndex        =   4
      Top             =   840
      Width           =   6375
      Begin VB.Image imgThinking 
         Height          =   500
         Left            =   3000
         Picture         =   "frmOthello.frx":0442
         Stretch         =   -1  'True
         Top             =   3960
         Width           =   500
      End
      Begin VB.Image imgSelect 
         Height          =   500
         Left            =   840
         Picture         =   "frmOthello.frx":112B
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   500
      End
      Begin VB.Image imgInvalid 
         Height          =   1095
         Left            =   720
         Picture         =   "frmOthello.frx":2427
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Timer tmrBot 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4440
      Top             =   3960
   End
   Begin VB.Frame fraSecond 
      Caption         =   "Second Player"
      Height          =   1455
      Left            =   6840
      TabIndex        =   1
      Top             =   3000
      Width           =   2535
      Begin VB.Image imgSecond 
         Height          =   550
         Left            =   240
         Picture         =   "frmOthello.frx":314B
         Stretch         =   -1  'True
         Top             =   480
         Width           =   550
      End
      Begin VB.Label lblSecondScore 
         Caption         =   "SecondScore"
         Height          =   615
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Frame fraFirst 
      Caption         =   "First Player"
      Height          =   1455
      Left            =   6840
      TabIndex        =   0
      Top             =   1440
      Width           =   2535
      Begin VB.Image imgFirst 
         Height          =   500
         Left            =   240
         Picture         =   "frmOthello.frx":4559
         Stretch         =   -1  'True
         Top             =   480
         Width           =   500
      End
      Begin VB.Label lblFirstScore 
         Caption         =   "FirstScore"
         Height          =   615
         Left            =   960
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Image imgBackground 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   550
      Left            =   0
      Picture         =   "frmOthello.frx":4C8F
      Stretch         =   -1  'True
      Top             =   45
      Visible         =   0   'False
      Width           =   550
   End
   Begin VB.Image imgPiece 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Index           =   0
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuMenu 
         Caption         =   "Menu"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "Highscores"
      End
      Begin VB.Menu mnuChangelog 
         Caption         =   "Changelog"
      End
   End
End
Attribute VB_Name = "frmOthello"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:Derrick Liang
'Purpose: The purpose of this application is to run a game of othello.
'Date: Tuesday, June 2nd, 2015
'Files: frmAbout.frm, frmAbout.frx, frmBackground.frm, frmBackground.frx, frmChangelog.frm, frmChangelog.frx, frmHighscore.frm
    'frmHighscore.frx, frmMenu.frm, frmMenu.frx, frmPicture.frm, frmPicture.frx, frmSplash.frm, frmSplash.frx, modOthello.bas
        'vbpOthello.vbp
Option Explicit

'Declare the form constants and variables

Const MAX = 10
Dim FirstScore As Integer
Dim SecondScore As Integer
Private Elapsed As Integer
Private BotMode As Integer
Private Bot As Move

'Initalize the form controls and score.

Sub Initialize()
    
    Dim K As Integer
    Dim M As Integer
    
    ReadScore
    
    If Mode <> 0 Then
        mnuHighScores.Enabled = True
    Else
        mnuHighScores.Enabled = False
    End If
    
    CellWidth = 550
    
    For K = 1 To MAX
        For M = 1 To MAX
            Load imgPiece(XYtoNum(K, M))
            With imgPiece(XYtoNum(K, M))
                .BorderStyle = 1
                .Appearance = 0
                .Enabled = True
                .Visible = True
                .Stretch = True
                .Height = CellWidth
                .Width = CellWidth
                .ZOrder vbBringToFront
                .Left = K * CellWidth
                .Top = M * CellWidth
            End With
        Next M
    Next K
    
    With imgBackground
        .Height = CellWidth * MAX
        .Width = CellWidth * MAX
        .ZOrder vbSendToBack
        .BorderStyle = 1
                .Appearance = 0
                .Enabled = True
                .Visible = True
                .Stretch = True
                .Left = CellWidth
                .Top = CellWidth
    End With
    With frmMenu
        imgBackground.Picture = !imgBackground.Picture
        imgFirst.Picture = !imgFirstPicture.Picture
        If Mode = 0 Then
            imgSecond.Picture = !imgSecondPicture.Picture
        Else
            imgSecond.Picture = !imgBot.Picture
            Elapsed = 0
        End If
    End With
    fraFirst.Caption = FirstName
    fraSecond.Caption = SecondName
End Sub

Private Sub Form_Load()
    Initialize
    StartGame
    UpdateBoard
End Sub

Private Sub UpdateBoard()
    
    Dim K As Integer
    Dim M As Integer
    Dim FirstScore As Integer
    Dim SecondScore As Integer
    Dim Amount As Integer
    
    'Check all imageboxes and replace with their respective images.
    
    For K = 1 To MAX
        For M = 1 To MAX
            If PieceMove(K, M) = 1 Then
                imgPiece(XYtoNum(K, M)).Picture = imgFirst.Picture
            ElseIf PieceMove(K, M) = 2 Then
                imgPiece(XYtoNum(K, M)).Picture = imgSecond.Picture
            Else
                imgPiece(XYtoNum(K, M)).Picture = Nothing
            End If
        Next M
    Next K
    
    'Change the background of the
    
    If Turn = 1 Then
        lblSecondScore.BackColor = &H8000000F
        lblFirstScore.BackColor = vbWhite
        fraSecond.BackColor = &H8000000F
        fraFirst.BackColor = vbWhite
        If Mode = 1 Then
            For K = 1 To GetMove()
                imgPiece(XYtoNum(MoveList(K).X, MoveList(K).Y)).Picture = imgSelect.Picture
            Next K
        End If
    Else
        lblFirstScore.BackColor = &H8000000F
        lblSecondScore.BackColor = vbWhite
        fraFirst.BackColor = &H8000000F
        fraSecond.BackColor = vbWhite
    End If
    
    Checkscore FirstScore, SecondScore
    lblFirstScore.Caption = "Score: " & FirstScore
    lblSecondScore.Caption = "Score: " & SecondScore
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMenu.Show
End Sub

Private Sub imgPiece_Click(Index As Integer)
    Dim Bot As Move
    Dim Check As Boolean
    
    'Check for different gamemodes.
    
    If Mode <> 0 Then
        
        If Turn = 1 Then
            Check = PutPiece(NumtoX(Index), NumtoY(Index))
            If Not Check And PieceMove(NumtoX(Index), NumtoY(Index)) = 0 Then
                imgPiece(Index).Picture = imgInvalid.Picture
            End If
        End If
        
        If CheckEnd() Then
            EndGame
            Exit Sub
        End If
        
        If Check Then
            BotMode = 0
            tmrBot.Enabled = True
            tmrTimer.Enabled = True
        End If
    Else
        Check = PutPiece(NumtoX(Index), NumtoY(Index))
        If Not Check And PieceMove(NumtoX(Index), NumtoY(Index)) = 0 Then
            imgPiece(Index).Picture = imgInvalid.Picture
        End If
    End If
    
    If CheckEnd() Then
        EndGame
        Exit Sub
    End If
    
End Sub

'This function places a piece and returns true if values have been changed.

Function PutPiece(X As Integer, Y As Integer) As Boolean
    Dim Check As Boolean
    Dim K As Integer
    Dim Temp As Boolean
    Dim Count As Integer
    K = 1
    
    Check = False
    If X = 0 Or Y = 0 Then
        EndGame
        Exit Function
    End If
    Temp = False
    If PieceMove(X, Y) = 0 Then
        Count = 0
        For K = 1 To 8
            Check = Valid(Turn, X, Y, K, True, , , Count)
            If Check Then
                PieceMove(X, Y) = Turn
                Temp = True
            End If
        Next
    End If
    
    If Temp = True Then
        If Turn = 1 Then
            Turn = 2
        Else
            Turn = 1
        End If
        UpdateBoard
    End If
    
    If CheckEnd() Then
        EndGame
        Exit Function
    End If
    
    PutPiece = Temp
End Function

Sub EndGame()

    Dim Answer As Integer
    Dim Temp As Score
    Dim K As Integer
    
    tmrTimer.Enabled = False
    Checkscore FirstScore, SecondScore
    
    'Check userscores and determine the winner.
    
    If FirstScore > SecondScore Then
    
        'Trap for blank names.
        
        If FirstName <> "" Then
            MsgBox FirstName & " wins!", , "Winner!"
        Else
            MsgBox "First player wins!", , "Winner!"
        End If
    ElseIf SecondScore > FirstScore Then
        If SecondName <> "" Then
            MsgBox SecondName & " wins!", , "Winner!"
        Else
            MsgBox "Second player wins!", , "Winner!"
        End If
    Else
        MsgBox "Tie!"
    End If
    
    'Check for different gamemodes.
    
    If Mode <> 0 And FirstScore > SecondScore Then
        With Scores(Mode, 6)
                .Time = Elapsed
            If FirstName = "" Then
                .Name = "Anonymous"
            Else
                .Name = FirstName
            End If
            .FirstScore = FirstScore
            .SecondScore = SecondScore
            
            'Check for high scores.
            
            If .Time < Scores(Mode, 5).Time Or Scores(Mode, 5).Time = 0 Then
                MsgBox "You made it to the top 5 scores!", , "High Score!"
            End If
        End With
        
        'Update high scores.
        
        For K = 6 To 2 Step -1
            If Scores(Mode, K).Time < Scores(Mode, K - 1).Time Or Scores(Mode, K - 1).Time = 0 Then
                Temp = Scores(Mode, K)
                Scores(Mode, K) = Scores(Mode, K - 1)
                Scores(Mode, K - 1) = Temp
            End If
        Next K
        
        frmHighScore.Show vbModal, Me
        SaveScore
        
    End If
    
    'Check for play again.
        
    Answer = MsgBox("Would you like to play again?", vbYesNo, "End!")
        
    If Answer = vbYes Then
        StartGame
        UpdateBoard
    Else
        frmMenu.Show
        Unload Me
    End If
    lblTime.Caption = "00:00"
    
End Sub

Sub Checkscore(ByRef First As Integer, ByRef Second As Integer)
    Dim K As Integer
    Dim M As Integer
    Dim TempFirst As Integer
    Dim TempSecond As Integer
    TempFirst = 0
    TempSecond = 0
    
    'Count all pieces on the board.
    
    For K = 1 To MAX
        For M = 1 To MAX
            If PieceMove(K, M) = 1 Then
                TempFirst = TempFirst + 1
            ElseIf PieceMove(K, M) = 2 Then
                TempSecond = TempSecond + 1
            End If
        Next M
    Next K
    
    First = TempFirst
    Second = TempSecond
            
End Sub
Sub StartGame()
    Dim K As Integer
    Dim M As Integer
    Dim Temp As Integer
    
    'Initialize piece placements.
    
    Randomize
    
    For K = 1 To MAX
        For M = 1 To MAX
            PieceMove(K, M) = 0
        Next M
    Next K
    
    Turn = 1
    
    Temp = MAX / 2
    
    PieceMove(Temp, Temp) = Int(Rnd() * 2 + 1)
    PieceMove(Temp + 1, Temp) = PieceMove(Temp, Temp) Mod 2 + 1
    PieceMove(Temp, Temp + 1) = PieceMove(Temp, Temp) Mod 2 + 1
    PieceMove(Temp + 1, Temp + 1) = PieceMove(Temp, Temp)
    
    Elapsed = 0
    
    If Mode <> 0 Then
        fraTimer.Visible = True
    Else
        fraTimer.Visible = False
    End If
End Sub

'Display the about form.

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'Display the changelog form.

Private Sub mnuChangelog_Click()
    frmChangelog.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    
    Dim Answer As Integer
    
    'Check if the user wants to exit.
    
    Answer = MsgBox("Your game in progress will not be saved! Are you sure you want to exit?", vbYesNo, "Warning!")
    
    If Answer = vbYes Then
        End
    End If
    
End Sub

'Display the high score form.

Private Sub mnuHighScores_Click()
    frmHighScore.Show vbModal, Me
End Sub

Private Sub mnuMenu_Click()
    Dim Answer As Integer
    
    'Check if the user wants to return to the menu.
    
    Answer = MsgBox("Your game in progress will not be saved! Are you sure you want to return to the main menu?", vbYesNo, "Warning!")
    
    If Answer = vbYes Then
        frmMenu.Show
        Unload Me
    End If
End Sub

Private Sub mnuNew_Click()
    Dim Answer As Integer
    
    'Check if the user wants to restart the game.
    
    Answer = MsgBox("Would you like to play again? All progress will be lost!", vbYesNo, "Restart")
    
    If Answer = vbYes Then
        StartGame
        UpdateBoard
    End If
End Sub

Private Sub tmrBot_Timer()
    
    'Indicate and display where the computer will move.
    
    If BotMode = 0 Then
        tmrTimer.Enabled = False
        If Mode = 1 Then
                Bot = EasyBot()
        ElseIf Mode = 2 Then
                Bot = MediumBot()
        ElseIf Mode = 3 Then
                Bot = Hardbot()
        End If
        If CheckEnd() Then
            EndGame
            Exit Sub
        End If
        imgPiece(XYtoNum(Bot.X, Bot.Y)).Picture = imgThinking.Picture
        BotMode = 1
    Else
        If Not PutPiece(Bot.X, Bot.Y) Then
            EndGame
            Exit Sub
        End If
        UpdateBoard
        BotMode = 0
        tmrBot.Enabled = False
        Turn = 1
        If CheckEnd() Then
            EndGame
            Exit Sub
        End If
        tmrTimer.Enabled = True
        
    End If
    If CheckEnd() Then
        EndGame
        Exit Sub
    End If
End Sub

Private Sub tmrTimer_Timer()
    Dim Converted As String
    Elapsed = Elapsed + 1
    
    'Display the elapsed time.
    
    Converted = Format$(Str$(Elapsed \ 60), "00") & ":" & Format$(Str$(Elapsed Mod 60), "00")
    
    lblTime.Caption = Converted
End Sub
