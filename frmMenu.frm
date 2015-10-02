VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu"
   ClientHeight    =   4770
   ClientLeft      =   12120
   ClientTop       =   4665
   ClientWidth     =   7545
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7545
   Begin VB.Frame fraComputer 
      Height          =   2295
      Left            =   3120
      TabIndex        =   16
      Top             =   2280
      Width           =   3015
      Begin VB.OptionButton optComputer 
         Caption         =   "Computer"
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optHuman 
         Caption         =   "Human"
         Height          =   375
         Left            =   480
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   2175
      Left            =   6240
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   2055
      Left            =   6240
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame fraOther 
      Caption         =   "Other"
      Height          =   2175
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
      Begin VB.Image imgBackground 
         Height          =   1335
         Left            =   1440
         Picture         =   "frmMenu.frx":0442
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Board Background:"
         Height          =   855
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraSecond 
      Caption         =   "Second Player"
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   2895
      Begin VB.TextBox txtSecondName 
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "Second Player"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Game Piece:"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Image imgSecondPicture 
         Height          =   975
         Left            =   1560
         Picture         =   "frmMenu.frx":10F5
         Stretch         =   -1  'True
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "First Player"
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "First Player"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Game Piece:"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.Image imgFirstPicture 
         Height          =   975
         Left            =   1560
         Picture         =   "frmMenu.frx":2503
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraDifficulty 
      Caption         =   "Difficulty"
      Height          =   2295
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   2895
      Begin VB.OptionButton optHard 
         Caption         =   "Hard"
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton optMedium 
         Caption         =   "Medium"
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton optEasy 
         Caption         =   "Easy"
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Image imgBot 
      Height          =   500
      Left            =   2280
      Picture         =   "frmMenu.frx":2C39
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   500
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Leave
End Sub

Private Sub Leave()

    Dim Answer As Integer
    
    Answer = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
    
    If Answer = vbYes Then
        End
    End If
    
End Sub

Private Sub cmdStart_Click()
    Dim Answer As Integer
    
    'Trap for same or blank names.
    
    If FirstPicture = SecondPicture Or UCase$(Trim$(txtFirstName.Text)) = UCase$(Trim$(txtSecondName.Text)) Or Trim$(txtFirstName.Text) = "" Or Trim$(txtSecondName.Text) = "" Then
        If Mode = 0 Then
            If UCase$(Trim$(txtFirstName.Text)) = UCase$(Trim$(txtSecondName.Text)) And FirstPicture <> SecondPicture Then
                Answer = MsgBox("Your names are the same, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            ElseIf FirstPicture = SecondPicture And UCase$(txtFirstName.Text) = UCase$(txtSecondName.Text) Then
                Answer = MsgBox("Your game pieces and names are the same, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            ElseIf FirstPicture = SecondPicture Then
                Answer = MsgBox("Your game pieces look the same, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            ElseIf Trim$(txtFirstName.Text) = "" Then
                Answer = MsgBox("Your first players name is blank, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            ElseIf Trim$(txtSecondName.Text) = "" Then
                Answer = MsgBox("Your second players name is blank, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            End If
        Else
            If Trim$(txtFirstName.Text) = "" Then
                Answer = MsgBox("Your name is blank, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            ElseIf UCase$(Trim$(txtFirstName.Text)) = "COMPUTER" Then
                Answer = MsgBox("Your names are the same, are you sure you want to continue?" _
                    , vbYesNo, "Caution!")
            End If
        End If
        If Answer = vbYes Then
            Start
        End If
    Else
        Start
    End If
End Sub

'Hide the menu form and initialize the names.

Private Sub Start()
    FirstName = txtFirstName.Text
    If Mode <> 0 Then
        SecondName = "Computer"
    Else
        SecondName = txtSecondName.Text
    End If
    Me.Hide
    frmOthello.Show
End Sub

Private Sub Form_Load()

    'Initialize the variables and form controls.
    
    FirstPicture = 1
    SecondPicture = 2
    Mode = 0
    fraDifficulty.Visible = False
    optEasy.Value = True
    optHuman.Value = True
End Sub

'Display the background form.

Private Sub imgBackground_Click()
    frmBackground.Show vbModal, Me
End Sub

'Display the picture select form.

Private Sub imgFirstPicture_Click()
    Turn = 1
    frmPicture.Show vbModal, Me
End Sub

'Display the picture selecct form.

Private Sub imgSecondPicture_Click()
    Turn = 2
    frmPicture.Show vbModal, Me
End Sub

'Display the computer difficulty form.

Private Sub optComputer_Click()
    optEasy_Click
    fraDifficulty.Visible = True
    fraDifficulty.ZOrder vbBringToFront
End Sub

Private Sub optEasy_Click()
    Mode = 1
End Sub

Private Sub optHard_Click()
    Mode = 3
End Sub

Private Sub optHuman_Click()
    Mode = 0
    fraDifficulty.Visible = False
End Sub

Private Sub optMedium_Click()
    Mode = 2
End Sub
