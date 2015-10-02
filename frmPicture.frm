VERSION 5.00
Begin VB.Form frmPicture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Your Picture"
   ClientHeight    =   2280
   ClientLeft      =   10380
   ClientTop       =   5805
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6045
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   7
      Left            =   4920
      Picture         =   "frmPicture.frx":0000
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   6
      Left            =   3720
      Picture         =   "frmPicture.frx":0736
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   5
      Left            =   2520
      Picture         =   "frmPicture.frx":0E6C
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   4
      Left            =   1320
      Picture         =   "frmPicture.frx":15A2
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   3
      Left            =   120
      Picture         =   "frmPicture.frx":1CD8
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   2
      Left            =   1320
      Picture         =   "frmPicture.frx":240E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1005
   End
   Begin VB.Image imgPicture 
      Height          =   1005
      Index           =   1
      Left            =   120
      Picture         =   "frmPicture.frx":381C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgPicture_Click(Index As Integer)
    If Turn = 1 Then
        frmMenu!imgFirstPicture.Picture = imgPicture(Index).Picture
        FirstPicture = Index
    Else
        frmMenu!imgSecondPicture.Picture = imgPicture(Index).Picture
        SecondPicture = Index
    End If
    Me.Hide
End Sub
