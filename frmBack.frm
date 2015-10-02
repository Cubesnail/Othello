VERSION 5.00
Begin VB.Form frmBackground 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select your background"
   ClientHeight    =   4470
   ClientLeft      =   10290
   ClientTop       =   4710
   ClientWidth     =   6615
   Icon            =   "frmBack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6615
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   5
      Left            =   4440
      Picture         =   "frmBack.frx":0442
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   4
      Left            =   2280
      Picture         =   "frmBack.frx":3F42
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   3
      Left            =   120
      Picture         =   "frmBack.frx":1B893
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   2
      Left            =   4440
      Picture         =   "frmBack.frx":1C546
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   1
      Left            =   2280
      Picture         =   "frmBack.frx":1D7CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgBackground 
      Height          =   2055
      Index           =   0
      Left            =   120
      Picture         =   "frmBack.frx":2326D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgBackground_Click(Index As Integer)
    frmMenu!imgBackground.Picture = imgBackground(Index).Picture
    Me.Hide
End Sub
