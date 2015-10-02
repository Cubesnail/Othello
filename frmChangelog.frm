VERSION 5.00
Begin VB.Form frmChangelog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changelog"
   ClientHeight    =   2115
   ClientLeft      =   12300
   ClientTop       =   5820
   ClientWidth     =   3675
   Icon            =   "frmChangelog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3675
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "2.0: Version caption changed from 1.0 to 2.0."
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "1.0: Game Created."
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Caption         =   "Changelog:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmChangelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Unload Me
End Sub
