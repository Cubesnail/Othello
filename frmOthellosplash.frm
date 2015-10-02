VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4230
   ClientLeft      =   10380
   ClientTop       =   5775
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmOthellosplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer tmrCloseForm 
         Interval        =   3000
         Left            =   5760
         Top             =   960
      End
      Begin VB.Image imgLogo 
         Height          =   2625
         Left            =   360
         Picture         =   "frmOthellosplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   2895
      End
      Begin VB.Label lblCompany 
         Caption         =   "Derrick Liang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   1
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   2
         Top             =   3360
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   3360
         TabIndex        =   3
         Top             =   1320
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    CloseForm
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version 2.0"
    lblProductName.Caption = "Othello"
End Sub

Private Sub CloseForm()
    frmMenu.Show
    Unload Me
End Sub

Private Sub tmrCloseForm_Timer()
    CloseForm
End Sub
