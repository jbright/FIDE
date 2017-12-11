VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5790
   ScaleWidth      =   8145
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7680
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   0
      Picture         =   "frmSplash.frx":9D99
      ScaleHeight     =   5535
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   1
      Tag             =   "Version"
      Top             =   5520
      Width           =   3330
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public TimeUp As Boolean

Private Sub Form_Load()
    TimeUp = False
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
   ' lblProductName.Caption = App.Title
End Sub

Private Sub Timer1_Timer()
    TimeUp = True
End Sub
