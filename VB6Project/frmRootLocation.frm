VERSION 5.00
Begin VB.Form frmRootLocation 
   Caption         =   "Fluke IDE Root Project Location"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   Icon            =   "frmRootLocation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6105
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox CurrentDirectory 
      Height          =   3015
      Left            =   360
      TabIndex        =   4
      Top             =   960
      Width           =   4095
   End
   Begin VB.DriveListBox CurrentDrive 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Select the location for the root project:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmRootLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Has our setting changed?
Private Changed  As Boolean

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' For the outside world.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get ProjectPath() As String
    ProjectPath = GetSetting(App.Title, "Settings", "ProjectPath", App.Path & "\Scripts")
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' For the outside world.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Let ProjectPath(strNewValue As String)
    SaveSetting App.Title, "Settings", "ProjectPath", strNewValue
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetProjectPath(ParentOwnerForm) As String

    Me.Show vbModal, ParentOwnerForm
    GetProjectPath = ProjectPath
    
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Okay, let's do this.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub OKButton_Click()
    Changed = True
    ProjectPath = CurrentDirectory.Path
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Okay, don't
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub CurrentDrive_Change()
    On Error Resume Next
    CurrentDirectory.Path = CurrentDrive.Drive
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Starting condition.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
    CurrentDrive.Drive = Left(ProjectPath, 1)
    On Error Resume Next
    CurrentDirectory.Path = ProjectPath
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ShowModal(ParentOwnerForm) As Boolean

    Changed = False
    Me.Show vbModal, ParentOwnerForm
    ShowModal = Changed
    
End Function


