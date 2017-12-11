VERSION 5.00
Begin VB.Form dlgAddNewFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Folder..."
   ClientHeight    =   1125
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "dlgAddNewFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFolderName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Folder:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "dlgAddNewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private NewFolderName As String

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetNewFolder(ParentOwnerForm) As String

    Me.Show vbModal, ParentOwnerForm
    GetNewFolder = NewFolderName
    
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Okay, don't
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CancelButton_Click()
    NewFolderName = ""
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Starting condition.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
    NewFolderName = ""
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Okay, let's do this.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub OKButton_Click()
    NewFolderName = txtFolderName.Text
    Unload Me
End Sub
