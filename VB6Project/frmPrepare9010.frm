VERSION 5.00
Begin VB.Form frmPrepare9010 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prepare 9010A"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmPrepare9010.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbDontPrompt 
      Caption         =   "Don't prompt again, just send the script. "
      Height          =   255
      Left            =   706
      TabIndex        =   2
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   946
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   2266
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "4) Click OK below"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "3) Press YES/ENTER"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "2) Press READ"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   2535
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      Caption         =   "1) Press AUX I/F on 9010A"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Prepare the Fluke 9010A:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   113
      X2              =   4553
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(This option can be changed back in the Script menu)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   4455
   End
End
Attribute VB_Name = "frmPrepare9010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Result As VbMsgBoxResult


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Cancel out of this dialog
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Save our settings
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdOK_Click()
    If cbDontPrompt.Value <> 0 Then
        FIDEMainModule.fMainForm.mnuScriptPromptBefore.Checked = False
    End If
    Result = vbOK
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Really ... we shouldn't have to set this because we won't get here if it's
' clicked.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
    Result = vbCancel
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ShowPrepare(ParentOwnerForm) As VbMsgBoxResult

    If Not (FIDEMainModule.fMainForm.mnuScriptPromptBefore.Checked) Then
        ShowPrepare = vbOK
    Else
        Me.Show vbModal, ParentOwnerForm
        ShowPrepare = Result
    End If
    
End Function

