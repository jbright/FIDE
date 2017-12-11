VERSION 5.00
Begin VB.Form frmBinaryToHexCalc 
   Caption         =   "Binary/Hex Calculator"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "frmBinaryToHexCalc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbHex 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   35
      Text            =   "$0000"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox tbValue 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   32
      Text            =   "0"
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   13
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   12
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   11
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Hex:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Decimal:"
      Height          =   255
      Left            =   1680
      TabIndex        =   34
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "A15"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   120
      TabIndex        =   31
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A14"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   30
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A13"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   29
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A12"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   840
      TabIndex        =   28
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A11"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   1320
      TabIndex        =   27
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A10"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   1560
      TabIndex        =   26
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A9"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   25
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A8"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   24
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A7"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   23
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A6"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2760
      TabIndex        =   22
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A5"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   21
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A4"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3240
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A3"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   19
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A2"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3960
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   17
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "A0"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4440
      TabIndex        =   16
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   3600
      Y1              =   120
      Y2              =   600
   End
End
Attribute VB_Name = "frmBinaryToHexCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' just unload.
Private Sub btnOK_Click()
    Unload Me
End Sub

' our values have changed...
Private Sub Check1_Click(Index As Integer)

    Dim i As Integer
    Dim hexVal As Long
    
    hexVal = 0
    For i = 15 To 0 Step -1
        ' shift bits.
        hexVal = hexVal * 2
        If Check1(i).Value Then
            hexVal = hexVal + 1
        End If
    Next
    
    tbValue.Text = str(hexVal)
    tbHex.Text = Hex(hexVal)
    While Len(tbHex.Text) < 4
        tbHex.Text = "0" & tbHex.Text
    Wend
    tbHex = "$" & tbHex.Text
    

End Sub
