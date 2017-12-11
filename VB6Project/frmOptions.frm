VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Serial Port and Application Settings"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   2022
         Left            =   307
         TabIndex        =   7
         Tag             =   "Sample 2"
         Top             =   305
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3420
      Index           =   0
      Left            =   840
      ScaleHeight     =   3475.161
      ScaleMode       =   0  'User
      ScaleWidth      =   5017.96
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   4965
      Begin VB.CheckBox cbShowDetailedMessages 
         Caption         =   "Show detailed pre-processor messages"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   2880
         Width           =   3615
      End
      Begin VB.ComboBox cmbParity 
         Height          =   315
         ItemData        =   "frmOptions.frx":058A
         Left            =   1320
         List            =   "frmOptions.frx":0597
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.ComboBox cmbDataBits 
         Height          =   315
         ItemData        =   "frmOptions.frx":05AC
         Left            =   1320
         List            =   "frmOptions.frx":05B6
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.ComboBox cmbStopBits 
         Height          =   315
         ItemData        =   "frmOptions.frx":05D4
         Left            =   1320
         List            =   "frmOptions.frx":05DE
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox cmbBaud 
         Height          =   315
         ItemData        =   "frmOptions.frx":05FC
         Left            =   1320
         List            =   "frmOptions.frx":060F
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cmbSerial 
         Height          =   315
         ItemData        =   "frmOptions.frx":062F
         Left            =   1320
         List            =   "frmOptions.frx":063F
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "9600"
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Odd"
         Height          =   255
         Left            =   3240
         TabIndex        =   24
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "1"
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "7"
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   2280
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   3274.56
         X2              =   5336.32
         Y1              =   243.871
         Y2              =   243.871
      End
      Begin VB.Label Recommended 
         Caption         =   "Recommended Setting"
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   0
         Width           =   2535
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Bits"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop Bits"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Parity"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Baud Rate"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Serial Port"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2490
      TabIndex        =   1
      Tag             =   "OK"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Tag             =   "&Apply"
      Top             =   4455
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   2022
         Left            =   505
         TabIndex        =   10
         Tag             =   "Sample 4"
         Top             =   502
         Width           =   2033
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3840.968
      ScaleMode       =   0  'User
      ScaleWidth      =   5745.64
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   2022
         Left            =   406
         TabIndex        =   9
         Tag             =   "Sample 3"
         Top             =   403
         Width           =   2033
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4245
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7488
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Program Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

        
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Serial port
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get SerialPort() As String
    SerialPort = Right(GetSetting(App.Title, "ComPort", "Serial", "COM1"), 1)
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Baud rate
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get Baud() As String
    Baud = GetSetting(App.Title, "ComPort", "Baud", "")
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Parity
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get Parity() As String
    Parity = Left(GetSetting(App.Title, "ComPort", "Parity", "O"), 1)
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' StopBits
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get StopBits() As String
    StopBits = Left(GetSetting(App.Title, "ComPort", "StopBits", "1"), 1)
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Databits
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get DataBits() As String
    DataBits = Left(GetSetting(App.Title, "ComPort", "DataBits", "7"), 1)
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Have we set the serial port settings?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get SerialSettingsSet() As Boolean
    If GetSetting(App.Title, "ComPort", "SettingsEntered", "") <> "True" Then
        SerialSettingsSet = False
    Else
        SerialSettingsSet = True
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show the all the dirty details?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get PreProcessorDetails() As Boolean
    If GetSetting(App.Title, "PreProcessorDetails", "Settings", 0) = 0 Then
        PreProcessorDetails = False
    Else
        PreProcessorDetails = True
    End If
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: cmdApply_Click
' PURPOSE: Apply changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdApply_Click()

    SaveSetting App.Title, "ComPort", "Serial", cmbSerial.Text
    SaveSetting App.Title, "ComPort", "Baud", cmbBaud.Text
    SaveSetting App.Title, "ComPort", "Parity", cmbParity.Text
    SaveSetting App.Title, "ComPort", "StopBits", cmbStopBits.Text
    SaveSetting App.Title, "ComPort", "DataBits", cmbDataBits.Text
    
    SaveSetting App.Title, "ComPort", "SettingsEntered", "True"
    
    SaveSetting App.Title, "PreProcessorDetails", "Settings", cbShowDetailedMessages.Value

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: cmdCancel_Click
' PURPOSE: Cancel changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdCancel_Click()
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: cmdOK_Click
' PURPOSE: Save changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub cmdOK_Click()
    cmdApply_Click
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    i = tbsOptions.SelectedItem.Index
    'handle ctrl+tab to move to the next tab
    If (Shift And 3) = 2 And KeyCode = vbKeyTab Then
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    ElseIf (Shift And 3) = 3 And KeyCode = vbKeyTab Then
        If i = 1 Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(tbsOptions.Tabs.Count)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i - 1)
        End If
    End If
End Sub


Private Sub tbsOptions_Click()
    

    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 210
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Load
' PURPOSE: Update the form with value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()

    cbShowDetailedMessages.Value = GetSetting(App.Title, "PreProcessorDetails", "Settings", 0)

    Dim i As Integer
    Dim strSetting As String
    strSetting = CStr(GetSetting(App.Title, "ComPort", "Serial", ""))
    For i = 0 To cmbSerial.ListCount - 1
        If CStr(cmbSerial.List(i)) = strSetting Then
            cmbSerial.ListIndex = i
            Exit For
        End If
    Next
    
    strSetting = CStr(GetSetting(App.Title, "ComPort", "Baud", ""))
    For i = 0 To cmbBaud.ListCount - 1
        If CStr(cmbBaud.List(i)) = strSetting Then
            cmbBaud.ListIndex = i
            Exit For
        End If
    Next
    
    strSetting = CStr(GetSetting(App.Title, "ComPort", "Parity", ""))
    For i = 0 To cmbParity.ListCount - 1
        If CStr(cmbParity.List(i)) = strSetting Then
            cmbParity.ListIndex = i
            Exit For
        End If
    Next

    strSetting = CStr(GetSetting(App.Title, "ComPort", "StopBits", ""))
    For i = 0 To cmbStopBits.ListCount - 1
        If CStr(cmbStopBits.List(i)) = strSetting Then
            cmbStopBits.ListIndex = i
            Exit For
        End If
    Next
    
    strSetting = CStr(GetSetting(App.Title, "ComPort", "DataBits", ""))
    For i = 0 To cmbDataBits.ListCount - 1
        If CStr(cmbDataBits.List(i)) = strSetting Then
            cmbDataBits.ListIndex = i
            Exit For
        End If
    Next
    
    
End Sub
