VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSignature 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fluke Signature Analysis"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12690
   Icon            =   "frmSignature.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   11400
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   7455
      Left            =   3720
      TabIndex        =   13
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13150
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   16000
      TextRTF         =   $"frmSignature.frx":173A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Script generation mode"
      Height          =   1100
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   3495
      Begin VB.OptionButton op0 
         Caption         =   "ROM check only"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton op2 
         Caption         =   "Check each address bit"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   800
         Width           =   3015
      End
      Begin VB.OptionButton op1 
         Caption         =   "Check base address"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   520
         Width           =   2775
      End
   End
   Begin VB.CheckBox bEachProgram 
      Caption         =   "Make each ROM test its own program"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CommandButton btnCopyToFile 
      Caption         =   "Copy to Current Script File"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   7680
      Width           =   7575
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton btnCalculateSig 
      Caption         =   "&Calculate Signature"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   3495
   End
   Begin VB.FileListBox CurrentFileList 
      Height          =   2430
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   3600
      Width           =   3495
   End
   Begin VB.DriveListBox CurrentDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.DirListBox CurrentDirectory 
      Height          =   1890
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Assumes starting address of 0000 (each file). Enter value in Hex as you would on the Fluke. 0000-0000 Calcs for entire file."
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Test Range 0000 - "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2580
      Width           =   1455
   End
End
Attribute VB_Name = "frmSignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NewCode As String

Private LabelCnt As Integer
Private LastLabel As String

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: GetNextLabel
' PURPOSE: Gets the next available label for us to use
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetNextLabel() As String
    LastLabel = "L" & LabelCnt
    GetNextLabel = LastLabel
    LabelCnt = LabelCnt + 1
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: HexLine
' PURPOSE: Gets a line of data from the given file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function HexLine(base As Integer) As String

    HexLine = _
    HexString(CByte(0)) & HexString(CByte(base)) & ":  " & _
    HexString(ByteData(base + 0)) & " " & _
    HexString(ByteData(base + 1)) & " " & _
    HexString(ByteData(base + 2)) & " " & _
    HexString(ByteData(base + 3)) & " " & _
    HexString(ByteData(base + 4)) & " " & _
    HexString(ByteData(base + 5)) & " " & _
    HexString(ByteData(base + 6)) & " " & _
    HexString(ByteData(base + 7)) & "    " & _
    HexString(ByteData(base + 8)) & " " & _
    HexString(ByteData(base + 9)) & " " & _
    HexString(ByteData(base + 10)) & " " & _
    HexString(ByteData(base + 11)) & " " & _
    HexString(ByteData(base + 12)) & " " & _
    HexString(ByteData(base + 13)) & " " & _
    HexString(ByteData(base + 14)) & " " & _
    HexString(ByteData(base + 15)) & " "

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: WriteErrorBlock
' PURPOSE: Tell the user what has failed
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub WriteErrorBlock(nIndex As Integer, strAddress As String)
    AddTextLine vbTab & "read @ " & strAddress
    AddTextLine vbTab & "if DAT = " & Hex(FlukeSignature.ByteData(nIndex)) & " goto " & GetNextLabel
    AddTextLine vbTab & "aux ERROR @ " & strAddress & " +"
    AddTextLine vbTab & "aux , EXPECT " & Hex(FlukeSignature.ByteData(nIndex)) & " +"
    AddTextLine vbTab & "aux , GOT $DAT +"
    AddTextLine vbTab & "goto ErrCond"
    AddTextLine " "
    AddTextLine LastLabel & ":"
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ProcessSignatureFile
' PURPOSE: Processes an individual file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ProcessSignatureFile(strDirectory As String, strFile As String)

    Dim strFileAndDir As String
    
    If Right(strDirectory, 1) = "\" Then
        strFileAndDir = strDirectory & strFile
    Else
        strFileAndDir = strDirectory & "\" & strFile
    End If
    
    ' Get the signature of the file
    Dim hSig As Long
    hSig = FlukeSignature.Signature(strFileAndDir, CLng("&H" & txtSize.Text))
    
    ' Is this its own program?
    If bEachProgram Then
        ' Reset label count
        LabelCnt = 0
        AddTextLine " "
        AddTextLine "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        AddTextLine "!! Program: " & strFile
        AddTextLine "!! Purpose: ROM test based on file "
        AddTextLine "!! " & vbTab & HexLine(0)
        AddTextLine "!! " & vbTab & HexLine(16)
        AddTextLine "!! " & vbTab & HexLine(32)
        AddTextLine "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
        AddTextLine "program " & strFile
        
        AddTextLine vbTab & "dpy ROM test @ 0000-" & Hex(UBound(FlukeSignature.ByteData)) & " sig " & Hex(hSig)
        AddTextLine vbTab & "aux ROM test @ 0000-" & Hex(UBound(FlukeSignature.ByteData)) & " sig " & Hex(hSig) & " +"
        AddTextLine vbTab
    End If
    
    If op1.Value = True Then
        WriteErrorBlock 0, "0000"
    ElseIf op2.Value = True Then
    
        Dim i As Long, nNextTestLine As Long
        For i = LBound(FlukeSignature.ByteData) To UBound(FlukeSignature.ByteData)
        
            If i = 0 Then
                WriteErrorBlock 0, "0000"
                nNextTestLine = 1
            ElseIf i = nNextTestLine Then
                WriteErrorBlock CInt(i), Hex(i)
                nNextTestLine = nNextTestLine * 2
            End If
        
        Next
        
    End If
    
    
    AddTextLine vbTab & "ROM test @ 0000-" & Hex(UBound(FlukeSignature.ByteData)) & " sig " & Hex(hSig) & vbTab & "! from file: " & strFile
    
    
    If bEachProgram And Not (op0.Value) Then
        AddTextLine vbTab & "goto ProgDone"
        AddTextLine " "
        AddTextLine "ErrCond:"
        AddTextLine vbTab & "!! sound bell; you can change this to ask for input"
        AddTextLine vbTab & "!! right now it logs it to the aux channel and continues"
        AddTextLine vbTab & "dpy # ERROR #"
        
        AddTextLine " "
        AddTextLine "ProgDone:"
        AddTextLine vbTab & "aux DONE"
        
    Else
        If op1.Value = True Or op2.Value = True Then
            AddTextLine vbTab & "!! TODO: You must define ErrCond "
        End If
    End If

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddTextLine
' PURPOSE: Adds a line to the text file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddTextLine(str As String)
    txtCode.Text = txtCode.Text & str & vbCrLf
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Load
' PURPOSE: Built in call. We get this when this form first starts up.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnCalculateSig_Click()
    txtCode.Text = "" & vbCrLf
    
    Dim i As Integer
    LabelCnt = 0
    
    ' Process each file
    With CurrentFileList
        For i = 0 To .ListCount - 1 'scan the entire list
            If .Selected(i) Then
                ProcessSignatureFile CurrentDirectory.Path, .List(i)
            End If
        Next
    End With
    
    ColorSyntax.CheckRange txtCode, 0, Len(txtCode.Text)
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: btnCancel_Click
' PURPOSE: Don't make our changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnCancel_Click()
    NewCode = ""
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: btnCopyToFile_Click
' PURPOSE: Copies the code, prepares to close diaglog
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnCopyToFile_Click()
    NewCode = txtCode.Text
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentDirectory_Change
' PURPOSE: Update the file list when the directory changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentDirectory_Change()
    On Error Resume Next
    CurrentFileList.Path = CurrentDirectory.Path
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentDrive_Change
' PURPOSE: Change the directory when the drive changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentDrive_Change()
    On Error Resume Next
    CurrentDirectory.Path = CurrentDrive.Drive
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: HexString
' PURPOSE: Returns xx in hex form
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function HexString(b As Byte) As String

    If b > 15 Then
        HexString = Hex(b)
    ElseIf b = 0 Then
        HexString = "00"
    Else
        HexString = "0" & Hex(b)
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GenerateCode(ParentOwnerForm) As String
    NewCode = ""
    Me.Show vbModal, ParentOwnerForm
    GenerateCode = NewCode
End Function


