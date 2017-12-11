VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmROMDiff 
   Caption         =   "ROM Set Identifier"
   ClientHeight    =   7440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12165
   Icon            =   "frmROMDiff.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7440
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCopyToFile 
      Caption         =   "Copy to Current Script File"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   6960
      Width           =   8295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files to be analzed"
      Height          =   3015
      Left            =   3720
      TabIndex        =   5
      Top             =   120
      Width           =   8295
      Begin VB.TextBox txtBaseAddr 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "0"
         Top             =   2565
         Width           =   615
      End
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   2520
         Width           =   2415
      End
      Begin VB.ListBox lstFiles 
         Height          =   1425
         ItemData        =   "frmROMDiff.frx":173A
         Left            =   120
         List            =   "frmROMDiff.frx":173C
         TabIndex        =   6
         Top             =   240
         Width           =   8055
      End
      Begin VB.Label Label2 
         Caption         =   $"frmROMDiff.frx":173E
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Base Address: 0x"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   1575
      End
   End
   Begin VB.FileListBox CurrentFileListA 
      Height          =   4185
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   3495
   End
   Begin VB.DirListBox CurrentDirectoryA 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.DriveListBox CurrentDriveA 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   3615
      Left            =   3720
      TabIndex        =   4
      Top             =   3240
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   16000
      TextRTF         =   $"frmROMDiff.frx":1833
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
   Begin VB.Label Label1 
      Caption         =   "ROMs"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmROMDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public NewCode As String

Private FileByteData(8, 18000) As Byte
Private FileLength(8) As Long
Private ShortestFile As Long

' TODO: just use these....
Private FileName(8) As String
Private FileNameAndPath(8) As String


Public Property Get BaseAddress() As Long

    BaseAddress = 0
    On Error Resume Next
    BaseAddress = CLng("&H" & txtBaseAddr.Text)
    txtBaseAddr.Text = HexStringLong(BaseAddress)

End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GenerateCode(ParentOwnerForm) As String
    NewCode = ""
    
    ' Try to set up the initial conditions...
    On Error Resume Next
    CurrentDriveA.Drive = GetSetting(App.Title, "ROMDiff", "CDRA", "C:\")
    CurrentDirectoryA.Path = GetSetting(App.Title, "ROMDiff", "CDLA", "C:\")
    CurrentFileListA.Path = CurrentDirectoryA.Path
    
    Me.Show vbModal, ParentOwnerForm
    GenerateCode = NewCode

    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: btnCancel_Click
' PURPOSE: Just close without doing anything
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnCancel_Click()
    NewCode = ""
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddTextLine
' PURPOSE: Adds a line to the text file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddTextLine(str As String)
    txtCode.Text = txtCode.Text & str & vbCrLf
End Sub


Private Sub btnCopyToFile_Click()
    NewCode = txtCode.Text
    Unload Me
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: btnGenerate_Click
' PURPOSE: Real work is done here as we try to figure out what the difference
'   is between two binary files.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnGenerate_Click()
On Error GoTo Hell

    If lstFiles.ListCount = 0 Then
        MsgBox "Add at least two ROM binaries to compare", vbOKOnly, "Specify at least 2 files"
        Exit Sub
    End If
    
    If lstFiles.ListCount > 8 Then
        MsgBox "Due to Fluke 9010A limitations, this code will only support up to 8 files.", vbOKOnly, "Specify less than 9 files"
        Exit Sub
    End If
    
    txtCode.Text = ""
    
    Dim fn As Integer
    Dim i As Integer, j As Integer
    ShortestFile = 18000
    Dim ByteData(18000) As Byte
    For i = 0 To lstFiles.ListCount - 1
    
        fn = FreeFile
        Open lstFiles.List(i) For Binary As fn
        FileLength(i) = LOF(fn)
        
        If (FileLength(i) < ShortestFile) Then ShortestFile = FileLength(i)
        
        ' Only get as much as we need -- which is the shortest file
        ' we have to analyze.
        Get fn, 1, ByteData()
        
        ' damn this sucks
        For j = 0 To FileLength(i)
            FileByteData(i, j) = ByteData(j)
        Next
        
        Close 1
    
    Next
    
  
    
    Dim bDiffFound As Boolean
    bDiffFound = False
    Dim bMisMatches As Integer
    bMisMatches = 0
    ' find a difference.
    For i = 0 To ShortestFile - 1
    
        Dim k As Integer
        Dim bByteMatches(8) As Boolean
        
        For j = 0 To lstFiles.ListCount - 1
            bByteMatches(j) = True
        Next
        
        For j = 0 To lstFiles.ListCount - 1
        
            For k = 0 To lstFiles.ListCount - 1
                ' don't compare ourselves!
                If j <> k Then
                    bByteMatches(j) = Not (FileByteData(j, i) <> FileByteData(k, i))
                Else
                    bByteMatches(j) = False
                End If
            Next
            
        Next
        
        ' okay, do we have a mismatch?
        Dim bDontHaveMatch As Boolean
        bDontHaveMatch = False
        For j = 0 To lstFiles.ListCount - 1
            If bByteMatches(j) = True Then bDontHaveMatch = True
        Next
        
        ' I.e. we have a friggin match!
        If Not (bDontHaveMatch) Then
            For j = 0 To lstFiles.ListCount - 1
                AddTextLine vbTab & "!  0x" & HexStringLong(BaseAddress + i) & " = " & HexString(FileByteData(j, i)) & " " & lstFiles.List(j)
            Next
            AddTextLine vbTab & "read @ " & HexStringLong(BaseAddress + i)
            For j = 0 To lstFiles.ListCount - 1
                AddTextLine vbTab & "if DAT = " & HexString(FileByteData(j, i)) & " goto RS" & CStr(j)
            Next
            
            AddTextLine ""
            bMisMatches = bMisMatches + 1
        End If
        
        If bMisMatches > 3 Then Exit For
    
    Next
    
    AddTextLine vbTab & "goto NOMATCH"
    AddTextLine ""
    
    For j = 0 To lstFiles.ListCount - 1
        AddTextLine "!  " & lstFiles.List(j)
        AddTextLine "RS" & CStr(j) & ":"
        AddTextLine vbTab & "dpy ROM SET " & CStr(j)
        AddTextLine vbTab & "aux ROM SET " & CStr(j)
        AddTextLine vbTab & "goto DONE"
        AddTextLine ""
    Next
    
    AddTextLine "NOMATCH:"
    AddTextLine vbTab & "dpy NO ROM MATCH FOUND"
    AddTextLine vbTab & "aux NO ROM MATCH FOUND"
    AddTextLine ""
    
    AddTextLine "DONE:"
    AddTextLine ""
    
    ColorSyntax.CheckRange txtCode, 0, Len(txtCode.Text)
    Exit Sub
    
Hell:
    MsgBox "An error occured. Make sure that this is the unzipped file, and that the files aren't large. " & Err.Description, _
        vbExclamation, "An error occurred"
    

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
' SUB: HexStringLong
' PURPOSE: Returns xxxx in hex form
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function HexStringLong(b As Long) As String

    If b > &HF00 Then
        HexStringLong = Hex(b)
    ElseIf b > &HF0 Then
        HexStringLong = "0" & Hex(b)
    ElseIf b > &HF Then
        HexStringLong = "00" & Hex(b)
    Else
        HexStringLong = "000" & Hex(b)
    End If

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentDirectoryA_Change
' PURPOSE: Update the file list when the directory changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentDirectoryA_Change()
    On Error Resume Next
    CurrentFileListA.Path = CurrentDirectoryA.Path
    SaveSetting App.Title, "ROMDiff", "CDLA", CurrentDirectoryA.Path
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentDriveA_Change
' PURPOSE: Change the directory when the drive changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentDriveA_Change()
    On Error Resume Next
    CurrentDirectoryA.Path = CurrentDriveA.Drive
    SaveSetting App.Title, "ROMDiff", "CDRA", CurrentDriveA.Drive
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentFileListA_DblClick
' PURPOSE: Change the directory when the drive changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentFileListA_DblClick()
    lstFiles.AddItem CurrentFileListA.Path & "\" & CurrentFileListA.FileName
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: lstFiles_DblClick
' PURPOSE: Remove the selected item
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub lstFiles_DblClick()
    lstFiles.RemoveItem (lstFiles.ListIndex)
End Sub
