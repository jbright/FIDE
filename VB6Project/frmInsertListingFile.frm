VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInsertListingFile 
   Caption         =   "Insert Assembly Listing File"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   Icon            =   "frmInsertListingFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtGeneratedCode 
      Height          =   2775
      Left            =   3720
      TabIndex        =   12
      Top             =   3840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4895
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   1.40000e5
      TextRTF         =   $"frmInsertListingFile.frx":173A
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
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9720
      TabIndex        =   11
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton btnApply 
      Caption         =   "Insert Listing Code into Program"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   6720
      Width           =   5895
   End
   Begin VB.CommandButton btnRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   9360
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtEndPosition 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   7680
      TabIndex        =   7
      Text            =   "17"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtStartPosition 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Text            =   "5"
      Top             =   120
      Width           =   495
   End
   Begin RichTextLib.RichTextBox txtListingFile 
      Height          =   2655
      Left            =   3720
      TabIndex        =   3
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4683
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   1.40000e5
      TextRTF         =   $"frmInsertListingFile.frx":17BA
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
   Begin VB.FileListBox CurrentFileList 
      Height          =   3405
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   3600
      Width           =   3495
   End
   Begin VB.DirListBox CurrentDirectory 
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.DriveListBox CurrentDrive 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Preview of code generated:"
      Height          =   255
      Left            =   3720
      TabIndex        =   14
      Top             =   3600
      Width           =   7095
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Preview of listing source:"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   600
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "characters."
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "characters and all characters after"
      Height          =   495
      Left            =   5160
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "Ignore first "
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmInsertListingFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

' The name of the current file
Private CurrentFileName As String
Private bFirstLineOutputted As Boolean
Private NewCode As String

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' IgnoreFirstChars
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get IgnoreFirstChars() As Integer

    ' default
    IgnoreFirstChars = 4
    On Error Resume Next
    IgnoreFirstChars = CInt(txtStartPosition.Text)

End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' IgnoreAfterChars
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get IgnoreAfterChars() As Integer

    ' default
    IgnoreAfterChars = 17
    On Error Resume Next
    IgnoreAfterChars = CInt(txtEndPosition.Text)

End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: btnApply_Click
' PURPOSE: We're going to use these changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnApply_Click()
    NewCode = txtGeneratedCode.Text
    Unload Me
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
' SUB: CurrentDrive_Change
' PURPOSE: Change the directory when the drive changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentDrive_Change()
    On Error Resume Next
    CurrentDirectory.Path = CurrentDrive.Drive
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
' SUB: btnApply_Click
' PURPOSE: Re-applies our changes
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnRefresh_Click()
    ApplyFile
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CurrentFileList_Click
' PURPOSE: We've selected a file. Show it.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CurrentFileList_Click()

    btnRefresh.Enabled = True
    btnApply.Enabled = True
    CurrentFileName = CurrentFileList.Path & "\" & CurrentFileList.FileName
    ApplyFile

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Form_Load
' PURPOSE: Set our initial conditions
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Form_Load()
    btnRefresh.Enabled = False
    btnApply.Enabled = False
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ApplyFile
' PURPOSE: Applies these settings to the current file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub ApplyFile()

    txtListingFile.Text = ""
    
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Set oTextStream = oFileSystem.OpenTextFile(CurrentFileName, ForReading)
    
    ' For the output code window
    bFirstLineOutputted = False
    
    With txtListingFile
        
        ' Lock window update
        LockWindowUpdate .hWnd
        
        Dim strLine As String
        Dim nCurSelStart As Integer
        While Not oTextStream.AtEndOfStream
            strLine = oTextStream.ReadLine
            nCurSelStart = .SelStart
            .SelText = strLine & vbCrLf

            ' Set the whole line to be black
            .SelStart = nCurSelStart
            .SelLength = nCurSelStart + Len(strLine)
            .SelColor = RGB(0, 0, 0)
            
            If Len(strLine) > IgnoreFirstChars Then
                ' Set the color of the first part of the line.
                .SelStart = nCurSelStart
                .SelLength = IgnoreFirstChars
                .SelColor = RGB(192, 192, 192)
            End If
            
            If Len(strLine) > IgnoreAfterChars Then
                ' Set the color of the second part of the line
                .SelStart = nCurSelStart + IgnoreAfterChars
                .SelLength = Len(strLine) - IgnoreAfterChars
                .SelColor = RGB(192, 192, 192)
            End If

            .SelStart = Len(.Text)
                
            AddCodeLine strLine
        Wend
        
        .SelStart = 0
        txtGeneratedCode.SelStart = 0
        
        ' Unlock window update
        LockWindowUpdate 0
    
    End With

    ColorSyntax.CheckRange txtGeneratedCode, 0, Len(txtGeneratedCode.Text)

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddCodeLine
' PURPOSE: Adds a code line based on our settings.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddCodeLine(strNextLine As String)

    If Len(Trim(strNextLine)) = 0 Then
        Exit Sub
    End If
    
    Dim strLine As String
    If Len(strNextLine) > IgnoreFirstChars Then
        strLine = Mid(strNextLine, IgnoreFirstChars + 1, IgnoreAfterChars - IgnoreFirstChars)
        strLine = Replace(strLine, "  ", "|")
        
        Dim arrLines
        arrLines = Split(Replace(strLine, "  ", "|"), "|")
        
        Dim arr
        Dim strWord
        Dim bFirstWord As Boolean
        Dim strCommentPart As String
        arr = Split(Trim(arrLines(LBound(arrLines))), " ")
        bFirstWord = True
        For Each strWord In arr
        
            ' Comment each line we're outputted
            If bFirstWord Then
                strCommentPart = vbTab & "! " & strNextLine
            Else
                strCommentPart = ""
            End If
        
            If Not (bFirstLineOutputted) Then
                bFirstLineOutputted = True
                ' first line has special code
                txtGeneratedCode.Text = ""
                txtGeneratedCode.SelText = vbTab & "REG9 = 0000 " & vbTab & vbTab & vbTab & "! Set to your base address" & vbCrLf
                txtGeneratedCode.SelText = vbTab & "write @ REGF = " & strWord & vbTab & strCommentPart & vbCrLf
            Else
                txtGeneratedCode.SelText = vbTab & "write @ REGF inc = " & strWord & strCommentPart & vbCrLf
            End If
            
            ' No longer the first word
            bFirstWord = False
        Next
        
    End If

End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Show this and return a value
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GenerateCode(ParentOwnerForm) As String
    NewCode = ""
    Me.Show vbModal, ParentOwnerForm
    GenerateCode = NewCode
End Function


