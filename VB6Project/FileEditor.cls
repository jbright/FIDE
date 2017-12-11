VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Public FileName As String           ' File Name
Public Directory As String          ' File Directory

' The current settings.
Private m_TextContents As String    ' Contents of the file.
Private m_RTFContents As String     ' Contents in RTF format
Private SelStart As Integer
Private SelLength As Integer

' Undo buffer
Private m_LastTextContents As String
Private m_LastRTFContents As String
Private LastSelStart As Integer
Private LastSelLength As Integer


Public FileHasChanged As Boolean    ' Has this file changed?


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' We'll change this to minus the extension later.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get WindowTitle() As String
    Dim strBaseName As String
    strBaseName = FileName
    If (InStr(FileName, ".")) > 0 Then
        strBaseName = Left(strBaseName, InStrRev(strBaseName, ".") - 1)
    End If
    
    If (FileHasChanged) Then
        strBaseName = strBaseName & " *"
    End If
    
    WindowTitle = strBaseName
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Returns the current contents of this file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get TextContents() As String
    TextContents = m_TextContents
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Sets the current contents of this file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Let TextContents(strNewValue As String)
    If (strNewValue <> m_TextContents) Then
        m_TextContents = strNewValue
        FileHasChanged = True
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Returns the current contents of this file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get RTFContents() As String
    RTFContents = m_RTFContents
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Sets the current contents of this file
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Let RTFContents(strNewValue As String)
    m_RTFContents = strNewValue
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Returns the directory and file name
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get FileNameAndDirectory() As String
    If Right(Directory, 1) = "\" Then
        FileNameAndDirectory = Directory & FileName
    Else
        FileNameAndDirectory = Directory & "\" & FileName
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SaveState
' PURPOSE: To save our positional information
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SaveState(rtb As RichTextBox)
    SelStart = rtb.SelStart
    SelLength = rtb.SelLength
    m_TextContents = rtb.Text
    m_RTFContents = rtb.TextRTF
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SelectAll
' PURPOSE: Ctrl+A command
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SelectAll(rtb As RichTextBox)
    rtb.SelStart = 0
    rtb.SelLength = Len(rtb.Text)
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SelectAll
' PURPOSE: Ctrl+A command
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub DoUndo(rtb As RichTextBox)
    
    FIDEMainModule.fMainForm.LockEditorWindow = True
    If (Len(m_LastRTFContents) > 0) Then
        
        With rtb
            ' Lock window update
            LockWindowUpdate .hWnd
            
            ' Save off the current UNDO buffer contents. They become our
            ' "new" undo contents.
            Dim sText As String, sRTF As String, nStart As Integer, nLen As Integer
            
            sText = .Text
            sRTF = .TextRTF
            nStart = .SelStart
            nLen = .SelLength
            
            .TextRTF = m_LastRTFContents
            .SelStart = LastSelStart
            .SelLength = 0 ' LastSelLength (test)
            LockWindowUpdate 0
            
            m_LastTextContents = sText
            m_LastRTFContents = sRTF
            LastSelStart = nStart
            LastSelLength = nLen
            
        End With
    
    End If
    FIDEMainModule.fMainForm.LockEditorWindow = False
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: UpdateTextContents
' PURPOSE: We're being asked to update. Keep track of the last changes so
'   that we can do an "undo" if necessary.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub UpdateTextContents(rtb As RichTextBox)

    If (TextContents <> rtb.Text) Then
        m_LastTextContents = m_TextContents
        m_LastRTFContents = m_RTFContents
        LastSelStart = SelStart
        LastSelLength = SelLength
        
            If False Then
                ' debug code
                OutputLog.AddOutputLine "UNDO Buffer:"
                OutputLog.AddOutputLine vbTab & "Txt: " & m_LastTextContents
                OutputLog.AddOutputLine vbTab & "RTF: " & m_LastRTFContents
                OutputLog.AddOutputLine vbTab & "SelStart: " & LastSelStart
                OutputLog.AddOutputLine vbTab & "SelLength: " & LastSelLength
                OutputLog.AddOutputLine " "
            End If
        
    End If
    
    
    ' Update our text
    TextContents = rtb.Text
    RTFContents = rtb.TextRTF
    ' Save our current location pointers.
    SelStart = rtb.SelStart
    SelLength = rtb.SelLength
        
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: RestoreState
' PURPOSE: To restore our position
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub RestoreState(rtb As RichTextBox)

    FIDEMainModule.fMainForm.LockEditorWindow = True
    With rtb
        LockWindowUpdate .hWnd
        .SetFocus
        .TextRTF = RTFContents
        If SelStart >= 0 Then .SelStart = SelStart
        If SelLength >= 0 Then .SelLength = SelLength
        LockWindowUpdate 0
    End With
    FIDEMainModule.fMainForm.LockEditorWindow = False
    
End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Highlight
' PURPOSE: Do we highlight this file?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function Highlight() As Boolean

    If Not (FIDEMainModule.fMainForm.mnuScriptColors.Checked) Then
        Highlight = False
        Exit Function
    End If
    
    If LCase(Right(Me.FileName, 4)) = ".9lc" Then
        Highlight = True
        Exit Function
    End If

    If LCase(Right(Me.FileName, 2)) = ".s" Then
        Highlight = True
        Exit Function
    End If
    
    Highlight = False
    

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CheckEditedText
' PURPOSE:
'   When a key is pressed in the TX window, this function determines
'   the last and the next word and checks if these are keywords.
'   If so, they are colored accordingly...
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckEditedText()
    If Highlight Then
        ColorSyntax.CheckSubRange FIDEMainModule.fMainForm.txtActiveFile
    End If
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CheckEntireDoc
' PURPOSE:
'   This function checks the entire document for keywords and
'   sets their front color
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckEntireDoc()
    If Highlight Then
        ColorSyntax.CheckRange FIDEMainModule.fMainForm.txtActiveFile, 0, Len(FIDEMainModule.fMainForm.txtActiveFile.Text)
    End If
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CheckRange
' PURPOSE: Just checks on range (if highlighting is enabled)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckRange(nOriginalStart As Integer, nLength As Integer)
    If Highlight Then
        With FIDEMainModule.fMainForm.txtActiveFile
            ColorSyntax.CheckRange FIDEMainModule.fMainForm.txtActiveFile, PrevCR(.Text, nOriginalStart), NextCR(.Text, nOriginalStart + nLength)
        End With
    End If
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Load
' PURPOSE: we're being asked to load up this file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub CreateNew(sFileName As String, sDirectory As String)

    FileName = sFileName
    Directory = sDirectory

    ' Now load the file from disk
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    Set oTextStream = oFileSystem.CreateTextFile(FileNameAndDirectory)
    oTextStream.Write " "
    oTextStream.Close
    ' Special case for empty files
    TextContents = " "
  
    FileHasChanged = False
    frmMain.sbStatusBar.Panels(1).Text = FileNameAndDirectory & " created."
    OutputLog.AddOutputLine "File '" & FileNameAndDirectory & "'...created."

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Load
' PURPOSE: we're being asked to load up this file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub Load(rtb As RichTextBox, sFileName As String, sDirectory As String)

    FileName = sFileName
    Directory = sDirectory

    ' Now load the file from disk
    Dim oFileSystem As New FileSystemObject
    Dim oFile As file
    Set oFile = oFileSystem.GetFile(FileNameAndDirectory)
    Dim oTextStream As TextStream
    Set oTextStream = oFile.OpenAsTextStream(ForReading)
    ' Special case for empty files
    TextContents = ""
    On Error Resume Next
    TextContents = oTextStream.ReadAll
    oTextStream.Close
    On Error GoTo 0
    
    On Error Resume Next
    ' Set the box Text contents. We'll parse it for color a little later.
    rtb.Text = TextContents
    On Error GoTo 0
    
    FileHasChanged = False
    frmMain.sbStatusBar.Panels(1).Text = FileNameAndDirectory & " loaded."
    OutputLog.AddTraceLine "File '" & FileNameAndDirectory & "'...loaded."

End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' TODO: figure out how I want to structure this object
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub Unload()

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: OnSave
' PURPOSE: We're being asked to save our file here.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function OnSave() As Boolean
    On Error GoTo CantSave

    OnSave = False
    ' Do we have unsaved changes?
    If FileHasChanged Then
       
        ' We do have a file name, so save
        Dim oFileSystem As New FileSystemObject
        Dim oTextStream As TextStream
        Set oTextStream = oFileSystem.CreateTextFile(FileNameAndDirectory, True)
         
        oTextStream.Write TextContents
        oTextStream.Close
        
    End If
    
    
    FileHasChanged = False
    frmMain.sbStatusBar.Panels(1).Text = FileNameAndDirectory & " saved."
    OutputLog.AddTraceLine FileNameAndDirectory & " saved."
    
    ' Returning True just means that everything went okay.
    OnSave = True
    
    Exit Function
    
CantSave:
    ' TODO: Big nasty error message here.
    MsgBox "COULD NOT SAVE! " & Err.Description
End Function



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CanClose
' PURPOSE: Can we close?
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CanClose() As Boolean

    ' By default, assume we can't
    CanClose = False
    
    ' If there are no changes, there there is no reason not to close
    If Not (FileHasChanged) Then
        CanClose = True
        Exit Function
    End If
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox(FileNameAndDirectory & " has changed. " & vbCrLf & vbCrLf & _
                "Would you like to save your changes?", vbYesNoCancel, "Save Changes?")
                
    ' Try to save
    If (answer = vbYes) Then
        CanClose = OnSave()
    ElseIf (answer = vbNo) Then
        ' They don't care, we can still close
        CanClose = True
    Else
        ' They are canceling this option, so in effect, we can't close
        CanClose = False        ' (to be explicit)
    End If
    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Class_Initialize
' PURPOSE: Init our values
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Class_Initialize()
    SelStart = -1
    SelLength = -1
End Sub


