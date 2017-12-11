VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Children that we may have.
Private m_childrenFiles As Collection

' The name of the file.
Public FileName As String

' The DOS file name. Needed for legacy reasons
Public ShortFileName As String

' Location of this file
Public Directory As String

' Location of where we end up
Public DestinationDirectory As String

' Used for notification purposes
Public ParentFileLineNumber As Integer
Public ParentSourceFile As String


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Gets the name of the file, including directory
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get DirectoryAndFileName() As String
    DirectoryAndFileName = MakePathName(Directory, FileName)
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Where we end up in life
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get DestinationDirectoryAndFileName() As String
    DestinationDirectoryAndFileName = MakePathName(DestinationDirectory, FileName)
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' returns a collection of active files (file editor modules)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get Children() As Collection

    If (m_childrenFiles Is Nothing) Then
        Set m_childrenFiles = New Collection
    End If
    Set Children = m_childrenFiles
    
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Gets the name of the compiled file. Should be the short DOS file name
' with the .h extention
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get CompiledFileName() As String
    CompiledFileName = Left(ShortFileName, InStrRev(ShortFileName, ".") - 1) & ".h"
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CopyToDestination
' PURPOSE: Copy this file, and all child files, to the destination. Also,
'   keep track of the new, short file name
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CopyToDestination(strDestinationDirectory As String) As Boolean

    CopyToDestination = False
    
    ' First, take care of myself.
    Dim oFileSystem As New FileSystemObject
    Dim strNewFile As String
    DestinationDirectory = strDestinationDirectory
    strNewFile = DestinationDirectoryAndFileName
    
    ' Check to see if the new file exists in destination. Just give a warning.
    If oFileSystem.FileExists(strNewFile) Then
        OutputLog.AddOutputLine "WARNING: Include file already exists. File may have been included"
        OutputLog.AddOutputLine "multiple times, or two files with the exact same file name may exist."
        OutputLog.AddOutputLine "It is recommended that you give your files distictive file names."
        OutputLog.AddOutputLine DirectoryAndFileName & " -> " & strNewFile & " (exists)."
    End If
    
    On Error GoTo CantCopyFile
    oFileSystem.CopyFile DirectoryAndFileName, strNewFile, True
    On Error GoTo 0
    
    On Error GoTo CantGetFile
    Dim oFile As File
    Set oFile = oFileSystem.GetFile(strNewFile)
    On Error GoTo 0
    
    ShortFileName = oFile.ShortName
    
    ' Cleanup
    Set oFile = Nothing
    Set oFileSystem = Nothing
    
    ' Now, process any children we may have (recursively)
    Dim i As Integer
    Dim objSource As SourceFile
    For i = 1 To Children.Count
        Set objSource = Children.Item(i)
        If Not (objSource.CopyToDestination(strDestinationDirectory)) Then
            ' Exit processing if any of these copy procedures fail
            Exit Function
        End If
    Next
    
    CopyToDestination = True
    OutputLog.AddTraceLine "File copied " & DirectoryAndFileName & " -> " & strNewFile & " (" & ShortFileName & ")"
    Exit Function
    

CantCopyFile:
    OutputLog.AddOutputLine "ERROR: Copying of " & DirectoryAndFileName & " to " & strNewFile & " failed."
    OutputLog.AddOutputLine Err.Description
    Exit Function

CantGetFile:
    OutputLog.AddOutputLine "ERROR: Could not file copied over:  " & strNewFile & "."
    OutputLog.AddOutputLine Err.Description
    Exit Function

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: PreProcess
' PURPOSE: Scan the file to see if there are any included files that must
'   be dealt with. They come in three varieties:
'
'       (1) include "foo.pod"
'               built-in pod file. can be ignored.
'       (2) include "goo.9lc"
'               must exist in our directory. must be copied as well
'       (3) include "~/lib/soo.9lc"
'               "~" means replace it with the project location.
'
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function PreProcess()
On Error GoTo PreProcessError

    PreProcess = False
    
    ' Scan this file.
    Dim oStream As TextStream
    Dim strCurrentLine
    Dim oFileSystem As New FileSystemObject
    
    ' Quick check to make sure we're found
    If Not (oFileSystem.FileExists(DirectoryAndFileName)) Then
        Err.Raise 1, "SoureFile", "File " & DirectoryAndFileName & " not found. Check include statement or file location." & vbCrLf & _
            "SOURCE: " & ParentSourceFile & ", line " & ParentFileLineNumber
    End If
    
    Dim nLineNumber As Integer
    Dim SourceCode As SourceFile
    Set oStream = oFileSystem.OpenTextFile(DirectoryAndFileName, ForReading, False)
    nLineNumber = 0
    While Not oStream.AtEndOfStream
        strCurrentLine = oStream.ReadLine
        nLineNumber = nLineNumber + 1
        ' remove any white space from beginning (and end)
        strCurrentLine = Trim(strCurrentLine)
        ' convert it to lower case--just easier (and I'm lazy)
        strCurrentLine = LCase(strCurrentLine)
        If Left(strCurrentLine, 7) = "include" Then
            ' We have a line that appears to have an include statement. We must peak at it.
            strCurrentLine = Right(strCurrentLine, Len(strCurrentLine) - 7)
            ' We should now have:
            ' "foo.ext"
            strCurrentLine = Trim(strCurrentLine)
            ' Split on quote mark.
            Dim arr
            arr = Split(strCurrentLine, """", 3)
            ' now arr(0), should be something. like a file.
            Dim strFileToInclude As String
            
            ' Missing quote maybe?
            If (UBound(arr)) < 2 Then
                Err.Raise 2, "SourceFile", "Include statement appears to be invalid in file " & FileName
            End If
            
            strFileToInclude = Trim(arr(LBound(arr) + 1))
            ' If we end in .pod, then we can ignore it. It's built in.
            If Right(strFileToInclude, 4) <> ".pod" Then
                
                strFileToInclude = Replace(strFileToInclude, "/", "\")
                
                ' Okay, we need to add this object to our child list.
                Set SourceCode = New SourceFile
                SourceCode.FileName = ParseForFileName(strFileToInclude)
                SourceCode.Directory = ParseForDirectoryName(strFileToInclude)
                SourceCode.ParentFileLineNumber = nLineNumber
                SourceCode.ParentSourceFile = DirectoryAndFileName
                Children.Add SourceCode
                
            End If
            
        End If
    Wend
    oStream.Close

    PreProcess = True
    OutputLog.AddTraceLine "Processed " & DirectoryAndFileName
    
    ' Now, process any children we may have (recursively)
    Dim i As Integer
    Dim objSource As SourceFile
    For i = 1 To Children.Count
        Set objSource = Children.Item(i)
        objSource.PreProcess
    Next
    
    Exit Function
    
PreProcessError:
    OutputLog.AddOutputLine "ERROR: Preprocessing of " & DirectoryAndFileName & " failed."
    OutputLog.AddOutputLine Err.Description
    
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: UpdateIncludes
' PURPOSE: See PreProcess for more details. Now that the file is in its
'   destination, we're going to scan it and do some replacements of include
'   file statements. There is some duplicated code, but this part is so
'   messy anyways, I don't really mind.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function UpdateIncludes() As Boolean
On Error GoTo UpdateIncludesError

    UpdateIncludes = False
    
    ' Scan this file.
    Dim oStream As TextStream
    Dim strCurrentLine
    Dim oFileSystem As New FileSystemObject
    
    ' Quick check to make sure we're found
    If Not (oFileSystem.FileExists(DestinationDirectoryAndFileName)) Then
        Err.Raise 3, "SoureFile", "File " & DestinationDirectoryAndFileName & " not found. Unrecoverable error."
    End If
    
    ' First, get the entire file.
    Dim strFileContents As String
    Set oStream = oFileSystem.OpenTextFile(DestinationDirectoryAndFileName, ForReading, False)
    strFileContents = oStream.ReadAll
    oStream.Close
    
    ' Now we have the file, and the stream is closed. Now we're going to open it up
    ' for reading, and re-write the contents.
    Dim arrLines
    arrLines = Split(strFileContents, vbCrLf)
    Dim nLine As Integer
    Dim nLineNumber As Integer
    Dim objIncludeSourceFile As SourceFile
    
    Set oStream = oFileSystem.OpenTextFile(DestinationDirectoryAndFileName, ForWriting, True)
    
    ' Go through each line
    nLineNumber = 0
    For nLine = LBound(arrLines) To UBound(arrLines) Step 1
    
        nLineNumber = nLineNumber + 1
        Set objIncludeSourceFile = GetIncludeLine(nLineNumber)
        If Not (objIncludeSourceFile Is Nothing) Then
            ' This is the include line we've been waiting for...
            oStream.WriteLine "include """ & objIncludeSourceFile.ShortFileName & """"
        Else
            ' Output the normal line.
            oStream.WriteLine arrLines(nLine)
        End If
    
    Next
    ' Done with this file.
    oStream.Close
    
    
    UpdateIncludes = True
    OutputLog.AddTraceLine "Include file updated for " & DestinationDirectoryAndFileName
    
    ' Now, process any children we may have (recursively)
    Dim i As Integer
    Dim objSource As SourceFile
    For i = 1 To Children.Count
        Set objSource = Children.Item(i)
        objSource.UpdateIncludes
    Next
    
    Exit Function
    
UpdateIncludesError:
    OutputLog.AddOutputLine "ERROR: Processing include of " & DestinationDirectoryAndFileName & " failed."
    OutputLog.AddOutputLine Err.Description

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: GetIncludeLine
' PURPOSE: Checks to see if we think we have an include line here.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function GetIncludeLine(nLineNumber As Integer) As SourceFile
    Set GetIncludeLine = Nothing
    
    Dim i As Integer
    Dim objSource As SourceFile
    For i = 1 To Children.Count
        Set objSource = Children.Item(i)
        If nLineNumber = objSource.ParentFileLineNumber Then
            Set GetIncludeLine = objSource
            Exit Function
        End If
    Next
    Exit Function
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ParseForFileName
' PURPOSE: Just want the file name part (including extenstion). Assumes that
'   the strLine passed in is not a .pod file, and that this line includes
'   everything in quotes.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function ParseForFileName(strLine As String) As String

    ' If the slash is found then we need to count backwards from that location.
    If InStr(strLine, "\") Then
        Dim nLoc As Integer
        nLoc = InStrRev(strLine, "\")
        ParseForFileName = Right(strLine, Len(strLine) - nLoc)
    Else
        ParseForFileName = strLine
    End If

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ParseForDirectoryName
' PURPOSE: Figure out the directory piece for this.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function ParseForDirectoryName(strLine As String) As String

    ' If the slash is found then we need to count backwards from that location.
    If InStr(strLine, "\") Then
        Dim subFolder As String
        Dim nLoc As Integer
        nLoc = InStrRev(strLine, "\")
        subFolder = Left(strLine, nLoc)
        
        Dim strBase As String
        If Left(strLine, 1) = "~" Then
            ' Base is the project's directory
            strBase = frmRootLocation.ProjectPath
            subFolder = Right(subFolder, Len(subFolder) - 2)
        Else
            ' Base is our current directory
            strBase = Directory
        End If
        
        ' Assemble everything now.
        ParseForDirectoryName = MakePathName(strBase, subFolder)
        
    Else
        ' In this case, it's assumed to be in the same directory
        ParseForDirectoryName = Directory
    End If

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' FUNCTION: MakePathName
' PURPOSE: Makes sure that we build a valid name
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MakePathName(strDirectory As String, strFile As String) As String

    If (Right(strDirectory, 1) = "\") Then
        MakePathName = strDirectory & strFile
    Else
        MakePathName = strDirectory & "\" & strFile
    End If
    
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' I forget how smart VB is, so we'll be nice and destroy everything.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub Class_Terminate()
    Dim i As Integer
    Dim objSource As SourceFile
    For i = 1 To Children.Count
        Set objSource = Children.Item(i)
        Set objSource = Nothing
    Next
    Set m_childrenFiles = Nothing
End Sub
