Attribute VB_Name = "Compiler"
Option Explicit

Public Const CompilerOutputFile = "9lc.out"

' Source file to be compiled
Private SourceCode As SourceFile


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Where we do our work
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get InternalDirectory() As String
    If (Right(App.Path, 1) = "\") Then
        InternalDirectory = App.Path & "Internal"
    Else
        InternalDirectory = App.Path & "\Internal"
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Name of the batch file we'll end up executing (includes path)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get BatchFile() As String
    BatchFile = SourceCode.MakePathName(InternalDirectory, "9lc.bat")
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Returns the name of the compiled file. (if it exists)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get CompiledDirectoryAndFileName() As String
    CompiledDirectoryAndFileName = ""
    If Not (SourceCode Is Nothing) Then
        CompiledDirectoryAndFileName = SourceCode.MakePathName(InternalDirectory, SourceCode.CompiledFileName)
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Compile
' PURPOSE: To actually compile a given file.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Compile() As Boolean

    Compile = False

    ' Clean up
    Set SourceCode = Nothing
    ' Start over
    Set SourceCode = New SourceFile
    ' Get internal directory
    Dim strInternalDir As String
    strInternalDir = InternalDirectory
    
    With FIDEMainModule.fMainForm.ActiveFileEditor
        SourceCode.FileName = .FileName
        SourceCode.Directory = .Directory
    End With
    
    ' Start with a blank slate
    OutputLog.ClearOutput
    
    OutputLog.AddOutputLine "Starting process..."
    If SourceCode.PreProcess() Then
        OutputLog.AddTraceLine "Internal pre-process complete."
    Else
        OutputLog.AddOutputLine "Compile failed."
        Exit Function
    End If
    
    CleanUpDirectory strInternalDir
    
    ' copy these files over
    If SourceCode.CopyToDestination(strInternalDir) Then
        OutputLog.AddTraceLine "Copy completed."
    Else
        OutputLog.AddOutputLine "Copy failed."
        Exit Function
    End If
    
    ' source files have to be processed to deal with long file names in the include names.
    ' we're going to take care of that now.
    If SourceCode.UpdateIncludes() Then
        OutputLog.AddTraceLine "Includes updated."
    Else
        OutputLog.AddOutputLine "Include processing failed."
        Exit Function
    End If
    
    ' Now we have to build a .bat file that will run our command.
    If CreateBatchFile() Then
        OutputLog.AddTraceLine "Batch file created."
    Else
        OutputLog.AddOutputLine "Batch file processing failed."
        Exit Function
    End If
    
    ' Now this is where the magic begins. Lauch the compile
    ' Shell out to legacy app
    Dim oScript As New WshShell
    
    OutputLog.AddTraceLine "Preparing to execute batch file."
    ' Run the file.
    oScript.Run """" & BatchFile & """", 0, True
    
    ' Display results to user.
    LogResults
    
    Compile = True
    OutputLog.AddOutputLine "Processing completed."
    Exit Function
    
'
'
'    ' Hack for Joe's machine.
'SecondAttempt:
'    OutputLog.AddTraceLine "Using late binding to launch batch file."
'    Dim oScript2
'    Set oScript2 = CreateObject("Wscript.Shell")
'    OutputLog.AddTraceLine "Bound to shell."
'    oScript2.Run """" & BatchFile & """", 0, True
'    LogResults
'    Compile = True
'    OutputLog.AddOutputLine "Processing completed."
    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' FUNCTION: LogResults
' PURPOSE: Tells the user what has happened
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function LogResults() As Boolean
    
    LogResults = False
    Dim oFileSystem As New FileSystemObject
    Dim oTextStream As TextStream
    
    On Error GoTo FileNotFound
    Dim strOutputFile As String
    strOutputFile = SourceCode.MakePathName(SourceCode.MakePathName(App.Path, "Internal"), CompilerOutputFile)
    Set oTextStream = oFileSystem.OpenTextFile(strOutputFile, ForReading, False)
    
    While Not oTextStream.AtEndOfStream
        OutputLog.AddOutputLine oTextStream.ReadLine
    Wend
    
    oTextStream.Close
    LogResults = True
    Exit Function
    
FileNotFound:
    OutputLog.AddOutputLine "ERROR Could not find output compiler results. "
    OutputLog.AddOutputLine "Check for file: " & strOutputFile & "."

End Function



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' FUNCTION: CreateBatchFile
' PURPOSE: Builds a temporary batch file to use to execute the command.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function CreateBatchFile() As Boolean
On Error GoTo CouldNotCreateBatch

    CreateBatchFile = False
    Dim oFileSystem As New FileSystemObject

    Dim oTextStream As TextStream
    Set oTextStream = oFileSystem.CreateTextFile(BatchFile, True)
    oTextStream.WriteLine "@Echo Off"
    oTextStream.WriteLine "cd """ & SourceCode.MakePathName(App.Path, "Internal") & """"
    oTextStream.WriteLine "9LC.exe " & SourceCode.ShortFileName & " > " & CompilerOutputFile
    oTextStream.Close

    CreateBatchFile = True
    Exit Function
    
CouldNotCreateBatch:
    OutputLog.AddOutputLine "ERROR: Could not create batch file " & BatchFile & "."
    OutputLog.AddOutputLine Err.Description
    Exit Function
    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CleanUpDirectory
' PURPOSE: Get rid of any files that might have been left over.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub CleanUpDirectory(strInternalDir As String)

    Dim oFileSystem As New FileSystemObject
    ' These are file names that we may have created.
    If False Then
        ' NOt sure why... but it was bombing out here.
        oFileSystem.DeleteFile (SourceCode.MakePathName(strInternalDir, "*.s"))
        oFileSystem.DeleteFile SourceCode.MakePathName(strInternalDir, "*.h")
        oFileSystem.DeleteFile SourceCode.MakePathName(strInternalDir, "*.9lc")
        oFileSystem.DeleteFile SourceCode.MakePathName(strInternalDir, "*.txt")
    End If
    Set oFileSystem = Nothing

End Sub
