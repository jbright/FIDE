Attribute VB_Name = "HelpOutput"
Option Explicit

Public Const Caption = "Help"

' Last topic we accessed
Private LastTopic As String

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Where we do our work
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get BrowseBase() As String
    If (Right(App.Path, 1) = "\") Then
        BrowseBase = App.Path & "Help\"
    Else
        BrowseBase = App.Path & "\Help\"
    End If
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Ouput that we will write to
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get Output() As WebBrowser
    Set Output = FIDEMainModule.fMainForm.wbHelpContext
End Property

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ShowControl
' PURPOSE: Make sure that the correct controls are hidden/shown
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ShowOutput()
    
    ' We don't update our state if the output is minimized
    If FIDEMainModule.fMainForm.OutputMinimized Then
        Exit Sub
    End If
    
    ' If we're hidden...
    ' Don't let them be shown.
    SerialOutputLog.HideOutput
    OutputLog.HideOutput
    
    ' but show us
    Output.Visible = True
    
    ' We have to switch to the correct tab
    If FIDEMainModule.fMainForm.tsOutput.SelectedItem <> HelpOutput.Caption Then
        FIDEMainModule.fMainForm.tsOutput.Tabs(3).Selected = True
        ' Make sure we don't pull the focus off the editor control!
        If FIDEMainModule.fMainForm.txtActiveFile.Visible Then
            FIDEMainModule.fMainForm.txtActiveFile.SetFocus
        End If
    End If
    
    FIDEMainModule.fMainForm.EnableTabButtons


End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: HideOutput
' PURPOSE: The other control may be showing
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub HideOutput()
    If Output.Visible Then
        Output.Visible = False
        FIDEMainModule.fMainForm.EnableTabButtons
    End If
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SetTopic
' PURPOSE: Opens up a topic
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub SetTopic(strTopic As String)

    If LCase(strTopic) <> LCase(LastTopic) Then
        ShowOutput
        FIDEMainModule.fMainForm.wbHelpContext.Navigate BrowseBase & strTopic
        LastTopic = strTopic
    End If
    
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: SetTopic
' PURPOSE: Opens up a topic
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ScanContent(strLine As String)

    ' For efficiency reasons, we'll check this here.
    ' If we're hidden, there is no reason to continue.
    If FIDEMainModule.fMainForm.OutputMinimized Then
        Exit Sub
    End If
    
    Dim strProcess As String
    strProcess = LCase(Trim(Replace(strLine, vbTab, " ")))
    If Len(strProcess) = 0 Then
        Exit Sub
    End If
    
    ' Okay, see if we have anything to show.
    Dim str7 As String
    str7 = Left(strProcess, 7)
    If str7 = "include" Then
        SetTopic "Include.htm"
        Exit Sub
    ElseIf str7 = "execute" Then
        SetTopic "execute.htm"
        Exit Sub
    ElseIf str7 = "exercis" Then
        SetTopic "exercise.htm"
        Exit Sub
    ElseIf str7 = "program" Then
        SetTopic "program.htm"
        Exit Sub
    ElseIf str7 = "declara" Then
        SetTopic "declaration.htm"
        Exit Sub
    End If
    
    Dim str6 As String
    str6 = Left(str7, 6)
    If str6 = "assign" Then
        SetTopic "declaration.htm"
        Exit Sub
    ElseIf str6 = "enable" Then
        SetTopic "enable.htm"
        Exit Sub

    End If
    
    Dim str5 As String
    str5 = Left(str6, 5)
    If str5 = "learn" Then
        SetTopic "learn.htm"
        Exit Sub
    ElseIf str5 = "probe" Then
        SetTopic "probe.htm"
        Exit Sub
    ElseIf str5 = "write" Then
        SetTopic "write.htm"
        Exit Sub

    End If
    
    Dim str4 As String
    str4 = Left(str5, 4)
    
    If str4 = "atog" Then
        SetTopic "atog.htm"
        Exit Sub
    ElseIf str4 = "auto" Then
        SetTopic "auto.htm"
        Exit Sub
    ElseIf str4 = "dtog" Then
        SetTopic "dtog.htm"
        Exit Sub
    ElseIf str4 = "goto" Then
        SetTopic "goto.htm"
        Exit Sub
    ElseIf str4 = "ramp" Then
        SetTopic "ramp.htm"
        Exit Sub
    ElseIf str4 = "read" Then
        SetTopic "read.htm"
        Exit Sub
    ElseIf str4 = "stop" Then
        SetTopic "stop.htm"
        Exit Sub
    ElseIf str4 = "sync" Then
        SetTopic "sync.htm"
        Exit Sub
    ElseIf str4 = "walk" Then
        SetTopic "walk.htm"
        Exit Sub
    ElseIf str4 = "beep" Then
        SetTopic "beep.htm"
        Exit Sub
    ElseIf str4 = "trap" Then
        SetTopic "trap.htm"
        Exit Sub
    End If
    
    Dim str3 As String
    str3 = Left(str4, 3)
    If str3 = "aux" Then
        SetTopic "aux1.htm"
        Exit Sub
    ElseIf str3 = "bus" Then
        SetTopic "bus.htm"
        Exit Sub
    ElseIf str3 = "dpy" Then
        SetTopic "dpy.htm"
        Exit Sub
    ElseIf str3 = "ram" Then
        SetTopic "ram.htm"
        Exit Sub
    ElseIf str3 = "reg" Then
        SetTopic "reg.htm"
        Exit Sub
    ElseIf str3 = "rom" Then
        SetTopic "rom.htm"
        Exit Sub
    ElseIf str3 = "run" Then
        SetTopic "run.htm"
        Exit Sub
    ElseIf str3 = "cpl" Or str3 = "dec" Or str3 = "inc" Or str3 = "shl" Or str3 = "str" Then
        SetTopic "unary.htm"
        Exit Sub
    ElseIf str3 = "pod" Then
        SetTopic "pod.htm"
        Exit Sub
    End If
    
    Dim str2 As String
    str2 = Left(str3, 2)
    If str2 = "if" Then
        SetTopic "if.htm"
        Exit Sub
    ElseIf str2 = "ex" Then
        SetTopic "execute.htm"
        Exit Sub
    ElseIf str2 = "io" Then
        SetTopic "io.htm"
        Exit Sub
    End If
    
    Dim str1 As String
    str1 = Left(str2, 1)
    If str1 = "!" Then
        SetTopic "comment.htm"
        Exit Sub
    End If
    
End Sub


