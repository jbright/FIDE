Attribute VB_Name = "OutputLog"
Option Explicit

Public Const Caption = "Messages"

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Ouput that we will write to
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get Output() As TextBox
    Set Output = FIDEMainModule.fMainForm.txtOutput
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
    HelpOutput.HideOutput
    ' but show us
    Output.Visible = True
    
    
    ' We have to switch to the correct tab
    If FIDEMainModule.fMainForm.tsOutput.SelectedItem <> OutputLog.Caption Then
        FIDEMainModule.fMainForm.tsOutput.Tabs(1).Selected = True
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
' SUB: AddOutputLine
' PURPOSE: Adds new text to the output.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddOutputLine(txt As String)

    ShowOutput
    With Output
        .SelStart = Len(.Text)
        .SelText = txt & vbCrLf
        .SelStart = Len(.Text)
    End With

End Sub

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddTraceLine
' PURPOSE: Tracing lines are used during development and may be removed
'   in production
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddTraceLine(txt As String)
    If frmOptions.PreProcessorDetails Then
        AddOutputLine "> " & txt
    End If
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddOutput
' PURPOSE: Adds new text to the output.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub AddOutput(txt As String)

    ShowOutput
    With Output
        .SelStart = Len(.Text)
        .SelText = txt
        .SelStart = Len(.Text)
    End With
    

End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ClearOutput
' PURPOSE: Adds new text to the output.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ClearOutput()

    ShowOutput
    Output.Text = ""

End Sub

