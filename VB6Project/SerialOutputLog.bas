Attribute VB_Name = "SerialOutputLog"
Option Explicit

Public Const Caption = "Serial Port"


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' Ouput that we will write to
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Property Get Output() As TextBox
    Set Output = FIDEMainModule.fMainForm.txtSerialOutput
End Property


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
    OutputLog.HideOutput
    HelpOutput.HideOutput
    ' but show us
    Output.Visible = True
    
    ' We have to switch to the correct tab
    If FIDEMainModule.fMainForm.tsOutput.SelectedItem <> SerialOutputLog.Caption Then
        FIDEMainModule.fMainForm.tsOutput.Tabs(2).Selected = True
    End If
    
    FIDEMainModule.fMainForm.EnableTabButtons

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
